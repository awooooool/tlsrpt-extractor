import Imap, { ImapMessageAttributes } from "imap";
import dotenv from "dotenv";
import { Base64Decode } from "base64-stream";
import fs from "fs";
import zlib from "node:zlib";

dotenv.config();

const dir: string = process.env.REPORTS_DIR || "./reports";
const readonly: boolean = process.env.NODE_ENV === "development" ? true : false;

interface IReportClass {
  writeReports(): void;
}

interface IPolicyDetails {
  "policy-type": "sts" | "tlsa" | "no-policy-found";
  "policy-string": string[];
  "policy-domain": string;
  "mx-host"?: string;
}

interface IPolicySummary {
  "total-successful-session-count": number;
  "total-failure-session-count"?: number;
}

interface IPolicyFailureDetail {
  "failed-session-count": number;
  "receiving-mx-hostname": string;
  "result-type": string;
  "receiving-ip": string;
}

interface IPolicy {
  policy: IPolicyDetails;
  summary: IPolicySummary;
  "failure-details"?: IPolicyFailureDetail[];
}

interface IReportMetadata {
  "organization-name": string;
  "date-range": {
    "start-datetime": Date;
    "end-datetime": Date;
  };
  "contact-info": string;
  "report-id": string;
}

interface IReport extends IReportMetadata {
  policies: IPolicy[];
}

interface IProcessedPolicy
  extends Omit<IPolicy, "summary" | "failure-details"> {
  summary: Omit<IPolicySummary, "total-failure-session-count">;
  "failure-details"?: IPolicyFailureDetail;
}

interface IProcessedReport extends IReportMetadata {
  policies: IProcessedPolicy;
}

class Report implements IReportClass {
  // report metadata {organization-name, date-range, contact-info, report-id}
  private metadata: IReportMetadata;

  // policies contained in report
  private policies: IPolicy[];

  // filename in attachment
  private filename: string;

  // processed flattened report ready for logstash
  private processed: IProcessedReport[] = [];

  constructor(report: IReport, filename: string) {
    // report metadata
    const { policies, ...reportMetadata } = report;
    this.metadata = reportMetadata;

    // report policies
    this.policies = policies;

    // report filename
    this.filename = filename;
  }

  // to process report & flatten policies and failure-details
  private processReport(): void {
    this.policies.forEach((policy) => {
      delete policy.summary["total-failure-session-count"];
      if (policy["failure-details"]) {
        policy["failure-details"].forEach((detail) => {
          const processedReport = {
            policies: {
              policy: policy.policy,
              summary: policy.summary,
              "failure-details": detail,
            },
            ...this.metadata,
          };
          this.processed.push(processedReport);
        });
      } else {
        const processedReport = {
          policies: {
            policy: policy.policy,
            summary: policy.summary,
          },
          ...this.metadata,
        };
        this.processed.push(processedReport);
      }
    });
  }

  // to write processed reports in this.processed
  public writeReports(): void {
    this.processReport();
    const uid = parseInt(process.env.UID || "0", 10);
    const gid = parseInt(process.env.GID || "0", 10);
    for (const [index, report] of this.processed.entries()) {
      const filename = this.filename.replace(/(\.[\w\d_-]+)$/i, `-${index}$1`);
      fs.writeFile(`${dir}/${filename}`, JSON.stringify(report), () => {
        console.log(`Done writing ${filename}`);
        fs.chown(`${dir}/${filename}`, uid, gid, () => {
          console.log(`Done owning ${filename}`);
        });
      });
    }
  }
}

// array for report objects
const reports: Report[] = [];

// check if numeric env is actually numeric
const getEnv = () => {
  let valid = true;
  let port = null;
  let regexNumOnly = new RegExp(/^\d*$/); // for numeric check
  // check if env has value
  if (
    !(
      process.env.IMAP_HOST &&
      process.env.IMAP_USER &&
      process.env.IMAP_PASS &&
      process.env.UID &&
      process.env.GID
    )
  ) {
    valid = false;
  }
  for (const value of [process.env.UID, process.env.GID]) {
    if (!value || !regexNumOnly.test(value)) {
      // check if env is numeric
      valid = false; // set valid to false if env is invalid
      break;
    }
  }
  if (process.env.IMAP_PORT && regexNumOnly.test(process.env.IMAP_PORT)) {
    port = parseInt(process.env.IMAP_PORT, 10); // convert to numeric, process.env always output string or undefined
  }
  if (!valid) {
    console.log(".env isn't properly set up");
    process.exit(1); // .env check, exit if haven't set up properly
  }
  return {
    host: process.env.IMAP_HOST as string,
    user: process.env.IMAP_USER as string,
    pass: process.env.IMAP_PASS as string,
    port: port || 993,
    uid: parseInt(process.env.UID as any, 10),
    gid: parseInt(process.env.GID as any, 10),
  };
};

// create folder if not exist
if (!fs.existsSync(dir)) {
  fs.mkdirSync(dir);
}

const imap = new Imap({
  user: getEnv().user,
  password: getEnv().pass,
  host: getEnv().host,
  port: getEnv().port,
  tls: true,
});

// open IMAP inbox
function openInbox(cb: (error: Error) => void) {
  imap.openBox("INBOX", readonly, cb);
}

// for finding attachment(s) contained in each emails
function findAttachment(attrs: ImapMessageAttributes): Array<any> {
  const attachments: Array<any> = [];
  attrs.struct?.forEach((element) => {
    if (!Array.isArray(element)) return;
    element.forEach((part) => {
      if (
        part.disposition &&
        ["INLINE", "ATTACHMENT"].indexOf(part.disposition.type.toUpperCase())
      ) {
        attachments.push(part);
      }
    });
  });

  return attachments;
}

// search for unread emails
function searchIMAP(errorSearch: Error, results: number[]) {
  if (errorSearch) throw errorSearch;
  const f = imap.fetch(results, {
    bodies: "HEADER.FIELDS (FROM TO SUBJECT DATE)",
    struct: true,
  });
  f.on("message", function (msg, seqno) {
    console.log("Message #%d", seqno);
    const prefix = `(#${seqno}) `;
    msg.once("attributes", function (attrs: ImapMessageAttributes) {
      const attachments = findAttachment(attrs);

      attachments.forEach((attachment) => {
        const fetch = imap.fetch(attrs.uid, {
          bodies: attachment.partID,
          struct: true,
        });
        fetch.on("message", function (msgFetch) {
          msgFetch.on("body", async function (stream) {
            const filename =
              attachment.disposition.params.filename.match(
                /(.+?)(\.[^.]*$|$)/
              )[1];

            if (attachment.encoding.toLowerCase() === "base64")
              stream = stream.pipe(new Base64Decode());
            if (attachment.subtype.search(/gzip/) !== -1)
              stream = stream.pipe(zlib.createGunzip());

            let tlsreport = "";
            stream
              .on("data", (d) => {
                tlsreport += d.toString();
              })
              .on("end", () => {
                const report = new Report(JSON.parse(tlsreport), filename);
                reports.push(report);
              });
          });
        });
      });
    });
    msg.once("end", function () {
      console.log(`${prefix}Finished`);
    });
  });
  f.once("error", function (fetchError) {
    console.log(`Fetch error: ${fetchError}`);
  });
  f.once("end", function () {
    console.log("Done fetching all messages!");
    imap.end();
  });
}

// main function, will run once imap.connect() successful
imap.once("ready", function () {
  openInbox(function (err: Error) {
    if (err) throw err;
    imap.search(["UNSEEN"], searchIMAP);
  });
});

// if things happen. SPOILER ALERT: it will.
imap.once("error", function (error: any) {
  console.log(error);
});

// mark as the end of connection
imap.once("end", function () {
  console.log("Connection ended");
});

imap.once("close", function () {
  // write file after connection is closed
  for (const report of reports) {
    report.writeReports();
  }
});

// attemps to connect
imap.connect();
