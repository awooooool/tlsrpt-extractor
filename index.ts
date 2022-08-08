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
  private report: IReport;

  private metadata: IReportMetadata;

  private policies: IPolicy[];

  private filename: string;

  private processed: IProcessedReport[] = [];

  constructor(report: IReport, filename: string) {
    // original report
    this.report = report;

    // report metadata
    const { policies, ...reportMetadata } = this.report;
    this.metadata = reportMetadata;

    // report policies
    this.policies = policies;

    // report filename
    this.filename = filename;
  }

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

  public writeReports(): void {
    this.processReport();
    this.processed.forEach((report, index) => {
      fs.writeFileSync(
        `${dir}/${this.filename.replace(/(\.[\w\d_-]+)$/i, `-${index}$1`)}`,
        JSON.stringify(report)
      );
    });
  }
}

if (
  !(process.env.IMAP_HOST && process.env.IMAP_USER && process.env.IMAP_PASS)
) {
  console.log(".env isn't properly set up");
  process.exit(1);
}

if (!fs.existsSync(dir)) {
  fs.mkdirSync(dir);
}

const imap = new Imap({
  user: process.env.IMAP_USER,
  password: process.env.IMAP_PASS,
  host: process.env.IMAP_HOST,
  port: (process.env.IMAP_PORT as any) || 993,
  tls: true,
});

function openInbox(cb: (error: Error) => void) {
  imap.openBox("INBOX", readonly, cb);
}

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
                report.writeReports();
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

imap.once("ready", function () {
  openInbox(function (err: Error) {
    if (err) throw err;
    imap.search(["UNSEEN"], searchIMAP);
  });
});

imap.once("error", function (error: any) {
  console.log(error);
});

imap.once("end", function () {
  console.log("Connection ended");
});

imap.connect();
