import Imap, { ImapMessageAttributes } from "imap";
import dotenv from "dotenv";
import { Base64Decode } from "base64-stream";
import fs from "fs";
import zlib from "node:zlib";

dotenv.config();

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

class Report {
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

    protected processReport(): void {
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
                `./reports/${this.filename.replace(
                    /(\.[\w\d_-]+)$/i,
                    `-${index}$1`
                )}`,
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

const dir = process.env.REPORTS_DIR || "./reports";

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

function openInbox(cb: (error: Error, mailbox: Imap.Box) => void) {
    imap.openBox("INBOX", true, cb);
}

// function toUpper(thing) {
//     return thing && thing.toUpperCase ? thing.toUpperCase() : thing;
// }

// function findAttachmentParts(struct, attachments?) {
//     attachments = attachments || [];
//     for (var i = 0, len = struct.length, r; i < len; ++i) {
//         if (Array.isArray(struct[i])) {
//             findAttachmentParts(struct[i], attachments);
//         } else {
//             if (
//                 struct[i].disposition &&
//                 ["INLINE", "ATTACHMENT"].indexOf(
//                     toUpper(struct[i].disposition.type)
//                 ) > -1
//             ) {
//                 attachments.push(struct[i]);
//             }
//         }
//     }
//     return attachments;
// }

// function buildAttMessageFunction(attachment) {
//     var filename = attachment.params.name;
//     var encoding = attachment.encoding;

//     return function (msg, seqno) {
//         var prefix = "(#" + seqno + ") ";
//         msg.on("body", function (stream, info) {
//             //Create a write stream so that we can stream the attachment to file;
//             console.log(
//                 prefix + "Streaming this attachment to file",
//                 filename,
//                 info
//             );
//             var writeStream = fs.createWriteStream(filename);
//             writeStream.on("finish", function () {
//                 console.log(prefix + "Done writing to file %s", filename);
//             });

//             //stream.pipe(writeStream); this would write base64 data to the file.
//             //so we decode during streaming using
//             if (toUpper(encoding) === "BASE64") {
//                 //the stream is base64 encoded, so here the stream is decode on the fly and piped to the write stream (file)
//                 stream.pipe(base64.decode()).pipe(writeStream);
//             } else {
//                 //here we have none or some other decoding streamed directly to the file which renders it useless probably
//                 stream.pipe(writeStream);
//             }
//         });
//         msg.once("end", function () {
//             console.log(prefix + "Finished attachment %s", filename);
//         });
//     };
// }

function findAttachment(attrs: ImapMessageAttributes): Array<any> {
    let attachments: Array<any> = [];
    attrs.struct?.forEach((element) => {
        if (!Array.isArray(element)) return;
        element.forEach((part) => {
            if (
                part.disposition &&
                ["INLINE", "ATTACHMENT"].indexOf(
                    part.disposition.type.toUpperCase()
                )
            ) {
                attachments.push(part);
            }
        });
    });

    return attachments;
}

function streamToString(stream: NodeJS.ReadableStream): Promise<string> {
    const chunks: any[] = [];
    return new Promise((resolve, reject) => {
        stream.on("data", (chunk) => chunks.push(Buffer.from(chunk)));
        stream.on("error", (err) => reject(err));
        stream.on("end", () => resolve(Buffer.concat(chunks).toString("utf8")));
    });
}

imap.once("ready", function () {
    openInbox(function (err: Error, box: Imap.Box) {
        if (err) throw err;
        const f = imap.seq.fetch(`1:${box.messages.total}`, {
            bodies: "HEADER.FIELDS (FROM TO SUBJECT DATE)",
            struct: true,
        });
        f.on("message", function (msg, seqno) {
            console.log("Message #%d", seqno);
            const prefix = "(#" + seqno + ") ";
            // msg.on("body", function (stream, _info) {
            //   var buffer = "";
            //   stream.on("data", function (chunk) {
            //     buffer += chunk.toString("utf8");
            //   });
            //   stream.once("end", function () {
            //     console.log(
            //       prefix + "Parsed header: %s",
            //       inspect(Imap.parseHeader(buffer), { colors: true })
            //     );
            //   });
            // });
            msg.once("attributes", function (attrs: ImapMessageAttributes) {
                // console.log(`${prefix} Attributes`, inspect(attrs, { depth: null, colors: true }));
                const attachments = findAttachment(attrs);
                // console.log(inspect(attachments, { colors: true, depth: null }))

                attachments.forEach((attachment) => {
                    let fetch = imap.fetch(attrs.uid, {
                        bodies: attachment.partID,
                        struct: true,
                    });
                    fetch.on("message", function (msg, seqno) {
                        msg.on("body", async function (stream, info) {
                            const filename =
                                attachment.disposition.params.filename.match(
                                    /(.+?)(\.[^.]*$|$)/
                                )[1];
                            // const writeStream = fs.createWriteStream(`./reports/${filename}`);

                            // const deflated = Buffer.from(await streamToString(stream), 'base64').toString('utf8');
                            // const inflated = zlib.inflateSync(Buffer.from(deflated, 'utf-8')).toString();

                            // if (attachment.encoding.toLowerCase() === 'base64' && attachment.subtype.search(/gzip/) !== -1) stream.pipe(new Base64Decode()).pipe(zlib.createGunzip()).on('data', function(d) {
                            //   data = JSON.parse(d.toString());
                            // });
                            // else if (attachment.encoding.toLowerCase() === 'base64') stream.pipe(new Base64Decode()).pipe(writeStream);
                            // else if (attachment.subtype.search(/gzip/) !== -1) stream.pipe(zlib.createGunzip()).pipe(writeStream);

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
                                    const report = new Report(
                                        JSON.parse(tlsreport),
                                        filename
                                    );
                                    report.writeReports();
                                });
                            // console.log(report);
                        });
                    });
                });

                // console.log(prefix + 'Attributes: %s', inspect(attrs, false, 8));
                // var attachments = findAttachmentParts(attrs.struct);
                // console.log(prefix + "Has attachments: %d", attachments.length);
                // for (var i = 0, len = attachments.length; i < len; ++i) {
                //     var attachment = attachments[i];
                /*This is how each attachment looks like {
  partID: '2',
  type: 'application',
  subtype: 'octet-stream',
  params: { name: 'file-name.ext' },
  id: null,
  description: null,
  encoding: 'BASE64',
  size: 44952,
  md5: null,
  disposition: { type: 'ATTACHMENT', params: { filename: 'file-name.ext' } },
  language: null
  }
  */
                // console.log(
                //     prefix + "Fetching attachment %s",
                //     attachment.params.name
                // );
                // var f = imap.fetch(attrs.uid, {
                //     //do not use imap.seq.fetch here
                //     bodies: [attachment.partID],
                //     struct: true,
                // });
                // //build function to process attachment message
                // f.on("message", buildAttMessageFunction(attachment));
                // }
            });
            msg.once("end", function () {
                console.log(prefix + "Finished");
            });
        });
        f.once("error", function (fetchError) {
            console.log("Fetch error: " + fetchError);
        });
        f.once("end", function () {
            console.log("Done fetching all messages!");
            imap.end();
        });
    });
});

imap.once("error", function (error: any) {
    console.log(error);
});

imap.once("end", function () {
    console.log("Connection ended");
});

imap.connect();
