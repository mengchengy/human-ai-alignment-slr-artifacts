import fs from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";

const packageRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const outputRoot = path.resolve(process.env.BP_OUTPUT_ROOT ?? path.join(packageRoot, "generated"));
const dir = path.join(outputRoot, "doi_205_finalization_2026-11-16");
const previewDir = `${dir}/preview`;
await fs.mkdir(previewDir, { recursive: true });

const s1 = await SpreadsheetFile.importXlsx(
  await FileBlob.load(`${dir}/Supplementary_File_S1_Included_Study_Register_205.xlsx`),
);
const s2 = await SpreadsheetFile.importXlsx(
  await FileBlob.load(`${dir}/Supplementary_File_S2_Evidence_Matrix_205.xlsx`),
);

const checks = {};
checks.s1Tail = s1.worksheets.getItem("Included studies").getRange("A203:L209").values;
checks.s2ReadMe = s2.worksheets.getItem("ReadMe").getRange("A1:B18").values;
checks.clusterTail = s2.worksheets.getItem("Cluster codes").getRange("A203:T209").values;
checks.anchorTail = s2.worksheets.getItem("Evidence anchors").getRange("A203:F209").values;
checks.publicationTail = s2.worksheets.getItem("Publication status").getRange("A203:Y210").values;
checks.sensitivity = s2.worksheets.getItem("Sensitivity").getRange("A1:H58").values;

const p311Row = s2.worksheets.getItem("Publication status").getRange("A1:Y210").values.find((row) => row[0] === "P311");
const p314Cluster = s2.worksheets.getItem("Cluster codes").getRange("A1:T209").values.find((row) => row[0] === "P314");
const p314Publication = s2.worksheets.getItem("Publication status").getRange("A1:Y210").values.find((row) => row[0] === "P314");
if (!p311Row || p311Row[17] !== "1629-1631") throw new Error("P311 page range not corrected");
if (!p314Cluster || p314Cluster[18] !== "FORMALLY_PUBLISHED") throw new Error("P314 cluster status not corrected");
if (!p314Publication || p314Publication[8] !== "FORMALLY_PUBLISHED" || p314Publication[14] !== "2026-10-11") {
  throw new Error("P314 publication audit not corrected");
}

for (const [name, workbook] of [["s1", s1], ["s2", s2]]) {
  const errors = await workbook.inspect({
    kind: "match",
    searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A|#NUM!|#NULL!|#SPILL!|#CALC!",
    options: { useRegex: true, maxResults: 300 },
    summary: `${name} formula error scan`,
  });
  checks[`${name}Errors`] = errors.ndjson;
}

for (const [name, workbook, sheetName, range] of [
  ["s1_tail", s1, "Included studies", "A200:L209"],
  ["s2_readme", s2, "ReadMe", "A1:B18"],
  ["s2_cluster_tail", s2, "Cluster codes", "A202:T209"],
  ["s2_anchor_tail", s2, "Evidence anchors", "A202:F209"],
  ["s2_publication_tail", s2, "Publication status", "A198:Y210"],
  ["s2_sensitivity", s2, "Sensitivity", "A1:H58"],
]) {
  const image = await workbook.render({ sheetName, range, scale: 1.5, format: "png" });
  await fs.writeFile(`${previewDir}/${name}.png`, new Uint8Array(await image.arrayBuffer()));
}

await fs.writeFile(`${dir}/validation_report.json`, `${JSON.stringify(checks, null, 2)}\n`, "utf8");
console.log(JSON.stringify({
  p311Pages: p311Row[17],
  p314Status: p314Publication[8],
  p314Date: p314Publication[14],
  s1TailIds: checks.s1Tail.map((row) => row[0]),
  clusterTailIds: checks.clusterTail.map((row) => row[0]),
  s1Errors: checks.s1Errors,
  s2Errors: checks.s2Errors,
}, null, 2));
