import fs from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";

const packageRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const inputRoot = path.resolve(process.env.BP_INPUT_ROOT ?? path.join(packageRoot, "inputs"));
const outputRoot = path.resolve(process.env.BP_OUTPUT_ROOT ?? path.join(packageRoot, "generated"));
const outputDir = path.join(outputRoot, "doi_205_finalization_2026-11-16");
const s1Input = path.join(outputRoot, "doi_final_integration_2026-11-16", "Supplementary_File_S1_Included_Study_Register_202.xlsx");
const s2Input = path.join(outputRoot, "doi_final_integration_2026-11-16", "Supplementary_File_S2_Evidence_Matrix_202.xlsx");
const deltaInput = path.join(packageRoot, "inputs", "p320_p322_delta.json");
const s1Output = path.join(outputDir, "Supplementary_File_S1_Included_Study_Register_205.xlsx");
const s2Output = path.join(outputDir, "Supplementary_File_S2_Evidence_Matrix_205.xlsx");

const appendedIds = new Set(["P320", "P321", "P322"]);
const delta = JSON.parse(await fs.readFile(deltaInput, "utf8"));

if (JSON.stringify(delta.cluster_ids) !== JSON.stringify([...appendedIds])) {
  throw new Error(`Unexpected delta cluster IDs: ${JSON.stringify(delta.cluster_ids)}`);
}
if (delta.s1_rows_A_to_L.length !== 3 ||
    delta.s2_cluster_codes_A_to_T.length !== 3 ||
    delta.s2_evidence_anchors_A_to_F.length !== 3 ||
    delta.s2_publication_status_A_to_Y.length !== 3 ||
    delta.s2_multilabel_rows_A_to_E.length !== 64) {
  throw new Error("The compact P320-P322 delta has unexpected row counts.");
}

function pct(n, denominator) {
  return Math.round((n / denominator) * 1000) / 10;
}

function findRow(sheet, clusterId, maxRow = 400) {
  const values = sheet.getRange(`A1:A${maxRow}`).values;
  const index = values.findIndex((row) => row[0] === clusterId);
  if (index < 0) throw new Error(`${clusterId} not found in ${sheet.name}`);
  return index + 1;
}

function countBy(rows, index) {
  const counts = {};
  for (const row of rows) {
    const key = row[index] ?? "";
    counts[key] = (counts[key] ?? 0) + 1;
  }
  return counts;
}

function countMulti(rows, dimensionCode, allowedIds = null) {
  const counts = {};
  for (const row of rows) {
    if (row[1] !== dimensionCode) continue;
    if (allowedIds && !allowedIds.has(row[0])) continue;
    counts[row[3]] = (counts[row[3]] ?? 0) + 1;
  }
  return counts;
}

function validateUniqueAssignments(rows) {
  const keys = new Set();
  for (const row of rows) {
    const key = `${row[0]}\u0000${row[1]}\u0000${row[3]}`;
    if (keys.has(key)) throw new Error(`Duplicate multi-label assignment: ${key}`);
    keys.add(key);
  }
}

function updateSensitivity(workbook, clusterRows, multilabelRows) {
  const sensitivity = workbook.worksheets.getItem("Sensitivity");
  const existing = sensitivity.getRange("A5:H58").values;
  const secondary = new Set(
    clusterRows
      .filter((row) => row[6] === "SECONDARY_SYNTHESIS_WITH_CONCEPTUAL_OUTPUT")
      .map((row) => row[0]),
  );
  const all = new Set(clusterRows.map((row) => row[0]));
  const nonSecondary = new Set([...all].filter((id) => !secondary.has(id)));
  const assignments = new Set(
    multilabelRows.map((row) => `${row[0]}\u0000${row[1]}\u0000${row[3]}`),
  );
  const stageStatus = new Map(clusterRows.map((row) => [row[0], row[14]]));
  const implementation = new Map(clusterRows.map((row) => [row[0], row[16]]));
  const dimensionCode = new Map([
    ["RQ2 assessment form", "rq2Canonical"],
    ["RQ3 application context", "contexts"],
    ["RQ3 teamwork stage", "stages"],
    ["RQ3 validation design", "validationDesign"],
  ]);

  const counted = existing.map((row) => {
    const [dimension, categoryCode, category] = row;
    const count = (ids) => {
      let n = 0;
      if (dimension === "RQ3 stage evidence status") {
        for (const id of ids) if (stageStatus.get(id) === categoryCode) n += 1;
      } else if (dimension === "RQ3 implementation form") {
        for (const id of ids) if (implementation.get(id) === categoryCode) n += 1;
      } else {
        const code = dimensionCode.get(dimension);
        if (!code) throw new Error(`Unknown sensitivity dimension: ${dimension}`);
        for (const id of ids) {
          if (assignments.has(`${id}\u0000${code}\u0000${categoryCode}`)) n += 1;
        }
      }
      return n;
    };
    const fullN = count(all);
    const restrictedN = count(nonSecondary);
    return [
      dimension,
      categoryCode,
      category,
      fullN,
      pct(fullN, all.size),
      restrictedN,
      pct(restrictedN, nonSecondary.size),
      fullN - restrictedN,
    ];
  });

  if (all.size !== 205) throw new Error(`Expected 205 clusters, got ${all.size}`);
  if (secondary.size !== 33) throw new Error(`Expected 33 secondary clusters, got ${secondary.size}`);
  if (nonSecondary.size !== 172) {
    throw new Error(`Expected 172 non-secondary clusters, got ${nonSecondary.size}`);
  }
  sensitivity.getRange("A2").values = [[
    "Full percentages use N=205. Restricted percentages use N=172 after excluding 33 secondary-synthesis contribution clusters.",
  ]];
  sensitivity.getRange("A3").values = [[
    "Counts were regenerated directly from the normalized final 205-cluster coding.",
  ]];
  sensitivity.getRange("A5:H58").values = counted;
  return { counted, secondary, nonSecondary };
}

async function buildS1() {
  const workbook = await SpreadsheetFile.importXlsx(await FileBlob.load(s1Input));
  const sheet = workbook.worksheets.getItem("Included studies");
  const oldTable = sheet.tables.getItem("IncludedStudies202");
  const tableStyle = oldTable.style;
  sheet.getRange("A207:L209").values = delta.s1_rows_A_to_L;

  const p314 = findRow(sheet, "P314");
  sheet.getRange(`L${p314}`).values = [[
    "OpenReview HEAI 2026 Poster page and API record; published 11 October 2026; checked 2026-11-16.",
  ]];

  oldTable.delete();
  const finalTable = sheet.tables.add("A4:L209", true, "IncludedStudies205");
  finalTable.style = tableStyle;
  sheet.getRange("A1").values = [["Supplementary File S1: Included Study Register"]];
  sheet.getRange("A2").values = [[
    "One row represents one study-contribution cluster. P250 contains two eligible companion reports, so 205 clusters represent 206 reports.",
  ]];
  sheet.getRange("A3").values = [[
    "Final register after applying the final I1--I5 criteria and completing full-text retrieval and version reconciliation.",
  ]];

  workbook.recalculate();
  const output = await SpreadsheetFile.exportXlsx(workbook);
  await output.save(s1Output);
}

async function buildS2() {
  const workbook = await SpreadsheetFile.importXlsx(await FileBlob.load(s2Input));
  const clusterSheet = workbook.worksheets.getItem("Cluster codes");
  const multiSheet = workbook.worksheets.getItem("Multilabel codes");
  const anchorSheet = workbook.worksheets.getItem("Evidence anchors");
  const publicationSheet = workbook.worksheets.getItem("Publication status");

  clusterSheet.getRange("A207:T209").values = delta.s2_cluster_codes_A_to_T;
  anchorSheet.getRange("A207:F209").values = delta.s2_evidence_anchors_A_to_F;
  publicationSheet.getRange("A208:Y210").values = delta.s2_publication_status_A_to_Y;
  multiSheet.getRange("A2837:E2900").values = delta.s2_multilabel_rows_A_to_E;

  const p311Publication = findRow(publicationSheet, "P311");
  publicationSheet.getRange(`R${p311Publication}`).values = [["1629-1631"]];

  const p313Publication = findRow(publicationSheet, "P313");
  publicationSheet.getRange(`S${p313Publication}`).values = [[
    "https://ceur-ws.org/Vol-4256/XR4CE_paper_7.pdf",
  ]];

  const p314Cluster = findRow(clusterSheet, "P314");
  clusterSheet.getRange(`S${p314Cluster}:T${p314Cluster}`).values = [[
    "FORMALLY_PUBLISHED",
    "NO_SUCCESSOR_SEARCH_REQUIRED_FOR_RECORDED_FORMAL_VERSION",
  ]];
  const p314Publication = findRow(publicationSheet, "P314");
  publicationSheet.getRange(`I${p314Publication}:Y${p314Publication}`).values = [[
    "FORMALLY_PUBLISHED",
    "NOT_SEPARATELY_VERIFIED",
    "HIGH_OFFICIAL_VENUE_AND_REPOSITORY_IDENTITY",
    "workshop-poster",
    "Second Workshop on Human-aware Embodied AI",
    "HEAI 2026 Poster (IROS 2026 Workshop)",
    "2026-10-11",
    null,
    null,
    "Submission 15",
    "https://openreview.net/forum?id=bqQ5QobPpT",
    "OpenReview public page and API note; venue, accepted-release invitation, and publication date checked 2026-11-16",
    "2026-11-16",
    null,
    "NO_SUCCESSOR_SEARCH_REQUIRED_FOR_RECORDED_FORMAL_VERSION",
    0,
    "OpenReview identifies the record as HEAI 2026 Poster and published it on 11 October 2026. No item-level DOI was recorded.",
  ]];

  for (const [sheet, tableName, range] of [
    [clusterSheet, "ClustercodesTable", "A4:T209"],
    [multiSheet, "MultilabelcodesTable", "A4:E2900"],
    [anchorSheet, "EvidenceanchorsTable", "A4:F209"],
    [publicationSheet, "PublicationstatusTable", "A4:Y210"],
  ]) {
    const oldTable = sheet.tables.getItem(tableName);
    const style = oldTable.style;
    oldTable.delete();
    const table = sheet.tables.add(range, true, tableName);
    table.style = style;
  }

  clusterSheet.getRange("A3").values = [[
    "Baseline canonical SHA-256: 31655db590049900d965f1a9b97c9d4554c864ccf77eba3a56c1524801fce7ce. P311--P322 were added during final protocol and retrieval reconciliation; the frozen 193-cluster source is preserved.",
  ]];
  publicationSheet.getRange("A3").values = [[
    "Baseline reports were checked 2026-11-10. Rows P311--P322 use the authoritative-source checks recorded from 2026-11-14 to 2026-11-16.",
  ]];

  const clusterRows = clusterSheet.getRange("A5:T209").values;
  const multilabelRows = multiSheet.getRange("A5:E2900").values;
  validateUniqueAssignments(multilabelRows);
  const sensitivity = updateSensitivity(workbook, clusterRows, multilabelRows);

  const readme = workbook.worksheets.getItem("ReadMe");
  readme.getRange("A1").values = [["Supplementary File S2: Evidence Matrix and Codebook"]];
  readme.getRange("A2").values = [[
    "This workbook contains the final 205-cluster evidence matrix, evidence anchors, codebook, publication/version audit, and sensitivity analysis.",
  ]];
  readme.getRange("B5").values = [["2026-11-16"]];
  readme.getRange("B6").values = [[
    "205 study-contribution clusters representing 206 eligible reports",
  ]];
  readme.getRange("B8").values = [[
    "Supplementary_File_S1_Included_Study_Register_205.xlsx",
  ]];
  readme.getRange("B9").values = [[
    "Frozen 193-cluster S2 plus full-text extraction and coding for P311--P322",
  ]];
  readme.getRange("B10").values = [[
    "Frozen 194-report audit plus authoritative-source checks for P311--P322",
  ]];
  readme.getRange("B13").values = [["N = 205 clusters"]];
  readme.getRange("B14").values = [[
    "N = 172 clusters after excluding 33 secondary-synthesis contribution clusters",
  ]];
  readme.getRange("B15").values = [[205]];
  readme.getRange("B16").values = [[multilabelRows.length]];
  readme.getRange("B17").values = [[206]];
  readme.getRange("B18").values = [[
    "Sensitivity rows were recalculated from the normalized final 205-cluster coding.",
  ]];

  const publicationCounts = countBy(clusterRows, 18);
  const summary = {
    generated: "2026-11-16",
    reports: 206,
    clusters: clusterRows.length,
    secondary: sensitivity.secondary.size,
    nonSecondary: sensitivity.nonSecondary.size,
    multilabelAssignments: multilabelRows.length,
    contributionType: countBy(clusterRows, 6),
    alignmentRole: countBy(clusterRows, 8),
    evidenceTier: countBy(clusterRows, 10),
    measurementMaturity: countBy(clusterRows, 12),
    stageEvidenceStatus: countBy(clusterRows, 14),
    implementation: countBy(clusterRows, 16),
    publicationStatus: publicationCounts,
    rq1Families: countMulti(multilabelRows, "rq1Families"),
    rq2Forms: countMulti(multilabelRows, "rq2Canonical"),
    contexts: countMulti(multilabelRows, "contexts"),
    stages: countMulti(multilabelRows, "stages"),
    validationDesign: countMulti(multilabelRows, "validationDesign"),
    nonSecondaryCounts: {
      rq2Forms: countMulti(multilabelRows, "rq2Canonical", sensitivity.nonSecondary),
      contexts: countMulti(multilabelRows, "contexts", sensitivity.nonSecondary),
      stages: countMulti(multilabelRows, "stages", sensitivity.nonSecondary),
      validationDesign: countMulti(multilabelRows, "validationDesign", sensitivity.nonSecondary),
      implementation: countBy(
        clusterRows.filter((row) => sensitivity.nonSecondary.has(row[0])),
        16,
      ),
      stageEvidenceStatus: countBy(
        clusterRows.filter((row) => sensitivity.nonSecondary.has(row[0])),
        14,
      ),
    },
  };

  if (publicationCounts.FORMALLY_PUBLISHED !== 156 ||
      publicationCounts.ONLINE_FIRST_FORMALLY_PUBLISHED !== 2 ||
      publicationCounts.PREPRINT_ARXIV !== 37 ||
      publicationCounts.PREPRINT_SSRN !== 8 ||
      publicationCounts.PREPRINT_RESEARCH_SQUARE !== 1 ||
      publicationCounts.PUBLISHED_PLATFORM_AWAITING_PEER_REVIEW !== 1) {
    throw new Error(`Unexpected publication status counts: ${JSON.stringify(publicationCounts)}`);
  }

  workbook.recalculate();
  const output = await SpreadsheetFile.exportXlsx(workbook);
  await output.save(s2Output);
  await fs.writeFile(
    `${outputDir}/final_205_summary.json`,
    `${JSON.stringify(summary, null, 2)}\n`,
    "utf8",
  );
  return summary;
}

await fs.mkdir(outputDir, { recursive: true });
await buildS1();
const summary = await buildS2();
console.log(JSON.stringify({ s1Output, s2Output, summary }, null, 2));
