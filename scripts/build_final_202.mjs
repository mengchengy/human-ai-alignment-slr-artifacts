import fs from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";
import {
  FileBlob,
  SpreadsheetFile,
} from "@oai/artifact-tool";

process.on("uncaughtException", (error) => {
  console.error(`UNCAUGHT ${error.name}: ${error.message}`);
  process.exit(1);
});
process.on("unhandledRejection", (error) => {
  console.error(`UNHANDLED ${error?.name ?? "Error"}: ${error?.message ?? error}`);
  process.exit(1);
});

const packageRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const inputRoot = path.resolve(process.env.BP_INPUT_ROOT ?? path.join(packageRoot, "inputs"));
const outputRoot = path.resolve(process.env.BP_OUTPUT_ROOT ?? path.join(packageRoot, "generated"));
const sourceDir = path.join(inputRoot, "doi_eight_candidate_integration_2026-11-16");
const outputDir = path.join(outputRoot, "doi_final_integration_2026-11-16");
const s1Source = path.join(sourceDir, "Supplementary_File_S1_Candidate_Integration_201.xlsx");
const s2Source = path.join(sourceDir, "Supplementary_File_S2_Candidate_Integration_201.xlsx");
const s1Output = path.join(outputDir, "Supplementary_File_S1_Included_Study_Register_202.xlsx");
const s2Output = path.join(outputDir, "Supplementary_File_S2_Evidence_Matrix_202.xlsx");

const f016 = {
  clusterId: "P319",
  reportId: "P319",
  authorYear: "Yu et al. (2025)",
  authors: "Yu, Zixiao and Cui, Tingru and Luo, Chen and Tan, Dilang",
  year: 2025,
  title:
    "A Systematic Literature Review on Human-Agent Teaming with Insights into Multi-Agent Interactions",
  sourceBranch: "Forward citation search",
  contributionTypeCode: "SECONDARY_SYNTHESIS_WITH_CONCEPTUAL_OUTPUT",
  contributionType: "Secondary synthesis with conceptual output",
  alignmentRoleCode: "RELATED_OPERATIONALIZATION",
  alignmentRole: "Related operationalization",
  evidenceTierCode: "D_CONCEPTUAL_PROPOSED_OR_SECONDARY",
  evidenceTier: "D — conceptual, proposed, or secondary support",
  measurementMaturityCode: "CONCEPTUAL_NO_DIRECT_MEASURE",
  measurementMaturity: "Conceptual, no direct measure",
  stageEvidenceStatusCode: "PROPOSED_OR_SYNTHESIZED_STAGE",
  stageEvidenceStatus: "Proposed or synthesized",
  implementationCode: "NO_FUNCTIONING_AI_SYSTEM",
  implementation: "No functioning AI system",
  publicationStatus: "FORMALLY_PUBLISHED",
  laterVersionStatus: "NO_SUCCESSOR_SEARCH_REQUIRED_FOR_RECORDED_FORMAL_VERSION",
  metadataSource:
    "Official AISeL PACIS 2025 proceedings record and checked full text; DBLP conf/pacis/YuCLT25; checked 2026-11-16",
  multi: {
    rq1Families: [
      "ROLES_RESPONSIBILITIES_ALLOCATION",
      "AUTONOMY_CONTROL_INITIATIVE",
      "WORKFLOW_CONSTRAINTS_OUTPUTS",
      "INFORMATION_COMMUNICATION",
      "CAPABILITY_RELIANCE_CALIBRATION",
      "GOALS_VALUES_PREFERENCES",
      "KNOWLEDGE_COMMON_GROUND_MENTAL_MODELS",
      "SOCIO_EMOTIONAL_EXPECTATIONS",
    ],
    rq2Canonical: [
      "CONCEPTUAL_FRAMEWORK_OR_DESIGN_GUIDANCE",
      "BENCHMARK_RUBRIC_TAXONOMY",
      "OUTCOME_OR_PROXY_MEASURE",
    ],
    contexts: ["DOMAIN_GENERAL_OR_CROSS_DOMAIN"],
    stages: ["DESIGN_TIME", "PRE_TASK", "RUNTIME", "TRANSITION", "POST_TASK"],
    validationDesign: [
      "SECONDARY_SYNTHESIS",
      "CORPUS_OR_TAXONOMY_VALIDATION",
      "CONCEPTUAL_OR_PROPOSAL",
    ],
  },
  anchors: {
    activeLocation:
      "Physical PDF p. 4, Human-Multi-Agent Teaming criteria; pp. 9–10, cross-composition synthesis",
    activeSummary:
      "The review requires HMAT agents to hold distinct or mission-critical roles and to exceed merely suggesting decisions and waiting for approval. It also synthesizes autonomous task allocation and reallocation in multi-agent teams.",
    alignmentLocation:
      "Physical PDF pp. 7–9, key variables and modified I-M-O model; pp. 10–12, composition findings and contextual dimensions",
    alignmentSummary:
      "The synthesis relates appropriate autonomy, task and role allocation, communication, shared mental models, transparency and reliability, value similarity, and calibrated trust to human expectations, team processes, and contextual requirements.",
    boundary:
      "This is review-level, Tier D evidence. The 58 underlying studies are heterogeneous, and some use Wizard-of-Oz or reactive systems; their implementations and outcomes are not coded as primary validation performed by this report. The taxonomy and modified I-M-O model are conceptual syntheses rather than a validated direct Alignment measure.",
  },
  publication: {
    peerReviewStatus: "NOT_SEPARATELY_VERIFIED",
    statusConfidence: "HIGH_OFFICIAL_VENUE_AND_REPOSITORY_IDENTITY",
    recordType: "proceedings-article",
    publisher: "Association for Information Systems",
    venue: "PACIS 2025 Proceedings",
    publicationDate: "2025-07",
    pageOrArticle: "20",
    authoritativeUrl: "https://aisel.aisnet.org/pacis2025/hci/hci/20/",
    verificationSource:
      "Official AISeL proceedings record and checked full text; DBLP conf/pacis/YuCLT25",
    dateChecked: "2026-11-16",
    potentialSuccessorCount: 0,
    notes:
      "Completed research paper. No item-level DOI was verified. Only this review's synthesis is coded; its 58 underlying studies are not counted as additional reports or clusters.",
  },
};

const multiLabels = new Map([
  ["ROLES_RESPONSIBILITIES_ALLOCATION", "Roles, responsibilities, and task allocation"],
  ["AUTONOMY_CONTROL_INITIATIVE", "Autonomy, control, and initiative"],
  ["WORKFLOW_CONSTRAINTS_OUTPUTS", "Workflows, constraints, and outputs"],
  ["INFORMATION_COMMUNICATION", "Information and communication"],
  ["CAPABILITY_RELIANCE_CALIBRATION", "Capability and reliance calibration"],
  ["GOALS_VALUES_PREFERENCES", "Goals, values, and preferences"],
  ["KNOWLEDGE_COMMON_GROUND_MENTAL_MODELS", "Knowledge, common ground, and mental models"],
  ["SOCIO_EMOTIONAL_EXPECTATIONS", "Socio-emotional expectations"],
  ["CONCEPTUAL_FRAMEWORK_OR_DESIGN_GUIDANCE", "Conceptual framework or design guidance"],
  ["BENCHMARK_RUBRIC_TAXONOMY", "Benchmark, rubric, or taxonomy"],
  ["OUTCOME_OR_PROXY_MEASURE", "Outcome or proxy measure"],
  ["DOMAIN_GENERAL_OR_CROSS_DOMAIN", "Domain-general or cross-domain"],
  ["DESIGN_TIME", "Design-time"],
  ["PRE_TASK", "Pre-task"],
  ["RUNTIME", "Runtime"],
  ["TRANSITION", "Transition or handoff"],
  ["POST_TASK", "Post-task"],
  ["SECONDARY_SYNTHESIS", "Secondary synthesis"],
  ["CORPUS_OR_TAXONOMY_VALIDATION", "Corpus or taxonomy validation"],
  ["CONCEPTUAL_OR_PROPOSAL", "Conceptual or proposal"],
]);

const dimensionLabels = {
  rq1Families: "RQ1 family",
  rq2Canonical: "RQ2 assessment form",
  contexts: "RQ3 application context",
  stages: "RQ3 teamwork stage",
  validationDesign: "RQ3 validation design",
};

function pct(n, denominator) {
  return Math.round((n / denominator) * 1000) / 10;
}

function normalizedMultilabelRows() {
  const rows = [];
  for (const [dimensionCode, codes] of Object.entries(f016.multi)) {
    for (const code of codes) {
      const label = multiLabels.get(code);
      if (!label) throw new Error(`Missing label for ${code}`);
      rows.push([
        f016.clusterId,
        dimensionCode,
        dimensionLabels[dimensionCode],
        code,
        label,
      ]);
    }
  }
  return rows;
}

function validateUniqueAssignments(rows) {
  const keys = new Set();
  for (const row of rows) {
    const key = `${row[0]}\u0000${row[1]}\u0000${row[3]}`;
    if (keys.has(key)) throw new Error(`Duplicate multi-label assignment: ${key}`);
    keys.add(key);
  }
}

async function buildS1() {
  const input = await FileBlob.load(s1Source);
  const workbook = await SpreadsheetFile.importXlsx(input);
  const sheet = workbook.worksheets.getItem("Included studies");
  const table = sheet.tables.getItem("IncludedStudies193");
  const tableStyle = table.style;

  sheet.getRange("F205").values = [[
    "Collaborative Gym: A Framework for Enabling and Evaluating Human-Agent Collaboration",
  ]];
  sheet.getRange("H205").values = [[null]];
  sheet.getRange("L205").values = [[
    "Official ICLR 2026 record and linked earlier arXiv version checked 2026-11-14",
  ]];

  sheet.getRange("A206:L206").values = [[
    f016.clusterId,
    f016.reportId,
    f016.authorYear,
    f016.authors,
    f016.year,
    f016.title,
    null,
    null,
    f016.sourceBranch,
    f016.contributionType,
    "D",
    f016.metadataSource,
  ]];
  table.delete();
  const finalTable = sheet.tables.add("A4:L206", true, "IncludedStudies202");
  finalTable.style = tableStyle;

  sheet.getRange("A1").values = [[
    "Supplementary File S1: Included Study Register (DOI protocol amendment)",
  ]];
  sheet.getRange("A2").values = [[
    "One row represents one study-contribution cluster. P250 contains two eligible companion reports, so 202 clusters represent 203 reports.",
  ]];
  sheet.getRange("A3").values = [[
    "Final integration of the nine reports restored after removal of the item-level DOI eligibility criterion. The frozen 193-cluster register remains the pre-amendment baseline.",
  ]];

  workbook.recalculate();
  const output = await SpreadsheetFile.exportXlsx(workbook);
  await output.save(s1Output);
}

function updateSensitivity(workbook, clusterRows, multilabelRows) {
  const sensitivity = workbook.worksheets.getItem("Sensitivity");
  const existing = sensitivity.getRange("A5:H58").values;
  const contributionIndex = 6;
  const stageStatusIndex = 14;
  const implementationIndex = 16;
  const secondaryClusters = new Set(
    clusterRows
      .filter((row) => row[contributionIndex] === "SECONDARY_SYNTHESIS_WITH_CONCEPTUAL_OUTPUT")
      .map((row) => row[0]),
  );
  const allClusters = new Set(clusterRows.map((row) => row[0]));
  const nonSecondary = new Set(
    [...allClusters].filter((clusterId) => !secondaryClusters.has(clusterId)),
  );
  const assignments = new Map();
  for (const row of multilabelRows) {
    assignments.set(`${row[0]}\u0000${row[1]}\u0000${row[3]}`, true);
  }
  const stageStatusByCluster = new Map(
    clusterRows.map((row) => [row[0], row[stageStatusIndex]]),
  );
  const implementationByCluster = new Map(
    clusterRows.map((row) => [row[0], row[implementationIndex]]),
  );
  const dimensionCode = new Map([
    ["RQ2 assessment form", "rq2Canonical"],
    ["RQ3 application context", "contexts"],
    ["RQ3 teamwork stage", "stages"],
    ["RQ3 validation design", "validationDesign"],
  ]);

  const counted = existing.map((row) => {
    const [dimension, categoryCode, category] = row;
    let fullN = 0;
    let restrictedN = 0;
    if (dimension === "RQ3 stage evidence status") {
      for (const clusterId of allClusters) {
        if (stageStatusByCluster.get(clusterId) === categoryCode) fullN += 1;
      }
      for (const clusterId of nonSecondary) {
        if (stageStatusByCluster.get(clusterId) === categoryCode) restrictedN += 1;
      }
    } else if (dimension === "RQ3 implementation form") {
      for (const clusterId of allClusters) {
        if (implementationByCluster.get(clusterId) === categoryCode) fullN += 1;
      }
      for (const clusterId of nonSecondary) {
        if (implementationByCluster.get(clusterId) === categoryCode) restrictedN += 1;
      }
    } else {
      const code = dimensionCode.get(dimension);
      if (!code) throw new Error(`Unknown sensitivity dimension: ${dimension}`);
      for (const clusterId of allClusters) {
        if (assignments.has(`${clusterId}\u0000${code}\u0000${categoryCode}`)) fullN += 1;
      }
      for (const clusterId of nonSecondary) {
        if (assignments.has(`${clusterId}\u0000${code}\u0000${categoryCode}`)) restrictedN += 1;
      }
    }
    return [
      dimension,
      categoryCode,
      category,
      fullN,
      pct(fullN, allClusters.size),
      restrictedN,
      pct(restrictedN, nonSecondary.size),
      fullN - restrictedN,
    ];
  });

  if (allClusters.size !== 202) throw new Error(`Expected 202 clusters, got ${allClusters.size}`);
  if (secondaryClusters.size !== 32) {
    throw new Error(`Expected 32 secondary clusters, got ${secondaryClusters.size}`);
  }
  if (nonSecondary.size !== 170) {
    throw new Error(`Expected 170 non-secondary clusters, got ${nonSecondary.size}`);
  }

  sensitivity.getRange("A2").values = [[
    "Full percentages use N=202. Restricted percentages use N=170 after excluding 32 secondary-synthesis contribution clusters.",
  ]];
  sensitivity.getRange("A3").values = [[
    "Counts were regenerated directly from the normalized 202-cluster DOI-amendment rows; the nine restored reports were checked against their full texts and current codebook boundaries.",
  ]];
  sensitivity.getRange("A5:H58").values = counted;
  return { counted, clusterCount: allClusters.size, secondaryCount: secondaryClusters.size };
}

async function buildS2() {
  const input = await FileBlob.load(s2Source);
  const workbook = await SpreadsheetFile.importXlsx(input);
  const clusterSheet = workbook.worksheets.getItem("Cluster codes");
  const multiSheet = workbook.worksheets.getItem("Multilabel codes");
  const anchorSheet = workbook.worksheets.getItem("Evidence anchors");
  const publicationSheet = workbook.worksheets.getItem("Publication status");
  const newMultiRows = normalizedMultilabelRows();

  clusterSheet.getRange("D205:E205").values = [[
    "Collaborative Gym: A Framework for Enabling and Evaluating Human-Agent Collaboration",
    null,
  ]];
  publicationSheet.getRange("F206").values = [[
    "Collaborative Gym: A Framework for Enabling and Evaluating Human-Agent Collaboration",
  ]];

  clusterSheet.getRange("A206:T206").values = [[
    f016.clusterId,
    f016.reportId,
    f016.authorYear,
    f016.title,
    null,
    f016.sourceBranch,
    f016.contributionTypeCode,
    f016.contributionType,
    f016.alignmentRoleCode,
    f016.alignmentRole,
    f016.evidenceTierCode,
    f016.evidenceTier,
    f016.measurementMaturityCode,
    f016.measurementMaturity,
    f016.stageEvidenceStatusCode,
    f016.stageEvidenceStatus,
    f016.implementationCode,
    f016.implementation,
    f016.publicationStatus,
    f016.laterVersionStatus,
  ]];
  multiSheet.getRange("A2817:E2836").values = newMultiRows;
  anchorSheet.getRange("A206:F206").values = [[
    f016.clusterId,
    f016.anchors.activeLocation,
    f016.anchors.activeSummary,
    f016.anchors.alignmentLocation,
    f016.anchors.alignmentSummary,
    f016.anchors.boundary,
  ]];
  publicationSheet.getRange("A207:Y207").values = [[
    f016.clusterId,
    f016.reportId,
    f016.authorYear,
    f016.authors,
    f016.year,
    f016.title,
    null,
    f016.sourceBranch,
    f016.publicationStatus,
    f016.publication.peerReviewStatus,
    f016.publication.statusConfidence,
    f016.publication.recordType,
    f016.publication.publisher,
    f016.publication.venue,
    f016.publication.publicationDate,
    null,
    null,
    f016.publication.pageOrArticle,
    f016.publication.authoritativeUrl,
    f016.publication.verificationSource,
    f016.publication.dateChecked,
    null,
    f016.laterVersionStatus,
    f016.publication.potentialSuccessorCount,
    f016.publication.notes,
  ]];

  for (const [sheet, tableName, tableRange] of [
    [clusterSheet, "ClustercodesTable", "A4:T206"],
    [multiSheet, "MultilabelcodesTable", "A4:E2836"],
    [anchorSheet, "EvidenceanchorsTable", "A4:F206"],
    [publicationSheet, "PublicationstatusTable", "A4:Y207"],
  ]) {
    const oldTable = sheet.tables.getItem(tableName);
    const style = oldTable.style;
    oldTable.delete();
    const newTable = sheet.tables.add(tableRange, true, tableName);
    newTable.style = style;
  }

  clusterSheet.getRange("A3").values = [[
    "Baseline canonical SHA-256: 31655db590049900d965f1a9b97c9d4554c864ccf77eba3a56c1524801fce7ce. P311–P319 are the nine DOI-amendment integration rows; the frozen 193-cluster source is preserved.",
  ]];
  publicationSheet.getRange("A3").values = [[
    "Baseline reports were checked 2026-11-10. DOI-amendment rows P311–P319 use the authoritative-source checks recorded on 2026-11-14 and 2026-11-16.",
  ]];

  const clusterRows = clusterSheet.getRange("A5:T206").values;
  const multilabelRows = multiSheet.getRange("A5:E2836").values;
  validateUniqueAssignments(multilabelRows);
  const sensitivityResult = updateSensitivity(workbook, clusterRows, multilabelRows);

  const readme = workbook.worksheets.getItem("ReadMe");
  readme.getRange("A1").values = [[
    "Supplementary File S2: Evidence Matrix and Codebook (DOI protocol amendment)",
  ]];
  readme.getRange("A2").values = [[
    "This workbook integrates the nine reports restored after removal of the item-level DOI eligibility criterion into the frozen 193-cluster baseline.",
  ]];
  readme.getRange("B5").values = [["2026-11-16"]];
  readme.getRange("B6").values = [[
    "202 study-contribution clusters representing 203 eligible reports",
  ]];
  readme.getRange("B8").values = [[
    "Supplementary_File_S1_Included_Study_Register_202.xlsx",
  ]];
  readme.getRange("B9").values = [[
    "Frozen 193-cluster S2 plus full-text extraction and coding for P311–P319",
  ]];
  readme.getRange("B10").values = [[
    "Frozen 194-report audit plus official-source checks for nine restored reports",
  ]];
  readme.getRange("B13").values = [["N = 202 clusters"]];
  readme.getRange("B14").values = [[
    "N = 170 clusters after excluding 32 secondary-synthesis contribution clusters",
  ]];
  readme.getRange("B15").values = [[202]];
  readme.getRange("B16").values = [[multilabelRows.length]];
  readme.getRange("B17").values = [[203]];
  readme.getRange("B18").values = [[
    "Sensitivity rows were recalculated from the normalized 202-cluster coding. The nine restored reports were checked against their full texts and current codebook boundaries.",
  ]];

  workbook.recalculate();
  const output = await SpreadsheetFile.exportXlsx(workbook);
  await output.save(s2Output);
  return sensitivityResult;
}

await fs.mkdir(outputDir, { recursive: true });
await buildS1();
const sensitivity = await buildS2();
await fs.writeFile(
  `${outputDir}/build_summary.json`,
  `${JSON.stringify({
    generated: "2026-11-16",
    reports: 203,
    clusters: 202,
    nonSecondary: 170,
    secondary: sensitivity.secondaryCount,
    f016ClusterId: f016.clusterId,
    f016MultilabelAssignments: normalizedMultilabelRows().length,
    outputs: [s1Output, s2Output],
  }, null, 2)}\n`,
  "utf8",
);

console.log(JSON.stringify({ s1Output, s2Output, sensitivity }, null, 2));
