# Supplementary files and scripts

This repository contains supplementary files for the bachelor project *A Systematic Literature Review on Characterizing and Measuring Human-AI Alignment in Hybrid Teams*.

## Supplementary files

- `Supplementary_File_S1_Included_Study_Register_205.xlsx` lists 206 reports grouped into 205 study-contribution clusters. One cluster contains two reports.
- `Supplementary_File_S2_Evidence_Matrix_205.xlsx` contains the coding, evidence locations, publication details, codebook, and sensitivity analysis for the 205 clusters.

These are the S1 and S2 files submitted with the thesis.

## Scripts

- `build_final_202.mjs` starts from two 201-cluster workbooks. They already contain P311–P318. The script adds P319 and updates the notes and sensitivity analysis.
- `build_final_205.mjs` adds P320–P322 from `inputs/p320_p322_delta.json` and updates the sensitivity analysis.
- `reconcile_final_workbooks.py` repeats text edits made after the 16 November 2026 build of the 205-cluster workbooks. It expands the short S1 label `Secondary synthesis`, corrects the P105 title in S1 and S2, and restores the full D4 definition in the S2 codebook. This script was written on 25 November because the earlier edits had not been recorded in a script.
- `validate_final_205.mjs` checks selected records and scans for formula errors.
- `compare_workbook_values.py` compares workbook cell values and formulas.
- `inspect_workbooks.mjs` lists workbook sheets and tables.

## Running the build

The two 201-cluster workbooks needed by `build_final_202.mjs` are included in `inputs/doi_eight_candidate_integration_2026-11-16/`. They are fixed input files based on the 193-cluster baseline with P311–P318 added. Their provenance and limits are described in `inputs/doi_eight_candidate_integration_2026-11-16/README.md`. The original generation script is not included in this repository; the documented build therefore starts from these fixed inputs.

Set `BP_INPUT_ROOT` to the `inputs/` folder. Set `BP_OUTPUT_ROOT` to a separate output folder. Then run the two build scripts, followed by `reconcile_final_workbooks.py`.

The build was rerun locally from the fixed 201-cluster workbooks onward. It was not rerun from the earlier 193-cluster data. Also, `p320_p322_delta.json` was copied from the final thesis workbooks. The match for P320–P322 is therefore expected and does not independently check those records.

The scripts that produced the 193-cluster baseline and the publication-status records, including the Crossref queries, are not part of this repository.

The scripts use Node.js and `@oai/artifact-tool` 2.8.188. This spreadsheet package is private and is not included here.

After text corrections, the generated workbooks matched the thesis S1 and S2 cell by cell, including formulas. The validation script found no formula errors.

## Use of generative AI

ChatGPT and Codex(GPT-5.6) helped write, revise, and run the scripts. The author reviewed the scripts and the results.
