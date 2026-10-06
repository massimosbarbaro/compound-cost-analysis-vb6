# Costi: cost analysis of compounds and finished products

*Analisi dei costi delle mescole e dei prodotti finiti*

**Visual Basic 6** · 2002 · version 1.0.0  
Author: **Massimo Sbarbaro** ([ORCID 0009-0006-8965-9013](https://orcid.org/0009-0006-8965-9013))

## Overview

Costi compares standard and actual (average) costs of the compounds and finished products of a plastic-film plant. For each compound it combines raw-material cost, scrap and recovery percentages, machine hourly cost and output in kg per hour, and computes the cost per kilogram, also for products laminated with PE or PET. Cost data are imported from the MFG/PRO ERP.

I designed and programmed this application in 2002. It is published here in 2026 as part of the documented record of my software work, with a permanent DOI on Zenodo.

## Main functions

- Compound cost analysis: standard versus current cost, scrap %, recovery %, €/h, kg/h, €/kg (`frmCostiS`, `frmATC`).
- Average costs by thickness range and lamination type.
- Import of standard and current costs from ERP text exports.
- Reports produced in Microsoft Access through automation (`frmGrafici`).

## Data

Microsoft Access database (`Costi.mdb`) through DAO. Main tables: `MescoleMaxCostoCRR`, `CostoMa514`, `Scarti`, `MFG`.

## Technology

DAO 3.51, Microsoft Access 8 object library, Crystal Reports, Winsock, Common Controls, Tab control, Masked Edit.

## Repository contents

| Path | Content |
|---|---|
| `src/` | Visual Basic 6 project (`.vbp`), forms (`.frm` with their binary resources `.frx`) and modules (`.bas`), as listed in the project file. |
| `config-example/` | Templates of the `.ini` configuration files read at start-up, with placeholder values. |

## What is not included

Crystal Reports layouts (`.rpt`), compiled executables, installers, scripts for the host connection and the production databases are **not** included, because they contain operational data of the organisations for which the program was built. Configuration files are published as templates in `config-example/` with placeholder values: server addresses, accounts and passwords have been removed. Names of client organisations and products have been removed from captions and comments; local paths have been normalised.

## Related repositories

- [compound-batch-dosing-vb6](https://github.com/massimosbarbaro/compound-batch-dosing-vb6)

## How to cite

Use the citation metadata in [`CITATION.cff`](CITATION.cff) (GitHub: *Cite this repository*). Each release is archived on Zenodo with its own DOI.

> Sbarbaro, Massimo. *Costi: cost analysis of compounds and finished products (Visual Basic 6, 2002)*. Software, version 1.0.0. GitHub: https://github.com/massimosbarbaro/compound-cost-analysis-vb6

## License

Released under the [MIT License](LICENSE). © 2002 Massimo Sbarbaro.
