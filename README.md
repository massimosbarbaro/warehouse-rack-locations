# Scaffali: rack location management for a finished-goods warehouse

*Gestione delle ubicazioni a scaffale del magazzino prodotti finiti*

**Visual Basic 6** · 2002–2003 · version 1.3.4  
Author: **Massimo Sbarbaro** ([ORCID 0009-0006-8965-9013](https://orcid.org/0009-0006-8965-9013))

## Overview

Scaffali manages where finished reels and pallets are stored in a racked warehouse. When a product is put away the operator assigns rack, row and column; the program then supports picking and moving stock between locations. Product and packing data are read from the packing database of the plant.

I designed and programmed this application in 2002–2003. It is published here in 2026 as part of the documented record of my software work, with a permanent DOI on Zenodo.

## Main functions

- Put-away of a production ID or order: rack, row, column, number of reels and pallet (`frmID`).
- Search, picking and moving of materials between locations.
- Load, unload and move menus with a separate configuration file (`SCAFFALI.ini`).
- Import of packing data from the plant's packing database (`frmImportMFG`).

## Data

Microsoft Access database (`Scaffali.mdb`) through DAO. Main tables: `Scaffali`, `Pallette`, `ODL`, `Legami`.

## Technology

DAO 3.51, Microsoft Access 8 object library, Common Controls, Tab control, Masked Edit, DBGrid.

## Repository contents

| Path | Content |
|---|---|
| `src/` | Visual Basic 6 project (`.vbp`), forms (`.frm` with their binary resources `.frx`) and modules (`.bas`), as listed in the project file. |
| `config-example/` | Templates of the `.ini` configuration files read at start-up, with placeholder values. |

## What is not included

Crystal Reports layouts (`.rpt`), compiled executables, installers, scripts for the host connection and the production databases are **not** included, because they contain operational data of the organisations for which the program was built. Configuration files are published as templates in `config-example/` with placeholder values: server addresses, accounts and passwords have been removed. Names of client organisations and products have been removed from captions and comments; local paths have been normalised.

## Related repositories

- [calender-production-orders-vb6](https://github.com/massimosbarbaro/calender-production-orders-vb6)
- [physical-inventory-reconciliation-vb6](https://github.com/massimosbarbaro/physical-inventory-reconciliation-vb6)

## How to cite

Use the citation metadata in [`CITATION.cff`](CITATION.cff) (GitHub: *Cite this repository*). Each release is archived on Zenodo with its own DOI.

> Sbarbaro, Massimo. *Scaffali: rack location management for a finished-goods warehouse (Visual Basic 6, 2002–2003)*. Software, version 1.3.4. GitHub: https://github.com/massimosbarbaro/warehouse-rack-locations-vb6

## License

Released under the [MIT License](LICENSE). © 2002 Massimo Sbarbaro.
