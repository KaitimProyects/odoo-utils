# Stock Inventory Validator

Validate a physical inventory Excel file before loading it.

## What it adds

Only transient models, so there is no table to migrate or clean up:

- `ktm.stock.inventory.importer.wizard`: holds the uploaded file (`file`, `filename`), the result (`state` draft/error/valid, `rows_checked`, `result_message`) and the `error_ids` One2many.
- `ktm.stock.inventory.importer.error`: one record per problem, with the Excel `row_number`, the column (`field_name`), the row's name, internal reference, location and lot, and the `message`.
- A form view opened as a dialog. After pressing Validate the same wizard reopens and lists every error. The "View Grouped Errors" button (shown only when there are errors) opens the errors in a new browser tab as a native list grouped by message, with collapsible groups and counts, so the wizard stays open. The tab URL is `/odoo/<wizard id>/action-ktm_stock_inventory_importer.ktm_stock_inventory_importer_error_action`; the action's domain reads the wizard id from `active_id`. The list is read only and can be regrouped by Column or filtered from the search box.
- A window action and a menu, `Inventory > Operations > Adjustments > Validate Inventory File`, restricted to `stock.group_stock_manager` (it sits next to Physical Inventory). Model access is granted to `stock.group_stock_manager` only, so a plain inventory user cannot run the wizard through RPC either.

The module only validates. It never writes to stock; a test asserts this.

## Why a separate module

The client needs the file checked before anyone decides how to load it. Keeping the check apart lets the later load step be a small module that depends on this one, without touching the validation.

## Dependencies

- `stock`: products, locations, `is_storable`, tracking and the Adjustments menu.

`openpyxl` reads the file. It is already an Odoo core requirement, so there is no `external_dependencies` entry and no `requirements.txt`.

## Expected file

First sheet, header in row 1, empty rows skipped. Headers are trimmed and compared case-insensitively; English and Spanish are accepted:

| Column | Accepted headers | Required |
|---|---|---|
| Name | `Name`, `Nombre` | yes |
| Internal reference | `Internal Reference`, `Referencia interna` | yes |
| Location | `Location`, `Ubicacion`, `Ubicación` | yes |
| Lot/Serial | `Lot/Serial`, `Lote/Serie` | no |
| Counted quantity | `Counted Quantity`, `Cantidad contada` | yes |

Errors reported per Excel row:

- duplicate line (same product, location and lot as an earlier row);
- product not found, or ambiguous: name and internal reference must match together;
- product not storable;
- lot/serial-tracked product without a lot, unless the counted quantity is 0 (nothing on hand, so no lot to name);
- location not found, not internal, or ambiguous;
- invalid quantity: empty, not a number, negative, or text containing a comma.

## Extension points

- `_parse_xlsx(content)` returns `(rows, errors)`. Override it to support another layout.
- `_validate_rows(rows)` takes parsed rows and returns error value dicts. It does not depend on the file, so a future load step can call it and then use the resolved records.
- Per-check methods to extend or override: `_resolve_products`, `_resolve_locations`, `_check_product_rules`, `_check_quantity`, `_check_duplicates`, plus `_company_domain` and the `HEADER_ALIASES` constant.

## Gotchas

- Some databases (e.g. with a module forbidding duplicate codes) reject a shared `default_code` at ORM level, so the tests that need two products with the same code insert them with SQL.
- A quantity written as text with a comma (`1,5`, `1,000`) is rejected on purpose: in Mexico the comma is a thousands separator, so it is ambiguous. Numeric cells or a period decimal are accepted.
- Archived products and locations are not found, because the searches use the default active filter.
- Locations match on `complete_name` first and fall back to `name`. Only internal locations are accepted; a unique name that exists only as a non-internal location reports "not an internal location".
- Records are searched for the wizard's company plus records with no company.

## License and category

LGPL-3, not the usual Kaitim OPL-1: the module is meant for a public utilities repository where others can use and contribute to it. The conventions lint reports this deviation on purpose; it has no exemption mechanism for a deliberate open licence.

The category is `Inventory/Inventory`, the same as core `stock`, not the usual `Kaitim/Inventory`: the module is meant for the Odoo Apps Store, which only browses its own fixed category list, so a `Kaitim` category would leave the app unlisted. The conventions lint reports this deviation on purpose too.
