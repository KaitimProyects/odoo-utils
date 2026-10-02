import base64
from io import BytesIO

from openpyxl import Workbook

from odoo.exceptions import UserError
from odoo.tests import TransactionCase, tagged
from odoo.tools.safe_eval import safe_eval

HEADERS = [
    'Name', 'Internal Reference', 'Location', 'Lot/Serial', 'Counted Quantity',
]


@tagged('post_install', '-at_install')
class TestInventoryImportWizard(TransactionCase):

    @classmethod
    def setUpClass(cls):
        super().setUpClass()
        Product = cls.env['product.product']
        Location = cls.env['stock.location']
        cls.product = Product.create({
            'name': 'RF Widget', 'default_code': 'RF-001',
            'is_storable': True,
        })
        cls.tracked = Product.create({
            'name': 'RF Tracked', 'default_code': 'RF-LOT',
            'is_storable': True, 'tracking': 'lot',
        })
        cls.consumable = Product.create({
            'name': 'RF Consumable', 'default_code': 'RF-CONS',
            'is_storable': False,
        })
        cls.twin_a = Product.create({
            'name': 'RF Twin A', 'default_code': 'RF-TWIN',
            'is_storable': True,
        })
        cls.twin_b = Product.create({
            'name': 'RF Twin B', 'default_code': 'RF-TWIN-B',
            'is_storable': True,
        })
        # Some modules forbid a shared code at the ORM level; legacy data can
        # still hold it, so force it with SQL.
        cls.env.cr.execute(
            "UPDATE product_product SET default_code = 'RF-TWIN' WHERE id = %s",
            [cls.twin_b.id],
        )
        cls.twin_b.invalidate_recordset(['default_code'])
        cls.numeric = Product.create({
            'name': 'RF Numeric', 'default_code': '1234',
            'is_storable': True,
        })
        cls.shelf = Location.create({
            'name': 'RF Shelf A', 'usage': 'internal',
        })
        cls.shelf_dup_1 = Location.create({
            'name': 'RF Twin Shelf', 'usage': 'internal',
            'location_id': cls.shelf.id,
        })
        cls.shelf_b = Location.create({
            'name': 'RF Shelf B', 'usage': 'internal',
        })
        cls.shelf_dup_2 = Location.create({
            'name': 'RF Twin Shelf', 'usage': 'internal',
            'location_id': cls.shelf_b.id,
        })
        cls.customer_loc = Location.create({
            'name': 'RF Customer Place', 'usage': 'customer',
        })

    # -- helpers -----------------------------------------------------------

    def _xlsx(self, rows, headers=HEADERS):
        workbook = Workbook()
        sheet = workbook.active
        if headers is not None:
            sheet.append(headers)
        for row in rows:
            sheet.append(row)
        buffer = BytesIO()
        workbook.save(buffer)
        return buffer.getvalue()

    def _validate(self, rows, headers=HEADERS):
        wizard = self.env['ktm.stock.inventory.importer.wizard'].create({
            'file': base64.b64encode(self._xlsx(rows, headers)),
            'filename': 'inventory.xlsx',
        })
        action = wizard.action_validate()
        return wizard, action

    def _messages(self, wizard, row=None):
        errors = wizard.error_ids
        if row is not None:
            errors = errors.filtered(lambda e: e.row_number == row)
        return errors.mapped('message')

    # -- tests -------------------------------------------------------------

    def test_valid_file(self):
        wizard, action = self._validate([
            ['RF Widget', 'RF-001', 'RF Shelf A', None, 5],
            ['RF Tracked', 'RF-LOT', 'RF Shelf A', 'LOT1', 2.5],
        ])
        self.assertFalse(wizard.error_ids)
        self.assertEqual(wizard.state, 'valid')
        self.assertEqual(wizard.rows_checked, 2)
        self.assertTrue(wizard.result_message)
        self.assertEqual(action['res_model'], wizard._name)
        self.assertEqual(action['res_id'], wizard.id)
        self.assertEqual(action['target'], 'new')

    def test_headers_aliases_spanish_and_case(self):
        wizard, _action = self._validate(
            [['RF Widget', 'RF-001', 'RF Shelf A', None, 1]],
            headers=[
                ' nombre ', 'REFERENCIA INTERNA', 'Ubicacion',
                'Lote/Serie', 'cantidad contada',
            ],
        )
        self.assertFalse(wizard.error_ids, self._messages(wizard))
        self.assertEqual(wizard.state, 'valid')

    def test_missing_header(self):
        wizard, _action = self._validate(
            [['RF Widget', 'RF-001', 'RF Shelf A', None, 1]],
            headers=['Name', 'Internal Reference', 'Location', 'Lot/Serial'],
        )
        self.assertEqual(wizard.state, 'error')
        self.assertEqual(len(wizard.error_ids), 1)
        self.assertIn('Counted Quantity', wizard.error_ids.message)

    def test_unreadable_file(self):
        wizard = self.env['ktm.stock.inventory.importer.wizard'].create({
            'file': base64.b64encode(b'this is not an xlsx'),
            'filename': 'broken.xlsx',
        })
        with self.assertRaises(UserError):
            wizard.action_validate()

    def test_no_file(self):
        wizard = self.env['ktm.stock.inventory.importer.wizard'].create({})
        with self.assertRaises(UserError):
            wizard.action_validate()

    def test_empty_rows_are_skipped(self):
        wizard, _action = self._validate([
            [None, None, None, None, None],
            ['RF Widget', 'RF-001', 'RF Shelf A', None, 1],
        ])
        self.assertFalse(wizard.error_ids)
        self.assertEqual(wizard.rows_checked, 1)

    def test_product_not_found(self):
        wizard, _action = self._validate([
            ['Nope', 'NOPE-1', 'RF Shelf A', None, 1],
            ['RF Widget', 'WRONG-CODE', 'RF Shelf A', None, 1],
        ])
        self.assertEqual(wizard.state, 'error')
        self.assertTrue(any('not found' in m for m in self._messages(wizard, 2)))
        self.assertTrue(any('not found' in m for m in self._messages(wizard, 3)))
        error = wizard.error_ids.filtered(lambda e: e.row_number == 2)[0]
        self.assertEqual(error.product_code, 'NOPE-1')
        self.assertEqual(error.product_name, 'Nope')

    def test_product_name_or_code_empty(self):
        wizard, _action = self._validate([
            [None, 'RF-001', 'RF Shelf A', None, 1],
            ['RF Widget', None, 'RF Shelf A', None, 1],
        ])
        self.assertTrue(self._messages(wizard, 2))
        self.assertTrue(self._messages(wizard, 3))

    def test_product_case_insensitive_name(self):
        wizard, _action = self._validate([
            ['rf widget', 'RF-001', 'RF Shelf A', None, 1],
        ])
        self.assertFalse(wizard.error_ids)

    def test_same_code_two_products_resolved_by_name(self):
        wizard, _action = self._validate([
            ['RF Twin A', 'RF-TWIN', 'RF Shelf A', None, 1],
            ['RF Twin B', 'RF-TWIN', 'RF Shelf A', None, 1],
        ])
        self.assertFalse(wizard.error_ids, self._messages(wizard))

    def test_product_ambiguous(self):
        clone = self.env['product.product'].create({
            'name': 'RF Widget', 'default_code': 'RF-001-CLONE',
            'is_storable': True,
        })
        self.env.cr.execute(
            "UPDATE product_product SET default_code = 'RF-001' WHERE id = %s",
            [clone.id],
        )
        clone.invalidate_recordset(['default_code'])
        wizard, _action = self._validate([
            ['RF Widget', 'RF-001', 'RF Shelf A', None, 1],
        ])
        self.assertTrue(any('ambiguous' in m for m in self._messages(wizard, 2)))

    def test_float_code_is_normalized(self):
        wizard, _action = self._validate([
            ['RF Numeric', 1234.0, 'RF Shelf A', None, 1],
        ])
        self.assertFalse(wizard.error_ids, self._messages(wizard))

    def test_product_not_storable(self):
        wizard, _action = self._validate([
            ['RF Consumable', 'RF-CONS', 'RF Shelf A', None, 1],
        ])
        self.assertTrue(
            any('not storable' in m for m in self._messages(wizard, 2)))

    def test_tracked_product_requires_lot(self):
        wizard, _action = self._validate([
            ['RF Tracked', 'RF-LOT', 'RF Shelf A', None, 1],
        ])
        self.assertTrue(
            any('lot/serial' in m.lower() for m in self._messages(wizard, 2)))

    def test_wizard_opens_extra_large(self):
        # The error list has many columns; it is unreadable in a narrow dialog.
        menu_action = self.env.ref(
            'ktm_stock_inventory_importer.'
            'ktm_stock_inventory_importer_wizard_action')
        self.assertIn("'dialog_size': 'extra-large'", menu_action.context)
        _wizard, action = self._validate([
            ['RF Widget', 'RF-001', 'RF Shelf A', None, 1],
        ])
        self.assertEqual(action['context'].get('dialog_size'), 'extra-large')

    def test_tracked_product_zero_count_needs_no_lot(self):
        # Counting zero means nothing is there, so there is no lot to name.
        wizard, _action = self._validate([
            ['RF Tracked', 'RF-LOT', 'RF Shelf A', None, 0],
            ['RF Tracked', 'RF-LOT', 'RF Shelf B', None, '0'],
        ])
        self.assertFalse(self._messages(wizard, 2))
        self.assertFalse(self._messages(wizard, 3))
        wizard, _action = self._validate([
            ['RF Tracked', 'RF-LOT', 'RF Shelf A', None, 0.5],
        ])
        self.assertTrue(
            any('lot/serial' in m.lower() for m in self._messages(wizard, 2)))

    def test_location_not_found(self):
        wizard, _action = self._validate([
            ['RF Widget', 'RF-001', 'Missing Place', None, 1],
            ['RF Widget', 'RF-001', None, None, 1],
        ])
        self.assertTrue(any('not found' in m for m in self._messages(wizard, 2)))
        self.assertTrue(self._messages(wizard, 3))

    def test_location_by_complete_name_and_name_fallback(self):
        wizard, _action = self._validate([
            ['RF Widget', 'RF-001', 'rf shelf a/rf twin shelf', None, 1],
            ['RF Widget', 'RF-001', 'RF SHELF A', None, 2],
        ])
        self.assertFalse(wizard.error_ids, self._messages(wizard))

    def test_location_ambiguous(self):
        wizard, _action = self._validate([
            ['RF Widget', 'RF-001', 'RF Twin Shelf', None, 1],
        ])
        self.assertTrue(any('ambiguous' in m for m in self._messages(wizard, 2)))

    def test_location_not_internal(self):
        wizard, _action = self._validate([
            ['RF Widget', 'RF-001', 'RF Customer Place', None, 1],
        ])
        self.assertTrue(
            any('not an internal' in m for m in self._messages(wizard, 2)))

    def test_quantity_invalid(self):
        wizard, _action = self._validate([
            ['RF Widget', 'RF-001', 'RF Shelf A', None, None],
            ['RF Widget', 'RF-001', 'RF Shelf A', 'A', 'abc'],
            ['RF Widget', 'RF-001', 'RF Shelf A', 'B', -3],
            ['RF Widget', 'RF-001', 'RF Shelf A', 'C', 0],
            ['RF Widget', 'RF-001', 'RF Shelf A', 'D', '7.5'],
            ['RF Widget', 'RF-001', 'RF Shelf A', 'E', '1,000'],
        ])
        self.assertTrue(self._messages(wizard, 2))
        self.assertTrue(self._messages(wizard, 3))
        self.assertTrue(self._messages(wizard, 4))
        self.assertFalse(self._messages(wizard, 5), 'zero is a valid count')
        self.assertFalse(self._messages(wizard, 6), 'numeric text is valid')
        # A comma is a thousands separator in Mexico and a decimal one
        # elsewhere: guessing could silently turn 1,000 into 1.
        self.assertTrue(self._messages(wizard, 7), 'comma text is ambiguous')

    def test_duplicates_reported_with_rows(self):
        wizard, _action = self._validate([
            ['RF Widget', 'RF-001', 'RF Shelf A', None, 1],
            ['RF Widget', 'RF-001', 'RF Shelf A', None, 2],
            ['RF Widget', 'RF-001', 'RF Shelf A', None, 3],
            ['RF Widget', 'RF-001', 'RF Shelf A', 'X', 3],
        ])
        self.assertFalse(self._messages(wizard, 2))
        dup_3 = self._messages(wizard, 3)
        dup_4 = self._messages(wizard, 4)
        self.assertEqual(len(dup_3), 1)
        self.assertIn('2', dup_3[0])
        self.assertEqual(len(dup_4), 1)
        self.assertIn('2', dup_4[0])
        self.assertIn('3', dup_4[0])
        self.assertFalse(self._messages(wizard, 5))

    def test_all_errors_reported_and_revalidation_resets(self):
        wizard, _action = self._validate([
            ['Nope', 'N', 'Missing Place', None, -1],
            ['RF Consumable', 'RF-CONS', 'RF Customer Place', None, 'x'],
        ])
        self.assertGreaterEqual(len(wizard.error_ids), 4)
        first_count = len(wizard.error_ids)
        wizard.action_validate()
        self.assertEqual(len(wizard.error_ids), first_count)

    def test_validation_does_not_touch_stock(self):
        quants = self.env['stock.quant'].search_count([])
        moves = self.env['stock.move'].search_count([])
        self._validate([['RF Widget', 'RF-001', 'RF Shelf A', None, 5]])
        self.assertEqual(self.env['stock.quant'].search_count([]), quants)
        self.assertEqual(self.env['stock.move'].search_count([]), moves)

    # -- grouped errors in a new browser tab ---------------------------------

    def _error_action(self):
        return self.env.ref(
            'ktm_stock_inventory_importer.'
            'ktm_stock_inventory_importer_error_action')

    def test_open_grouped_errors_returns_url_to_new_tab(self):
        wizard, _action = self._validate([
            ['Nope', 'N1', 'RF Shelf A', None, 1],
        ])
        result = wizard.action_open_grouped_errors()
        action = self._error_action()
        self.assertEqual(result['type'], 'ir.actions.act_url')
        self.assertEqual(result['target'], 'new')
        self.assertEqual(
            result['url'],
            '/odoo/%s/action-ktm_stock_inventory_importer.'
            'ktm_stock_inventory_importer_error_action' % wizard.id)
        self.assertEqual(action.res_model, wizard.error_ids._name)

    def test_error_action_domain_is_scoped_to_the_wizard(self):
        wizard, _action = self._validate([
            ['Nope', 'N1', 'RF Shelf A', None, 1],
            ['Nope', 'N2', 'RF Shelf A', None, 1],
        ])
        other, _action = self._validate([
            ['Other', 'O1', 'RF Shelf A', None, 1],
        ])
        domain = safe_eval(
            self._error_action().domain, {'active_id': wizard.id})
        found = self.env['ktm.stock.inventory.importer.error'].search(domain)
        self.assertEqual(found, wizard.error_ids)
        self.assertFalse(found & other.error_ids)

    def test_error_action_groups_by_message_by_default(self):
        action = self._error_action()
        self.assertIn('list', action.view_mode)
        self.assertTrue(
            safe_eval(action.context)['search_default_group_by_message'])
        search_view = self.env['ktm.stock.inventory.importer.error'] \
            .get_view(action.search_view_id.id, 'search')
        self.assertIn('name="group_by_message"', search_view['arch'])
        wizard, _action = self._validate([
            ['Nope', 'N1', 'RF Shelf A', None, 1],
            ['RF Widget', 'RF-001', 'Missing Place', None, 1],
            ['Nope', 'N2', 'RF Shelf A', None, 1],
            ['Nope', 'N3', 'RF Shelf A', None, 1],
        ])
        domain = safe_eval(action.domain, {'active_id': wizard.id})
        groups = self.env['ktm.stock.inventory.importer.error'].read_group(
            domain, ['message'], ['message'])
        counts = {g['message']: g['message_count'] for g in groups}
        self.assertEqual(len(counts), 2)
        self.assertEqual(
            sorted(counts.values(), reverse=True), [3, 1])

    def test_row_number_is_not_summed_in_groups(self):
        # A summed Row column is meaningless, and as the first aggregated
        # column it squeezes the group name to a few letters in the list.
        Error = self.env['ktm.stock.inventory.importer.error']
        self.assertFalse(Error.fields_get(['row_number'])['row_number'].get('aggregator'))

    def test_error_list_is_read_only(self):
        arch = self.env['ktm.stock.inventory.importer.error'].get_view(
            view_type='list')['arch']
        for attribute in ('create="0"', 'edit="0"', 'delete="0"'):
            self.assertIn(attribute, arch)
