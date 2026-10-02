{
    'name': 'KTM Stock Inventory Importer',
    'summary': 'Validate a physical inventory Excel file before loading it',
    'version': '18.0.1.0.0',
    'category': 'Kaitim/Inventory',
    'author': 'Kaitim',
    'maintainers': ['jesusmaherrera'],
    'license': 'LGPL-3',
    'depends': ['stock'],
    'data': [
        'security/ir.model.access.csv',
        'wizard/stock_inventory_import_wizard_views.xml',
    ],
}
