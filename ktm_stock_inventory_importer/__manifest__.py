{
    'name': 'Stock Inventory Validator',
    'summary': 'Validate a physical inventory Excel file before loading it',
    'version': '18.0.1.0.1',
    'category': 'Inventory/Inventory',
    'author': 'Kaitim',
    'maintainers': ['jesusmaherrera'],
    'license': 'LGPL-3',
    'depends': ['stock'],
    'data': [
        'security/ir.model.access.csv',
        'wizard/stock_inventory_import_wizard_views.xml',
    ],
    'images': ['static/description/banner.png'],
}
