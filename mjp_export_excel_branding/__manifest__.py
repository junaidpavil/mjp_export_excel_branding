# -*- coding: utf-8 -*-
{
    'name': 'Excel Export with Company Branding',
    'version': '1.0.0',
    'summary': 'Adds company logo and name to Excel exports across all models',
    'description': """
        Excel Export with Company Branding
        ==================================
        
        This module enhances the default Odoo export functionality by adding
        company branding elements to Excel exports.
        
        Features:
        ---------
        - Adds company logo at the top of exported Excel files
        - Displays company name as header
        - Works across all models (Sales, Invoices, Purchases, etc.)
        - Improves professionalism of exported documents
        
        Use Case:
        ---------
        By default, Odoo exports only selected fields as plain data.
        This module ensures exported Excel files include company identity,
        making them suitable for sharing with clients and stakeholders.
    """,
    'author': 'Muhammed Junaid P',
    'category': 'Tools',
    'license': 'LGPL-3',
    'depends': ['base','base_setup'],
    'data': [
        'views/res_config_view.xml',
    ],
    'images': ['static/description/banner.png'],
    'installable': True,
    'application': False,
    'auto_install': False,
}