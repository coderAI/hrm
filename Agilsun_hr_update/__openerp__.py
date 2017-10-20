# -*- coding: utf-8 -*-
{
    'name': 'Agilsun Human Resources Update Module',
    'version': '1.0',
    'author': 'Agilsun',
    'category': 'Agilsun',
    'website': 'http://agilsun.com',
    'description': """
        Agilsun Human Resources Update Module: This module custom module hr
    """,
    'author': 'Agilsun',
    'website': 'http://www.openerp.com',
    'images': [
    ],
    "depends": [
        "Agilsun_hr",
        "report_xlsx",
        "base",
    ],
    'data': [
        'reports/define_reports.xml',
        "wizard/hr_employee_update_view.xml",
    ],
    'demo': [],
    'test': [
    ],
    'installable': True,
    'application': False,
    'auto_install': False,
    'css': [],
}
