{
    'name': 'HR OT Sheet - Bulk OT Import',
    'version': '1.0.0',
    'summary': 'Import overtime (normal + holiday) and late deductions into payslips',
    'category': 'Human Resources/Payroll',
    'author': 'Chandika Rathnayake',
    'license': 'AGPL-3',
    'depends': ['hr_payroll', 'hr'],
    'data': [
        'security/ir.model.access.csv',
        'views/hr_ot_sheet_views.xml',
        'views/import_wizard_views.xml',
        'data/ot_input_codes.xml',
    ],
    'installable': True,
    'application': False,
}