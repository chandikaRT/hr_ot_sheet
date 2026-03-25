from odoo import fields, models


class HrEmployee(models.Model):
    _inherit = 'hr.employee'

    x_studio_ot_rate = fields.Float(
        string='OT Rate',
        digits=(10, 2),
        help='Hourly rate used to calculate normal and holiday overtime amounts.',
    )
    x_studio_employee_rate = fields.Float(
        string='Employee Rate',
        digits=(10, 2),
        help='Hourly rate used to calculate late deduction amounts.',
    )
