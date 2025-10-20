from odoo import models, fields

class HROTSheet(models.Model):
    _name = 'hr.ot.sheet'
    _description = 'OT Sheet Header'

    name = fields.Char(string='Reference', default=lambda self: 'New')
    month = fields.Selection([(str(i), str(i)) for i in range(1,13)], string='Month', required=True)
    year = fields.Integer(string='Year', required=True)
    line_ids = fields.One2many('hr.ot.sheet.line','sheet_id', string='OT Lines')
    state = fields.Selection([('draft','Draft'),('done','Done')], default='draft')

class HROTSheetLine(models.Model):
    _name = 'hr.ot.sheet.line'
    _description = 'OT Sheet Line'

    sheet_id = fields.Many2one('hr.ot.sheet', ondelete='cascade')
    employee_id = fields.Many2one('hr.employee', string='Employee', required=True)
    ot_normal = fields.Monetary(string='Normal OT Amount', currency_field='company_currency')
    ot_holiday = fields.Monetary(string='Holiday/Sunday OT Amount', currency_field='company_currency')
    late_deduction = fields.Monetary(string='Late Deduction Amount', currency_field='company_currency')
    company_currency = fields.Many2one('res.currency', related='employee_id.company_id.currency_id', readonly=True)

    applied = fields.Boolean(default=False)
