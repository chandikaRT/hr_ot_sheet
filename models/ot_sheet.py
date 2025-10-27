from odoo import api, models, fields
import calendar

class HROTSheet(models.Model):
    _name = 'hr.ot.sheet'
    _description = 'OT Sheet Header'

    name = fields.Char(string='Reference', copy=False, index=True)
    month = fields.Selection([
        ('1', 'January'), ('2', 'February'), ('3', 'March'),
        ('4', 'April'),   ('5', 'May'),      ('6', 'June'),
        ('7', 'July'),    ('8', 'August'),   ('9', 'September'),
        ('10', 'October'),('11', 'November'),('12', 'December')
    ], required=True)
    year  = fields.Integer(required=True, default=fields.Date.today().year)
    line_ids = fields.One2many('hr.ot.sheet.line','sheet_id', string='OT Lines')
    state = fields.Selection([('draft','Draft'),('done','Done')], default='draft')
    
    @api.model_create_multi
    def create(self, vals_list):
        for vals in vals_list:
            if vals.get('name', 'New') == 'New':
                month = int(vals.get('month'))
                year  = int(vals.get('year'))
                seq = self.env['ir.sequence'].next_by_code('hr.ot.sheet') or '001'
                vals['name'] = f"{seq}/{month:02d}/{year}"
        return super().create(vals_list)
    
    def action_apply_to_payslips(self):
        """Create / update payslip inputs for every line in the sheet."""
        self.ensure_one()
        Payroll = self.env['hr.payslip']
        Input   = self.env['hr.payslip.input']
        InputType = self.env['hr.payslip.input.type']

        # cache input types once
        codes = ('OT_NORMAL', 'OT_HOLIDAY', 'LATE_DEDUCTION')
        input_types = {
            c: InputType.search([('code', '=', c)], limit=1) or
               InputType.create({'name': c.replace('_', ' ').title(), 'code': c})
            for c in codes
        }

        for line in self.line_ids:
            year  = int(self.year)
            month = int(self.month)
            last_day = calendar.monthrange(year, month)[1]
            date_from = date(year, month, 1)
            date_to   = date(year, month, last_day)

            slip = Payroll.search([
                ('employee_id', '=', line.employee_id.id),
                ('date_from', '<=', date_to),
                ('date_to', '>=', date_from),
            ], limit=1)

            if not slip:
                contract = self.env['hr.contract'].search([
                    ('employee_id', '=', line.employee_id.id),
                    ('state', 'in', ('open', 'close'))
                ], limit=1)
                if not contract:
                    continue
                struct = (contract.structure_type_id.default_struct_id or
                          self.env['hr.payroll.structure'].search([], limit=1))
                if not struct:
                    raise UserError(
                        _('No salary structure for employee %s') % line.employee_id.name)
                slip = Payroll.create({
                    'employee_id': contract.employee_id.id,
                    'contract_id': contract.id,
                    'struct_id': struct.id,
                    'date_from': date_from,
                    'date_to': date_to,
                    'name': self.env['hr.ot.import.wizard']._generate_payslip_name(
                        contract.employee_id, struct.id, date_from, date_to),
                })

            # upsert only non-zero values
            def upsert(code, amount):
                if not amount:
                    return
                Input.search([
                    ('payslip_id', '=', slip.id),
                    ('input_type_id', '=', input_types[code].id)
                ], limit=1).write({'amount': amount}) or Input.create({
                    'payslip_id': slip.id,
                    'input_type_id': input_types[code].id,
                    'amount': amount,
                })

            upsert('OT_NORMAL', line.ot_normal)
            upsert('OT_HOLIDAY', line.ot_holiday)
            upsert('LATE_DEDUCTION', -abs(line.late_deduction) if line.late_deduction else 0)

            line.applied = True

        self.state = 'done'
        return {'type': 'ir.actions.act_window_close'}

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
