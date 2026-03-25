# -*- coding: utf-8 -*-
import base64
import io
import calendar
from datetime import date
from odoo import api, fields, models, _
from odoo.exceptions import UserError

try:
    from openpyxl import load_workbook
except ImportError:
    load_workbook = None

# ---------------------------------------------------------------------------
# Header-based column mapping
# Each canonical key maps to a list of accepted header aliases (lowercase).
# If no matching header is found in row 1, the positional default is used.
# ---------------------------------------------------------------------------
_HEADER_ALIASES = {
    'emp_code':       ['emp_code', 'employee code', 'barcode', 'emp code', 'code'],
    'emp_name':       ['emp_name', 'employee name', 'name', 'emp name'],
    'ot_normal_hrs':  ['ot_normal_hrs', 'normal ot', 'ot normal', 'normal ot hours',
                       'ot normal hours', 'normal overtime hours', 'normal overtime'],
    'ot_holiday_hrs': ['ot_holiday_hrs', 'holiday ot', 'ot holiday', 'holiday ot hours',
                       'ot holiday hours', 'holiday overtime hours', 'holiday overtime'],
    'late_ded_hrs':   ['late_ded_hrs', 'late deduction', 'late hours', 'late ded',
                       'late deduction hours'],
    'description':    ['description', 'desc', 'notes', 'remarks'],
}

_POSITIONAL_DEFAULTS = {
    'emp_code': 0,
    'emp_name': 1,
    'ot_normal_hrs': 2,
    'ot_holiday_hrs': 3,
    'late_ded_hrs': 4,
    'description': 5,
}


def _build_col_map(header_row):
    """Return a dict mapping canonical field names to column indices.

    Matches header cells against known aliases first; falls back to positional
    defaults for any column that cannot be matched.
    """
    normalised = [str(c).strip().lower() if c else '' for c in header_row]
    col_map = {}
    for canonical, aliases in _HEADER_ALIASES.items():
        for i, cell in enumerate(normalised):
            if cell in aliases:
                col_map[canonical] = i
                break
    for canonical, default_idx in _POSITIONAL_DEFAULTS.items():
        col_map.setdefault(canonical, default_idx)
    return col_map


class HrOtSheet(models.Model):
    _name = 'hr.ot.sheet'
    _description = 'OT Sheet Header'

    # ------------------------------------------------------------------
    # header fields
    # ------------------------------------------------------------------
    name = fields.Char(string='Reference', copy=False, index=True)
    month = fields.Selection([
        ('1', 'January'), ('2', 'February'), ('3', 'March'),
        ('4', 'April'), ('5', 'May'), ('6', 'June'),
        ('7', 'July'), ('8', 'August'), ('9', 'September'),
        ('10', 'October'), ('11', 'November'), ('12', 'December')
    ], required=True)
    year = fields.Integer(required=True, default=lambda _: date.today().year)
    line_ids = fields.One2many('hr.ot.sheet.line', 'sheet_id', string='OT Lines')
    state = fields.Selection([('draft', 'Draft'), ('done', 'Done')], default='draft')

    import_file = fields.Binary(string='OT Excel file')
    import_filename = fields.Char()

    # ------------------------------------------------------------------
    # sequence for name 001/10/2025
    # ------------------------------------------------------------------
    @api.model_create_multi
    def create(self, vals_list):
        for vals in vals_list:
            if not vals.get('name'):
                month = int(vals.get('month'))
                year = int(vals.get('year'))
                seq = self.env['ir.sequence'].next_by_code('hr.ot.sheet') or '001'
                vals['name'] = f"{seq}/{month:02d}/{year}"
        return super().create(vals_list)

    # ------------------------------------------------------------------
    # helper: build payslip name matching Odoo standard format
    # ------------------------------------------------------------------
    def _generate_payslip_name(self, employee, struct_id, date_from, date_to):
        struct = self.env['hr.payroll.structure'].browse(struct_id)
        if struct and struct.payslip_name:
            return struct.payslip_name
        return _('Payslip %s (%s - %s)') % (
            employee.name or '',
            date_from.strftime('%B %Y'),
            date_to.strftime('%B %Y'),
        )

    # ------------------------------------------------------------------
    # import Excel button
    # ------------------------------------------------------------------
    def action_import_excel(self):
        self.ensure_one()
        if not self.import_file:
            raise UserError(_('Please choose an Excel file first.'))
        if not load_workbook:
            raise UserError(_(
                'The Python library "openpyxl" is not installed on the server.\n'
                'Ask your system administrator to run: pip install openpyxl'
            ))

        data = base64.b64decode(self.import_file)
        ws = load_workbook(io.BytesIO(data), data_only=True).active
        rows = list(ws.iter_rows(min_row=1, values_only=True))
        if not rows:
            raise UserError(_('The uploaded file is empty.'))

        col_map = _build_col_map(rows[0])
        data_rows = rows[1:]  # row 0 is always treated as header

        error_lines, created = [], 0
        for idx, row in enumerate(data_rows, 2):
            row = list(row) + [None] * 10  # pad to avoid index errors

            emp_code    = row[col_map['emp_code']]
            emp_name    = row[col_map['emp_name']]
            description = row[col_map['description']]

            try:
                ot_normal_hrs  = float(row[col_map['ot_normal_hrs']]  or 0)
                ot_holiday_hrs = float(row[col_map['ot_holiday_hrs']] or 0)
                late_ded_hrs   = float(row[col_map['late_ded_hrs']]   or 0)
            except Exception:
                error_lines.append((idx, 'Invalid numeric value'))
                continue

            employee = None
            if emp_code:
                employee = self.env['hr.employee'].search(
                    [('barcode', '=', str(emp_code))], limit=1)
            if not employee and emp_name:
                employee = self.env['hr.employee'].search(
                    [('name', 'ilike', str(emp_name))], limit=1)
            if not employee:
                error_lines.append((idx, f'Employee not found: {emp_code}/{emp_name}'))
                continue

            self.env['hr.ot.sheet.line'].create({
                'sheet_id': self.id,
                'employee_id': employee.id,
                'ot_normal_hrs': ot_normal_hrs,
                'ot_holiday_hrs': ot_holiday_hrs,
                'late_ded_hrs': late_ded_hrs,
                'description': description or '',
            })
            created += 1

        self.import_file = False
        message = _('%d row(s) imported successfully.') % created
        if error_lines:
            message += '\n' + _('Skipped rows: ') + ', '.join(
                [f'Row {r}: {m}' for r, m in error_lines])

        return {
            'type': 'ir.actions.client',
            'tag': 'display_notification',
            'params': {
                'title': _('Import Complete'),
                'message': message,
                'type': 'warning' if error_lines else 'success',
                'sticky': bool(error_lines),
                'next': {
                    'type': 'ir.actions.act_window',
                    'res_model': self._name,
                    'res_id': self.id,
                    'views': [[False, 'form']],
                    'target': 'current',
                },
            },
        }

    # ------------------------------------------------------------------
    # apply to payslips button
    # ------------------------------------------------------------------
    def action_apply_to_payslips(self):
        self.ensure_one()
        Payroll   = self.env['hr.payslip']
        Input     = self.env['hr.payslip.input']
        InputType = self.env['hr.payslip.input.type']

        codes = ('OT_NORMAL', 'OT_HOLIDAY', 'LATE_DEDUCTION')
        input_types = {
            c: InputType.search([('code', '=', c)], limit=1) or
               InputType.create({'name': c.replace('_', ' ').title(), 'code': c})
            for c in codes
        }

        warnings = []

        for line in self.line_ids:
            year      = int(self.year)
            month     = int(self.month)
            last_day  = calendar.monthrange(year, month)[1]
            date_from = date(year, month, 1)
            date_to   = date(year, month, last_day)

            slip = Payroll.search([
                ('employee_id', '=', line.employee_id.id),
                ('date_from', '<=', date_to),
                ('date_to',   '>=', date_from),
            ], limit=1)

            # Skip payslips that are already validated or paid
            if slip and slip.state in ('done', 'paid'):
                warnings.append(_(
                    '%s: payslip is already validated/paid — skipped.'
                ) % line.employee_id.name)
                continue

            if not slip:
                contract = self.env['hr.contract'].search([
                    ('employee_id', '=', line.employee_id.id),
                    ('state', 'in', ('open', 'close'))
                ], limit=1)
                if not contract:
                    warnings.append(_(
                        '%s: no active contract found — skipped.'
                    ) % line.employee_id.name)
                    continue

                struct = (contract.structure_type_id.default_struct_id or
                          self.env['hr.payroll.structure'].search([], limit=1))
                if not struct:
                    warnings.append(_(
                        '%s: no salary structure found — skipped. '
                        'Assign a salary structure to the contract first.'
                    ) % line.employee_id.name)
                    continue

                slip = Payroll.create({
                    'employee_id': contract.employee_id.id,
                    'contract_id': contract.id,
                    'struct_id':   struct.id,
                    'date_from':   date_from,
                    'date_to':     date_to,
                    'name':        self._generate_payslip_name(
                                       contract.employee_id, struct.id,
                                       date_from, date_to),
                })

            def upsert(code, amount, desc=None):
                if not amount:
                    return
                existing = Input.search([
                    ('payslip_id',    '=', slip.id),
                    ('input_type_id', '=', input_types[code].id)
                ], limit=1)
                if existing:
                    existing.write({'amount': amount, 'name': desc or ''})
                else:
                    Input.create({
                        'payslip_id':    slip.id,
                        'input_type_id': input_types[code].id,
                        'amount':        amount,
                        'name':          desc or '',
                    })

            upsert('OT_NORMAL',      line.ot_normal,
                   line._desc_for_code('OT_NORMAL'))
            upsert('OT_HOLIDAY',     line.ot_holiday,
                   line._desc_for_code('OT_HOLIDAY'))
            upsert('LATE_DEDUCTION',
                   -abs(line.late_deduction) if line.late_deduction else 0,
                   line._desc_for_code('LATE_DEDUCTION'))

            line.applied = True

        self.state = 'done'

        if warnings:
            return {
                'type': 'ir.actions.client',
                'tag': 'display_notification',
                'params': {
                    'title': _('Applied with warnings'),
                    'message': '\n'.join(warnings),
                    'type': 'warning',
                    'sticky': True,
                },
            }
        return {'type': 'ir.actions.act_window_close'}


class HrOtSheetLine(models.Model):
    _name = 'hr.ot.sheet.line'
    _description = 'OT Sheet Line'

    sheet_id    = fields.Many2one('hr.ot.sheet', ondelete='cascade')
    employee_id = fields.Many2one('hr.employee', string='Employee', required=True)

    # hours entered by user
    ot_normal_hrs  = fields.Float(string='Normal OT Hours',      digits=(4, 2))
    ot_holiday_hrs = fields.Float(string='Holiday OT Hours',     digits=(4, 2))
    late_ded_hrs   = fields.Float(string='Late Deduction Hours', digits=(4, 2))

    # rates from employee master
    ot_rate       = fields.Float(related='employee_id.x_studio_ot_rate',
                                 string='OT Rate', readonly=True)
    employee_rate = fields.Float(related='employee_id.x_studio_employee_rate',
                                 string='Employee Rate', readonly=True)

    # computed amounts
    ot_normal      = fields.Monetary(string='Normal OT Amount',
                                     currency_field='company_currency',
                                     compute='_compute_amounts', store=True)
    ot_holiday     = fields.Monetary(string='Holiday OT Amount',
                                     currency_field='company_currency',
                                     compute='_compute_amounts', store=True)
    late_deduction = fields.Monetary(string='Late Deduction Amount',
                                     currency_field='company_currency',
                                     compute='_compute_amounts', store=True)

    company_currency = fields.Many2one('res.currency',
                                       related='employee_id.company_id.currency_id',
                                       readonly=True)
    applied     = fields.Boolean(default=False)
    description = fields.Char(string='Description',
                               compute='_compute_description', store=True)

    @api.depends('ot_normal_hrs', 'ot_holiday_hrs', 'late_ded_hrs',
                 'ot_rate', 'employee_rate')
    def _compute_amounts(self):
        for rec in self:
            rec.ot_normal      = rec.ot_normal_hrs  * rec.ot_rate
            rec.ot_holiday     = rec.ot_holiday_hrs * rec.ot_rate
            rec.late_deduction = rec.late_ded_hrs   * rec.employee_rate

    @api.depends('ot_normal_hrs', 'ot_holiday_hrs', 'late_ded_hrs',
                 'ot_rate', 'employee_rate')
    def _compute_description(self):
        for rec in self:
            desc = []
            if rec.ot_normal_hrs:
                desc.append(f"{rec.ot_normal_hrs} OT hours @ {rec.ot_rate:.2f}")
            if rec.ot_holiday_hrs:
                desc.append(f"{rec.ot_holiday_hrs} Holiday OT hours @ {rec.ot_rate:.2f}")
            if rec.late_ded_hrs:
                desc.append(f"{rec.late_ded_hrs} Late hours @ {rec.employee_rate:.2f}")
            rec.description = ' | '.join(desc) if desc else ''

    def _desc_for_code(self, code):
        self.ensure_one()
        if code == 'OT_NORMAL' and self.ot_normal_hrs:
            return f"{self.ot_normal_hrs} OT hours @ {self.ot_rate:.2f}"
        if code == 'OT_HOLIDAY' and self.ot_holiday_hrs:
            return f"{self.ot_holiday_hrs} OT hours @ {self.ot_rate:.2f}"
        if code == 'LATE_DEDUCTION' and self.late_ded_hrs:
            return f"{self.late_ded_hrs} Late hours @ {self.employee_rate:.2f}"
        return ''
