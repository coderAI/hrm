# -*-coding:utf-8-*-


from openerp.osv import osv, fields
from datetime import datetime
from time import time
from openerp.addons.report_xlsx.utils import _render  # @UnresolvedImport
from openerp.report import report_sxw  # @UnresolvedImport
from openerp.addons.report_xlsx import report_xlsx_utils  # @UnresolvedImport
from openerp import SUPERUSER_ID, tools
import sys
import codecs
import base64
import datetime, xlrd
from xlrd.sheet import ctype_text

class update_employee(osv.osv_memory):
    _name = 'update.employee'


    _columns = {
        'title': fields.selection([('0', 'CMND'), ('1', 'Bank Account'), ('2', 'Education'), ('3', 'Employee'),
                                   ('4', 'Contract'), ('5', 'Working Record'), ('6', 'Contract Mass'), ('7', 'Working Record Mass')],
                                  string='Title', required=True),
        'file': fields.binary('File'),
        'filename': fields.char('File Name'),
        'list_emp': fields.html('List Employee', translate=True),
    }

    def bt_update_emp(self, cr, uid, ids, context=None):
        data_file = self.browse(cr, uid, ids)
        if data_file[0].file == False:
            raise osv.except_osv('Validation Error !', 'Dont have file excel to import')
        base_data = base64.decodestring(data_file[0].file)
        xl_workbook  = xlrd.open_workbook(file_contents = base_data)
        sheet_names = xl_workbook.sheet_names()
        xl_sheet = xl_workbook.sheet_by_name(sheet_names[0])
        if xl_sheet.cell(0, 1).value:
            if xl_sheet.cell(0, 1).value != "Document Type":
                raise osv.except_osv(('Error!'),
                                     ('This template wrong!'))
        persional_document_obj = self.pool.get('vhr.personal.document')
        vhr_personal_document_obj_field = ['document_type_id', 'number,issue_date', 'expiry_date', 'country_id',
                                           'city_id', 'district_id', 'is_received_hard_copy', 'active,status_id',
                                           'note', 'id_personal_document']
        vhr_personal_document_obj_field = ['number','issue_date', 'expiry_date', 'country_id',
                                           'city_id', 'district_id', 'is_received_hard_copy', 'active','status_id',
                                           'note', 'document_type_id']
        for row_idx in range(4, xl_sheet.nrows):    # Iterate through rows
            id_personal_document = xl_sheet.cell(row_idx, 0).value,
            document_type_id = xl_sheet.cell(row_idx, 2).value or '',
            number = xl_sheet.cell(row_idx, 3).value or '',
            issue_date = xl_sheet.cell(row_idx, 4).value or '',
            expiry_date = xl_sheet.cell(row_idx, 5).value or '',
            country_id = xl_sheet.cell(row_idx, 6).value or '',
            city_id =  xl_sheet.cell(row_idx, 8).value or '',
            district_id = xl_sheet.cell(row_idx, 10).value or '',
            is_received_hard_copy = xl_sheet.cell(row_idx, 12).value or '',
            active = xl_sheet.cell(row_idx, 13).value or '',
            status_id = xl_sheet.cell(row_idx, 14).value or '',
            note = xl_sheet.cell(row_idx, 16).value or '',
            persional_document_old = persional_document_obj.read(cr, SUPERUSER_ID,
                                                                 int(id_personal_document[0]),vhr_personal_document_obj_field)
            for tmp in document_type_id:
                if tmp:
                    document_type_id = int(tmp)
                else:
                    document_type_id = 'NULL'
            for tmp in number:
                number = tmp
            for tmp in issue_date:
                if tmp:
                    try:
                        issue_date = datetime.date.strptime(tmp, "%Y-%m-%d")
                    except:
                        issue_date = tmp
                else:
                    issue_date = False
            for tmp in expiry_date:
                if tmp:
                    expiry_date = tmp
                else:
                    expiry_date = 'NULL'
            for tmp in country_id:
                if tmp:
                    country_id = int(tmp)
                else:
                    country_id = 'NULL'
            for tmp in district_id:
                if tmp:
                    district_id = int(tmp)
                else:
                    district_id = 'NULL'
            for tmp in city_id:
                if tmp:
                    city_id = int(tmp)
                else:
                    city_id = 'NULL'
            for tmp in is_received_hard_copy:
                if tmp:
                    is_received_hard_copy = tmp
                else:
                    is_received_hard_copy =False
            for tmp in active:
                if tmp and tmp == 1:
                    active = 'TRUE'
                else:
                    active ='False'
            for tmp in status_id:
                if tmp:
                    status_id = int(tmp)
                else:
                    status_id = 'NULL'
            for tmp in note:
                if tmp:
                    note = tmp
                else:
                    note  = 'NULL'
            for tmp in id_personal_document:
                if tmp:
                    id_personal_document = int(tmp)
                else:
                    id_personal_document = 'NULL'
            vals_update={
                'document_type_id': int(xl_sheet.cell(row_idx, 2).value) or '',
                'number' : int(xl_sheet.cell(row_idx, 3).value) or '',
                'issue_date' : xl_sheet.cell(row_idx, 4).value or '',
                'expiry_date' : xl_sheet.cell(row_idx, 5).value or '',
                'country_id' : int(xl_sheet.cell(row_idx, 6).value) or '',
                'city_id' : int(xl_sheet.cell(row_idx, 8).value) or '',
                'district_id' : xl_sheet.cell(row_idx, 10).value and int(xl_sheet.cell(row_idx, 10).value) or '',
                'is_received_hard_copy' : xl_sheet.cell(row_idx, 12).value or '',
                'active': xl_sheet.cell(row_idx, 13).value or '',
                'status_id': xl_sheet.cell(row_idx, 14).value and int(xl_sheet.cell(row_idx, 14).value) or '',
                'note': xl_sheet.cell(row_idx, 16).value or '',
            }
            sql = '''
                UPDATE vhr_personal_document SET document_type_id = %s, 
                number = %s, 
                issue_date = '%s:00:00+00',
                expiry_date = %s,
                country_id = %s,
                city_id = %s,
                district_id = %s,
                is_received_hard_copy = %s,
                active = %s,
                status_id = %s,
                note = %s
                WHERE id= %s
            ''' % (document_type_id,number,issue_date,expiry_date,country_id,city_id,district_id,is_received_hard_copy,active,status_id,note,id_personal_document)
            cr.execute(sql)
            model_obj = self.pool.get('ir.model')
            ir_model_fields_obj = self.pool.get('ir.model.fields')
            model_id = model_obj.search(cr,SUPERUSER_ID,[('model','=',"vhr.personal.document")])
            audittrail_log_obj = self.pool.get('audittrail.log')
            audittrail_log_ids = audittrail_log_obj.search(cr,SUPERUSER_ID,[('object_id','=',model_id[0]),('res_id','=',id_personal_document)])
            if audittrail_log_ids:
                audittrail_log_id = audittrail_log_ids[0]
            else:
                audittrail_log_id = audittrail_log_obj.create(cr, SUPERUSER_ID, {'object_id': model_id[0],
                                                                                 'res_id': id_personal_document,
                                                                                 'method': 'write'
                                                                                 })
            audittrail_log_line_obj = self.pool.get('audittrail.log.line')
            for i in vhr_personal_document_obj_field:
                field_id = ir_model_fields_obj.search(cr, SUPERUSER_ID, [('model_id', '=', model_id[0]), ('name', '=', i)])
                vals_log = {
                    "res_id": id_personal_document,
                    "log_id": audittrail_log_id,
                    "field_id": field_id[0],
                    "old_value": persional_document_old.get(i),
                    "new_value": vals_update.get(i),
                    "old_value_text": persional_document_old.get(i),
                    "new_value_text": vals_update.get(i),
                    "field_description": 'Mass Update '+i,
                }
                audittrail_log_line_obj.create(cr,SUPERUSER_ID,vals_log)
        return
    def action_export(self, cr, uid, ids, context=None):
        if not isinstance(ids, list):
            ids = [ids]
        ######
        persional_document_obj = self.pool.get('vhr.personal.document')
        persional_document_type_obj = self.pool.get('vhr.personal.document.type')
        hr_obj = self.pool.get('hr.employee')
        lines = []
        for record in self.browse(cr, uid, ids):
            domain = []
            type = record.title or False
            list_emp = record.list_emp or False
            code_emp = []
            if list_emp:
                code = ''
                num = ['0', '1', '2', '3', '4', '5', '6', '7', '8', '9']
                for data in list_emp:
                    if data in num:
                        code += data
                    else:
                        if code != '':
                            code_emp.append(int(code))
                            code = ''
            else:
                raise osv.except_osv(('Error!'),
                                     ('You have not entered the employee code list'))
            if code_emp is None:
                raise osv.except_osv(('Error!'),
                                     ('You have not entered the employee code list'))
            emp_id = hr_obj.search(cr, uid, [('code', 'in', code_emp)])
            domain.append(('employee_id', 'in', emp_id))
            if type == '0':
                per_type_id = persional_document_type_obj.search(cr, uid, [('code', '=', 'PDT-005')])[0]
                domain.append(('document_type_id', '=', per_type_id))
                per_list = persional_document_obj.search(cr, uid, domain)
                if per_list:
                    for line in persional_document_obj.browse(cr, uid, per_list):
                        lines.append({
                            'emp_code': line.employee_id.code or '',
                            'document_id': line.id or '',
                            'document_type_id': line.document_type_id.id or '',
                            'number': line.number or '',
                            'issue_date': line.issue_date or '',
                            'expiry_date':line.expiry_date or '',
                            'country_id': line.country_id.id or '',
                            'country_name': line.country_id.name or '',
                            'city_id': line.city_id.id or '',
                            'city_name': line.city_id.name or '',
                            'district_id': line.district_id.id or '',
                            'district_name': line.district_id.name or '',
                            'status_id': line.status_id.id or  '',
                            'status_name': line.status_id.name or '',
                            'place': line.city_id.id or '',
                            'active': line.active or '',
                            'is_received_hard_copy':line.is_received_hard_copy or '',
                            'note': line.note or '',
                        })
            elif type == '6':
                contract_obj = self.pool.get('hr.contract')
                contract_id = contract_obj.search(cr, uid, domain)
                if contract_id:
                    for line in contract_obj.browse(cr, uid, contract_id):
                        lines.append({
                            'con_id': line.id or '',
                            'emp_code': line.employee_id.code or '',
                            'company': line.company_id.id or '',
                            'job_type': line.job_type_id.id or '',
                            'con_type': line.type_id.id or '',
                            'div': line.division_id.id or '',
                            'depart': line.department_id.id or '',
                        })
        # #####
        record = self.read(cr, uid, ids,
                           ['title', 'file', 'file_name'], context=context)

        datas = {
            'ids': ids,
            'model': context.get('active_model', 'update.employee'),
            'active_ids': context.get('active_ids'),
            'param': [],
            'form': record,
            'lines': lines
        }
        return {
            'type': 'ir.actions.report.xml',
            'report_name': 'update_employee',
            'datas': datas,
            'name': u'Update Employee'
        }

update_employee()


class rpt_update_emp_xlsx_parser(report_sxw.rml_parse):
    def __init__(self, cr, uid, name, context):
        super(rpt_update_emp_xlsx_parser, self).__init__(cr, uid, name, context=context)
        self.context = context
        self.localcontext.update({
            'datetime': datetime,
        })


class rpt_update_emp_xlsx(report_xlsx_utils.generic_report_xlsx_base):
    def __init__(self, name, table, rml=False, parser=False, header=True, store=False):
        super(rpt_update_emp_xlsx, self).__init__(name, table, rml, parser, header, store)

        self.xls_styles.update({
            'fontsize_350': 'font: height 360;'
        })

    def generate_xls_report(self, _p, _xs, data, objects, wb):

        #
        # ws = super(rpt_update_emp_xlsx, self).generate_xls_report(_p, _xs, data,
        #                                                                   objects, wb,
        #
        #
        #                                                               report_name)
        sheet1 = wb.add_worksheet(u'Information')
        format_head = wb.add_format(
            {'font_size': 14, 'align': 'left', 'bold': True, 'font_color': 'blue'})
        format_head1 = wb.add_format(
            {'font_size': 11, 'align': 'left', 'bold': True})
        format_title_02 = wb.add_format(
            {'font_size': 8, 'align': 'left', 'border': True, 'align': 'center', 'valign': 'vcenter',
             'bg_color': '#FFFF00', 'bold': True, 'text_wrap': True})
        format_cell_02 = wb.add_format(
            {'font_size': 8, 'align': 'left', 'border': True, 'align': 'center', 'valign': 'vcenter',
             'bg_color': '#9ACBF1', 'text_wrap': True, 'bold': True})
        format_cell_03 = wb.add_format(
            {'font_size': 8, 'align': 'left', 'border': True, 'align': 'center', 'valign': 'vcenter',
             'bg_color': '#CFFEFE', 'bold': True, 'text_wrap': True})
        format_cell_04 = wb.add_format(
            {'font_size': 8, 'align': 'left', 'border': True, 'align': 'center', 'valign': 'vcenter',
             'bg_color': '#FFFEA8', 'bold': True, 'text_wrap': True})
        format_cell_05 = wb.add_format(
            {'font_size': 8, 'align': 'left', 'border': True, 'align': 'center', 'valign': 'vcenter',
             'bg_color': '#00B050', 'bold': True, 'text_wrap': True})
        format_data = wb.add_format(
            {'font_size': 8, 'align': 'left', 'border': True, 'align': 'left', 'valign': 'vcenter',
             'bg_color': '#CDFCD0', 'bold': True, 'text_wrap': True})
        list_title1 = ['Persional document \nId', 'Employee Code',
                       'Document Type', 'Số / ID / Code',
                       u'Ngày cấp',u'Ngày hết han',
                       'Country Id', 'Country Name',
                       'City Id', 'City Name',
                       'District Id','District Name',
                       'Is Received Hard Copy Active', 'Active',u'Tình trạng id',u'Tình trạng', u'Ghi chú']


        sheet1.portrait = 0  # Landscape
        sheet1.merge_range('B1:F1', u'Document Type', format_head)
        sheet1.set_column(0, 0, 3)
        col = 0
        row = 3
        # title 1
        for i in range(len(list_title1)):
            sheet1.write(row, col, list_title1[i], format_title_02)
            col += 1
        row = 4
        if data.get('lines', []):
            stt = 0
            for line in data['lines']:
                col_line = 0
                stt += 1
                sheet1.write(row, col_line, line['document_id'], format_data)
                sheet1.write(row, col_line + 1, line['emp_code'], format_data)
                sheet1.write(row, col_line + 2, line['document_type_id'], format_data)
                sheet1.write(row, col_line + 3, line['number'], format_data)
                sheet1.write(row, col_line + 4, line['issue_date'], format_data)
                sheet1.write(row, col_line + 5, line['expiry_date'], format_data)
                sheet1.write(row, col_line + 6, line['country_id'], format_data)
                sheet1.write(row, col_line + 7, line['country_name'], format_data)
                sheet1.write(row, col_line + 8, line['city_id'], format_data)
                sheet1.write(row, col_line + 9, line['city_name'], format_data)
                sheet1.write(row, col_line + 10, line['district_id'], format_data)
                sheet1.write(row, col_line + 11, line['district_name'], format_data)
                sheet1.write(row, col_line + 12, line['is_received_hard_copy'], format_data)
                sheet1.write(row, col_line + 13, line['active'], format_data)
                sheet1.write(row, col_line + 14, line['status_id'], format_data)
                sheet1.write(row, col_line + 15, line['status_name'], format_data)
                sheet1.write(row, col_line + 16, line['note'], format_data)
                row += 1
        wb.close()


rpt_update_emp_xlsx('report.update_employee', 'update.employee',
                    parser=rpt_update_emp_xlsx_parser)
