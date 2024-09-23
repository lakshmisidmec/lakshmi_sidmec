from odoo import models,fields
import xlsxwriter

class SaleQuotation(models.AbstractModel):
    _name = 'report.xlsx_module.report_sale_xlsx'
    _inherit = 'report.report_xlsx.abstract'

    def generate_xlsx_report(self, workbook, data, orders):
        sheet = workbook.add_worksheet('Quotation Report')
        sheet.hide_gridlines(2)

        # Insert logo
        logo_path = '/home/lakshmi/odoo17/custom_addons/xlsx_module/static/src/img/file.png'
        sheet.insert_image('A3', logo_path)
        company = self.env.company

        header_format = workbook.add_format({
            'bold': True,
            'font_color': '#1A237E',
            'bg_color': 'white',
            'font_size': 25,
            'align': 'left',
            'valign': 'top',
            'text_wrap': True,
            'border': 1,
        })

        head1_format = workbook.add_format({
            'font_color': 'black',
            'bg_color': '#90CAF9',
            'align': 'center',
            'font_size': 11,
            'bold': True,
            'border': 1,
        })

        normal_format = workbook.add_format({
            'font_color': 'black',
            'align': 'top',
            'font_size': 11,
        })

        table_cell_format = workbook.add_format({
            'font_color': 'black',
            'align': 'top',
            'font_size': 11,
            'border': 1,
            'text_wrap': True,
        })

        total_format = workbook.add_format({
            'font_color': 'black',
            'bold': True,
            'font_size': 11,
            'border': 1,
        })

        footer_format = workbook.add_format({
            'bold': True,
            'align': 'center',
            'font_size': 10,
        })

        sheet.merge_range(2,8,4,11, 'Quotation', header_format)
        for order in orders:
                sheet.write(6, 8, 'Date:',table_cell_format)
                sheet.merge_range(6, 9, 6, 10, str(order.date_order),table_cell_format)
                sheet.write(7, 8, 'Order#', table_cell_format)
                sheet.merge_range(7, 9, 7, 10, order.name, table_cell_format)

                sheet.merge_range(9, 0, 9, 3, 'Invoicing Address:',head1_format)
                sheet.write('A11', orders.partner_invoice_id.name, normal_format)
                # sheet.merge_range(10, 0, 10, 3, orders.partner_invoice_id.street, normal_format)
                sheet.write('A12', company.street, normal_format)
                sheet.write('A13', company.country_id.name, normal_format)

                sheet.merge_range(9, 8, 9, 11, 'Shipping Address:',head1_format)
                sheet.merge_range(10, 8, 10, 11, orders.partner_shipping_id.name,normal_format)
                sheet.merge_range(11, 8, 11, 11, company.street,normal_format)
                sheet.merge_range(12, 8, 12, 11, company.country_id.name,normal_format)

                sheet.merge_range(14, 0, 14, 1, 'Sales Person',head1_format)
                sheet.merge_range(14, 2, 14, 3, 'Shipping Method',head1_format)
                sheet.merge_range(14, 4, 14, 5, 'Shipping Terms',head1_format)
                sheet.merge_range(14, 6, 14, 7, 'Payment Terms',head1_format)
                sheet.merge_range(14, 8, 14, 9, 'Due Date',head1_format)
                sheet.merge_range(14, 10, 14, 11, 'Delivery Date',head1_format)

                sheet.merge_range(15, 0, 16, 1, order.user_id.name,table_cell_format)  # Sales Person
                sheet.merge_range(15, 2, 16, 3, 'Standard Delivery',table_cell_format)  # Shipping Method
                sheet.merge_range(15, 4, 16, 5, order.payment_term_id.name,table_cell_format)  # Shipping terms
                sheet.merge_range(15, 6, 16, 7, order.payment_term_id.name,table_cell_format)  # Payment Terms
                sheet.merge_range(15, 8, 16, 9, str(order.validity_date),table_cell_format) # Due Date
                sheet.merge_range(15, 10, 16, 11, str(order.commitment_date),table_cell_format)  # Delivery Date

                sheet.merge_range(18, 0, 18, 1, 'Product', head1_format)
                sheet.merge_range(18, 2, 18, 5, 'Description', head1_format)
                sheet.write(18, 6, 'Qty', head1_format)
                sheet.write(18, 7, 'UOM', head1_format)
                sheet.write(18, 8, 'Unit Price', head1_format)
                sheet.write(18, 9, 'Taxes', head1_format)
                sheet.merge_range(18, 10, 18, 11, 'Line Total', head1_format)
                row = 19
                asc_row = 0
                for line in order.order_line:
                    description_lines = line.name.split('\n')
                    # print("description_lines: ",description_lines)
                    num_of_lines = len(description_lines)
                    if num_of_lines>1:
                        for i, desc in enumerate(description_lines):
                            sheet.merge_range(row + i, 2, row + i, 5, '* '+desc, table_cell_format)
                            asc_row = row + i
                        sheet.merge_range(row, 0, asc_row, 1, line.product_template_id.name, table_cell_format)
                        sheet.merge_range(row, 6, asc_row, 6, line.product_uom_qty, table_cell_format)
                        sheet.merge_range(row, 7, asc_row, 7, line.product_uom.name, table_cell_format)
                        sheet.merge_range(row, 8, asc_row, 8, line.price_unit, table_cell_format)
                        sheet.merge_range(row, 9, asc_row, 9, line.price_tax, table_cell_format)
                        sheet.merge_range(row, 10, asc_row, 11, line.price_subtotal, table_cell_format)
                        asc_row += 1
                        row = asc_row
                    else:
                        sheet.merge_range(row, 0, row, 1, line.product_template_id.name, table_cell_format)
                        sheet.merge_range(row, 2, row, 5, '* '+line.name,table_cell_format)
                        sheet.write(row, 6, line.product_uom_qty, table_cell_format)
                        sheet.write(row, 7, line.product_uom.name, table_cell_format)
                        sheet.write(row, 8, line.price_unit, table_cell_format)
                        sheet.write(row, 9, line.price_tax, table_cell_format)
                        sheet.merge_range(row, 10, row, 11, line.price_subtotal, table_cell_format)
                        row+=1

                ins_row = row+1
                sheet.merge_range(ins_row, 0, ins_row, 7, 'Special Notes and Instructions', head1_format)
                sheet.merge_range(ins_row+1, 0, ins_row+2, 7, 'Happy Shopping.....', normal_format)
                sheet.write(ins_row, 9, 'Sub total', table_cell_format)
                sheet.write(ins_row+1, 9, 'Taxes', table_cell_format)
                sheet.write(ins_row+2, 9, 'Total', table_cell_format)

                sheet.write(ins_row, 10, company.currency_id.name,table_cell_format)
                sheet.write(ins_row + 1, 10, company.currency_id.name,table_cell_format)
                sheet.write(ins_row + 2, 10, company.currency_id.name,total_format)

                sheet.write(ins_row , 11, order.amount_untaxed,table_cell_format)
                sheet.write(ins_row + 1, 11, order.amount_tax,table_cell_format)
                sheet.write(ins_row + 2, 11, order.amount_total,total_format)

                footer_row = ins_row+5
                sheet.merge_range(footer_row, 0, footer_row, 11, 'Thank you for your business!', footer_format)
                sheet.merge_range(footer_row+1, 0, footer_row+1, 11, 'Should you have any concern regarding the quotation, '
                                                                     'please feel free to contact on +1 555 555 5555', footer_format)

