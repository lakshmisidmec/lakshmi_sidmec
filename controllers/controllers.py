# -*- coding: utf-8 -*-
# from odoo import http


# class XlsxModule(http.Controller):
#     @http.route('/xlsx_module/xlsx_module', auth='public')
#     def index(self, **kw):
#         return "Hello, world"

#     @http.route('/xlsx_module/xlsx_module/objects', auth='public')
#     def list(self, **kw):
#         return http.request.render('xlsx_module.listing', {
#             'root': '/xlsx_module/xlsx_module',
#             'objects': http.request.env['xlsx_module.xlsx_module'].search([]),
#         })

#     @http.route('/xlsx_module/xlsx_module/objects/<model("xlsx_module.xlsx_module"):obj>', auth='public')
#     def object(self, obj, **kw):
#         return http.request.render('xlsx_module.object', {
#             'object': obj
#         })

