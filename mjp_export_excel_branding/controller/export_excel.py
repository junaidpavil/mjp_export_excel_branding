import base64
import io
import json
import logging
import operator

from odoo import fields, http
from odoo.addons.web.controllers.export import (
    ExcelExport,
    ExportXlsxWriter,
    GroupExportXlsxWriter,
    GroupsTreeNode,
)
from odoo.http import content_disposition, request
from odoo.tools import osutil

_logger = logging.getLogger(__name__)
# CUSTOM_LAYOUT_MODELS = {"sale.order"}

def _get_custom_layout_models():
    param = request.env["ir.config_parameter"].sudo().get_param(
        "excel_export.custom_layout_models", ""
    )
    return set(param.split(",")) if param else set()

class ExportLayoutXlsxMixin:
    def __init__(
        self,
        export_fields,
        columns_headers,
        row_count,
        company=None,
        user=None,
        record_count=0,
    ):
        self.company = (company or request.env.company).sudo()
        self.user = user or request.env.user
        self.record_count = record_count

        self.header_row = 3
        self.first_data_row = 4

        super().__init__(export_fields, columns_headers, row_count + self.first_data_row)

        self.company_name_style = self.workbook.add_format(
            {
                "bold": True,
                "font_size": 16,
                "align": "left",
                "valign": "vcenter",
            }
        )
        self.meta_style = self.workbook.add_format(
            {
                "font_size": 11,
                "align": "left",
                "valign": "vcenter",
            }
        )
        self.count_style = self.workbook.add_format(
            {
                "bold": True,
                "font_size": 11,
                "align": "left",
                "valign": "vcenter",
            }
        )
        self.blank_style = self.workbook.add_format(
            {
                "align": "center",
                "valign": "vcenter",
            }
        )
        self.header_style = self.workbook.add_format(
            {
                "bold": True,
                "text_wrap": True,
                "align": "center",
                "valign": "vcenter",
                "bg_color": "#D9EAF7",
                "border": 1,
            }
        )

    def _get_logo_b64(self):
        if self.company._fields.get("logo_web") and self.company.logo_web:
            return self.company.logo_web
        return self.company.logo

    def _get_generated_datetime(self):
        user_dt = fields.Datetime.context_timestamp(self.user, fields.Datetime.now())
        return user_dt.strftime("%d-%m-%Y %H:%M:%S")

    def write_header(self):
        data_last_col = max(len(self.columns_headers) - 1, 0)
        meta_last_col = data_last_col

        self.worksheet.set_column(0, meta_last_col, 20)
        self.worksheet.set_row(0, 80)
        self.worksheet.set_row(1, 22)
        self.worksheet.set_row(2, 22)
        self.worksheet.set_row(self.header_row, 24)

        logo_b64 = self._get_logo_b64()
        has_logo = bool(logo_b64)

        logo_blank_style = self.workbook.add_format({
            "align": "center",
            "valign": "vcenter",
            "border": 0,
            "left": 0,
            "right": 0,
            "top": 0,
            "bottom": 0,
        })

        company_name_style_no_border = self.workbook.add_format({
            "bold": True,
            "font_size": 16,
            "align": "right",
            "valign": "vcenter",
            "border": 0,
            "left": 0,
            "right": 0,
            "top": 0,
            "bottom": 0,
        })

        meta_style_no_border = self.workbook.add_format({
            "font_size": 11,
            "align": "left",
            "valign": "vcenter",
            "border": 0,
        })

        count_style_no_border = self.workbook.add_format({
            "bold": True,
            "font_size": 11,
            "align": "left",
            "valign": "vcenter",
            "border": 0,
        })

        if has_logo:
            # Logo only in col 0 (single cell, no merge)
            self.worksheet.write(0, 0, "", logo_blank_style)
            self.worksheet.insert_image(
                0,
                0,
                "company_logo.png",
                {
                    "image_data": io.BytesIO(base64.b64decode(logo_b64)),
                    "x_scale": 0.55,
                    "y_scale": 0.55,
                    "x_offset": 8,
                    "y_offset": 6,
                    "object_position": 1,
                },
            )
            text_start_col = 1  # Start company name from col 1
        else:
            text_start_col = 0

        # Row 0: Company name — merge from text_start_col to last col if possible
        if text_start_col < meta_last_col:
            self.worksheet.merge_range(
                0, text_start_col, 0, meta_last_col,
                self.company.name or "",
                company_name_style_no_border,
            )
        else:
            # Only one cell available, just write directly
            self.worksheet.write(
                0, text_start_col,
                self.company.name or "",
                company_name_style_no_border,
            )

        # Row 1: Generated by — merge all cols
        if meta_last_col > 0:
            self.worksheet.merge_range(
                1, 0, 1, meta_last_col,
                f"Generated by {self.user.name} on {self._get_generated_datetime()}",
                meta_style_no_border,
            )
        else:
            self.worksheet.write(
                1, 0,
                f"Generated by {self.user.name} on {self._get_generated_datetime()}",
                meta_style_no_border,
            )

        # Row 2: Total count — merge all cols
        if meta_last_col > 0:
            self.worksheet.merge_range(
                2, 0, 2, meta_last_col,
                f"Total selected records count: {self.record_count}",
                count_style_no_border,
            )
        else:
            self.worksheet.write(
                2, 0,
                f"Total selected records count: {self.record_count}",
                count_style_no_border,
            )

        # Row 3: Column headers
        for col, column_header in enumerate(self.columns_headers):
            self.write(self.header_row, col, column_header, self.header_style)

        self.worksheet.freeze_panes(self.first_data_row, 0)


class CustomExportXlsxWriter(ExportLayoutXlsxMixin, ExportXlsxWriter):
    pass


class CustomGroupExportXlsxWriter(ExportLayoutXlsxMixin, GroupExportXlsxWriter):
    pass


class ExcelExportInherit(ExcelExport):
    @http.route()
    def web_export_xlsx(self, data):
        return super().web_export_xlsx(data)

    def base(self, data):
        params = json.loads(data)
        model, export_fields, ids, domain, import_compat = operator.itemgetter(
            "model", "fields", "ids", "domain", "import_compat"
        )(params)

        Model = request.env[model].with_context(
            import_compat=import_compat, **params.get("context", {})
        )

        if not Model._is_an_ordinary_table():
            export_fields = [field for field in export_fields if field["name"] != "id"]

        field_names = [field["name"] for field in export_fields]
        if import_compat:
            columns_headers = field_names
        else:
            columns_headers = [field["label"].strip() for field in export_fields]

        groupby = params.get("groupby")
        if not import_compat and groupby:
            groupby_type = [
                Model._fields[x.split(":", 1)[0].split(".", 1)[0]].type for x in groupby
            ]
            domain = [("id", "in", ids)] if ids else domain
            read_context = Model.env.context
            if ids:
                Model = Model.with_context(active_test=False)

            groups_data = Model.read_group(domain, ["__count"], groupby, lazy=False)
            tree = GroupsTreeNode(Model, field_names, groupby, groupby_type, read_context)

            records = Model.browse()
            for leaf in groups_data:
                records |= tree.insert_leaf(leaf)

            record_count = len(records)
            response_data = self.from_group_data(
                export_fields, columns_headers, tree, record_count, model=model
            )
        else:
            records = (
                Model.browse(ids)
                if ids
                else Model.search(domain, offset=0, limit=False, order=False)
            )
            export_data = records.export_data(field_names).get("datas", [])
            record_count = len(records)
            response_data = self.from_data(
                export_fields, columns_headers, export_data, record_count, model=model
            )

        _logger.info(
            "User %d exported %d %r records from %s. Fields: %s. %s: %s",
            request.env.user.id,
            len(records.ids),
            records._name,
            request.httprequest.environ["REMOTE_ADDR"],
            ",".join(field_names),
            "IDs sample" if ids else "Domain",
            records.ids[:10] if ids else domain,
        )

        return request.make_response(
            response_data,
            headers=[
                (
                    "Content-Disposition",
                    content_disposition(
                        osutil.clean_filename(self.filename(model) + self.extension)
                    ),
                ),
                ("Content-Type", self.content_type),
            ],
        )



    def from_data(self, export_fields, columns_headers, rows, record_count, model=None):

        if model in _get_custom_layout_models():
            writer_class = CustomExportXlsxWriter
            kwargs = dict(
                company=request.env.company.sudo(),
                user=request.env.user,
                record_count=record_count,
            )
        else:
            writer_class = ExportXlsxWriter
            kwargs = {}

        with writer_class(export_fields, columns_headers, len(rows), **kwargs) as xlsx_writer:
            start_row = getattr(xlsx_writer, "first_data_row", 1)
            for row_index, row in enumerate(rows, start=start_row):
                for cell_index, cell_value in enumerate(row):
                    xlsx_writer.write_cell(row_index, cell_index, cell_value)
        return xlsx_writer.value

    def from_group_data(self, export_fields, columns_headers, groups, record_count, model=None):

        if model in _get_custom_layout_models():
            writer_class = CustomGroupExportXlsxWriter
            kwargs = dict(
                company=request.env.company.sudo(),
                user=request.env.user,
                record_count=record_count,
            )
        else:
            writer_class = GroupExportXlsxWriter
            kwargs = {}

        with writer_class(export_fields, columns_headers, groups.count, **kwargs) as xlsx_writer:
            row, col = getattr(xlsx_writer, "first_data_row", 1), 0
            for group_name, group in groups.children.items():
                row, col = xlsx_writer.write_group(row, col, group_name, group)
        return xlsx_writer.value