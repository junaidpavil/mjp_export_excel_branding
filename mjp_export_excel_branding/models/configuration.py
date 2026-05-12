from odoo import api, fields, models


class ResConfigSettings(models.TransientModel):
    _inherit = "res.config.settings"

    excel_export_model_ids = fields.Many2many(
        "ir.model",
        "excel_export_config_model_rel",
        string="Custom Layout Export Models",
        help="Models that will use the branded header layout during Excel export.",
    )

    def get_values(self):
        res = super().get_values()
        param = self.env["ir.config_parameter"].sudo().get_param(
            "excel_export.custom_layout_models", ""
        )
        if param:
            model_ids = self.env["ir.model"].sudo().search(
                [("model", "in", param.split(","))]
            )
            res["excel_export_model_ids"] = [fields.Command.set(model_ids.ids)]
        else:
            res["excel_export_model_ids"] = [fields.Command.set([])]
        return res

    def set_values(self):
        super().set_values()
        model_names = self.excel_export_model_ids.mapped("model")
        self.env["ir.config_parameter"].sudo().set_param(
            "excel_export.custom_layout_models",
            ",".join(model_names),
        )
        
