import unittest

import excel_bancos


class UnificacaoBancosTests(unittest.TestCase):
    def setUp(self):
        self.fields = excel_bancos.get_banco_fields("bancos")

    def test_conjunto_usa_categoria_bancos_e_grupo_30(self):
        category = {"label": "20 - BANCOS", "sheet_name": "20 - BANCOS"}
        self.assertEqual(
            excel_bancos.pn_code_prefix(
                category,
                self.fields,
                {"grupo_codigo": "30", "cj_sufixo": "CJ"},
            ),
            "3020",
        )

    def test_banco_unitario_mantem_prefixo_1020(self):
        category = {"label": "20 - BANCOS", "sheet_name": "20 - BANCOS"}
        self.assertEqual(
            excel_bancos.pn_code_prefix(
                category,
                self.fields,
                {"grupo_codigo": "10", "pre_fixo": "1- BCO"},
            ),
            "1020",
        )

    def test_descricao_do_conjunto_nao_mistura_campos_de_insumo(self):
        data = {
            "grupo_codigo": "30",
            "cj_sufixo": "CJ",
            "cj_encosto": "RECLINAVEL",
            "cj_fornecedor": "MC",
            "cj_linha": "LB",
            "cj_layout": "3,2,3",
            "cj_tipo_cinto": "3P",
            "cj_tipo_revestimento": "TECIDO",
            "cj_especificidade": ["E/S/ J"],
            "cj_acessibilidade_secundaria": "N/A",
        }
        description = excel_bancos.build_descriptions(self.fields, data, "bancos")

        self.assertEqual(
            description["primaria"],
            "CJ BANCOS REC - MC - LB - 3,2,3 - 3P - TECIDO - E/S/ J",
        )
        self.assertEqual(
            description["secundaria"],
            "CJ BANCOS REC - MC - LB - 3,2,3 - 3P - TECIDO - E/S/ J ACESSIBILIDADE: N/A",
        )
        self.assertEqual(description["sufixo"], "CJ")

    def test_campos_do_conjunto_sao_visiveis_somente_no_grupo_30(self):
        conjunto = excel_bancos._visible_field_keys(
            self.fields,
            "bancos",
            {"grupo_codigo": "30", "cj_sufixo": "CJ"},
        )
        insumo = excel_bancos._visible_field_keys(
            self.fields,
            "bancos",
            {"grupo_codigo": "10", "pre_fixo": "1- BCO"},
        )

        self.assertIn("cj_layout", conjunto)
        self.assertNotIn("pre_fixo", conjunto)
        self.assertIn("pre_fixo", insumo)
        self.assertNotIn("cj_layout", insumo)

    def test_conjunto_nao_expoe_detalhe_de_revestimento(self):
        keys = {field["key"] for field in self.fields}
        self.assertNotIn("cj_detalhe_revestimento", keys)

    def test_descricao_normaliza_campos_legados_e_ordem_tecnica(self):
        data = {
            "grupo_codigo": "30",
            "cj_sufixo": "CJ",
            "cj_encosto": "RECLINAVEL",
            "cj_fornecedor": "MC REC",
            "cj_linha": "LB",
            "cj_layout": "4;3;3;3",
            "cj_tipo_cinto": "2P",
            "cj_tipo_revestimento": "TECIDO",
            "cj_especificidade": ["4L REC BJD", "ESJ"],
            "cj_acessibilidade": "FOCA",
        }

        description = excel_bancos.build_descriptions(self.fields, data, "bancos")
        self.assertEqual(
            description["primaria"],
            "CJ BANCOS REC - MC - LB - 4,3,3,3 - 2P - TECIDO - BJD - E/S/ J - 4L REC - FOCA",
        )

    def test_categoria_legacy_e_resolvida_para_bancos(self):
        catalog = {
            "categories": [
                {"key": "bancos", "label": "20 - BANCOS", "fields": []},
                {"key": "cat_20_bco", "label": "20 - CJ-BCO", "fields": []},
            ]
        }
        self.assertEqual(excel_bancos._find_category(catalog, "cat_20_bco")["key"], "bancos")


if __name__ == "__main__":
    unittest.main()
