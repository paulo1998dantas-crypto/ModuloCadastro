import unittest

import excel_bancos
import supabase_store


class UnificacaoBancosTests(unittest.TestCase):
    def setUp(self):
        self.fields = excel_bancos.get_banco_fields("bancos")

    def test_conjunto_usa_categoria_bancos_e_grupo_30(self):
        category = {"label": "20 - BANCOS", "sheet_name": "20 - BANCOS"}
        self.assertEqual(
            excel_bancos.pn_code_prefix(
                category,
                self.fields,
                {"grupo_codigo": "30", "pre_fixo": "8- CJ"},
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
            "pre_fixo": "8- CJ",
            "encosto": "2- RECLINAVEL",
            "fornecedor": "1- MC",
            "linha": "1- LB",
            "cj_layout": "3,2,3",
            "tipo_cinto": "2- 3P",
            "tipo_revestimento": "1- TECIDO",
            "especificidade": ["11- E/S/J"],
            "cj_acessibilidade_secundaria": "N/A",
        }
        description = excel_bancos.build_descriptions(self.fields, data, "bancos")

        self.assertEqual(
            description["primaria"],
            "CJ BANCOS REC - MC - LB - 3,2,3 - 3P - TECIDO - E/S/ J",
        )
        self.assertEqual(
            description["secundaria"],
            "CJ BANCOS REC - MC - LB - 3,2,3 - 3P - TECIDO - E/S/ J | ACESSIBILIDADE: N/A",
        )
        self.assertEqual(description["sufixo"], "CJ")

    def test_campos_do_conjunto_sao_visiveis_somente_no_grupo_30(self):
        conjunto = excel_bancos._visible_field_keys(
            self.fields,
            "bancos",
            {"grupo_codigo": "30", "pre_fixo": "8- CJ"},
        )
        insumo = excel_bancos._visible_field_keys(
            self.fields,
            "bancos",
            {"grupo_codigo": "10", "pre_fixo": "1- BCO"},
        )

        self.assertIn("cj_layout", conjunto)
        self.assertIn("fornecedor", conjunto)
        self.assertIn("linha", conjunto)
        self.assertIn("especificidade", conjunto)
        self.assertIn("tipo_costura", conjunto)
        self.assertIn("cor_da_linha", conjunto)
        self.assertIn("cor_do_revestimento", conjunto)
        self.assertIn("pre_fixo", conjunto)
        self.assertIn("pre_fixo", insumo)
        self.assertNotIn("cj_sufixo", conjunto)
        self.assertNotIn("cj_sufixo", insumo)
        self.assertNotIn("cj_layout", insumo)
        self.assertNotIn("cj_acessibilidade", insumo)

    def test_conjunto_nao_expoe_detalhe_de_revestimento(self):
        keys = {field["key"] for field in self.fields}
        self.assertNotIn("cj_detalhe_revestimento", keys)

    def test_descricao_normaliza_campos_legados_e_ordem_tecnica(self):
        data = {
            "grupo_codigo": "30",
            "cj_sufixo": "CJ",
            "encosto": "2- RECLINAVEL",
            "fornecedor": "1- MC REC",
            "linha": "1- LB",
            "cj_layout": "4;3;3;3",
            "tipo_cinto": "1- 2P",
            "tipo_revestimento": "1- TECIDO",
            "especificidade": ["4L REC BJD", "ESJ"],
            "cj_acessibilidade": "FOCA",
        }

        description = excel_bancos.build_descriptions(self.fields, data, "bancos")
        self.assertEqual(
            description["primaria"],
            "CJ BANCOS REC - MC - LB - 4,3,3,3 - 2P - TECIDO - BJD - E/S/ J - 4L REC - FOCA",
        )

    def test_campos_compartilhados_nao_sao_duplicados_no_conjunto(self):
        keys = {field["key"] for field in self.fields}
        self.assertIn("pre_fixo", keys)
        self.assertNotIn("cj_sufixo", keys)
        self.assertNotIn("cj_fornecedor", keys)
        self.assertNotIn("cj_linha", keys)
        self.assertNotIn("cj_encosto", keys)
        self.assertNotIn("cj_tipo_cinto", keys)
        self.assertNotIn("cj_tipo_revestimento", keys)
        self.assertNotIn("cj_especificidade", keys)

    def test_linha_executiva_usa_detalhes_canonicos_na_descricao_primaria(self):
        data = {
            "grupo_codigo": "30",
            "cj_sufixo": "CJ",
            "encosto": "2- RECLINAVEL",
            "fornecedor": "2- CS",
            "linha": "2- LE",
            "cj_layout": "3,3",
            "tipo_cinto": "2- 3P",
            "tipo_revestimento": "2- COURVIN",
            "especificidade": ["11- E/S/J", "EXECUTIVO"],
            "cor_do_revestimento": "3- CAPA LE MARROM",
            "tipo_costura": "3- CS COSTURA DIAMANTE",
            "cor_da_linha": "6- COR LINHA DOURADO",
        }

        description = excel_bancos.build_descriptions(self.fields, data, "bancos")
        primaria = (
            "CJ BANCOS REC - CS - LE - 3,3 - 3P - "
            "COURVIN MARROM/DIAMANTE/LINHA DOURADA - E/S/ J - EXECUTIVO"
        )
        self.assertEqual(description["primaria"], primaria)
        self.assertEqual(description["secundaria"], primaria)

    def test_linha_lb_mantem_detalhes_de_revestimento_somente_na_secundaria(self):
        data = {
            "grupo_codigo": "30",
            "pre_fixo": "8- CJ",
            "encosto": "2- RECLINAVEL",
            "fornecedor": "1- MC",
            "linha": "1- LB",
            "cj_layout": "4,3,3,3",
            "tipo_cinto": "2- 3P",
            "tipo_revestimento": "1- TECIDO",
            "especificidade": ["3- BJD"],
            "cor_do_revestimento": "1- CAPA LB PADRAO JI",
            "tipo_costura": "1- COSTURA LB",
            "cor_da_linha": "1- COR LINHA LB",
            "cj_acessibilidade": "1- N/A",
        }

        description = excel_bancos.build_descriptions(self.fields, data, "bancos")
        primaria = "CJ BANCOS REC - MC - LB - 4,3,3,3 - 3P - TECIDO - BJD"
        self.assertEqual(description["primaria"], primaria)
        self.assertEqual(
            description["secundaria"],
            primaria + " | REVESTIMENTO: CAPA LB PADRAO JI/COSTURA LB/LINHA LB | ACESSIBILIDADE: N/A",
        )

    def test_acessibilidade_do_conjunto_e_unica_numerada_e_condicional(self):
        accessibility_fields = [field for field in self.fields if field["key"] == "cj_acessibilidade"]
        self.assertEqual(len(accessibility_fields), 1)
        self.assertNotIn("cj_acessibilidade_secundaria", {field["key"] for field in self.fields})
        self.assertEqual(
            accessibility_fields[0]["options"],
            [
                "1- N/A",
                "2- FOCA",
                "3- ELEVITTA",
                "4- PLATAFORMA BI-PARTIDA",
                "5- PLATAFORMA FECHADA",
            ],
        )

        base = {
            "grupo_codigo": "30",
            "pre_fixo": "8- CJ",
            "encosto": "2- RECLINAVEL",
            "fornecedor": "1- MC",
            "linha": "1- LB",
            "cj_layout": "3,3",
            "tipo_cinto": "2- 3P",
            "tipo_revestimento": "1- TECIDO",
        }
        description_na = excel_bancos.build_descriptions(
            self.fields, {**base, "cj_acessibilidade": "1- N/A"}, "bancos"
        )
        self.assertNotIn(" - N/A", description_na["primaria"])
        self.assertIn("ACESSIBILIDADE: N/A", description_na["secundaria"])

        description_foca = excel_bancos.build_descriptions(
            self.fields, {**base, "cj_acessibilidade": "2- FOCA"}, "bancos"
        )
        self.assertTrue(description_foca["primaria"].endswith(" - FOCA"))

    def test_regras_condicionais_do_sistema_sao_expostas_em_opcoes(self):
        rules = excel_bancos.get_conditional_rules("bancos")
        for target in ("cor_do_revestimento", "tipo_costura", "cor_da_linha"):
            matching = [
                rule
                for rule in rules
                if rule["source_field_key"] == "linha"
                and rule["target_field_key"] == target
                and rule["action"] == "set_primary"
                and "LE" in rule["source_value_labels"][0]
            ]
            self.assertEqual(len(matching), 1)
            self.assertIn(matching[0]["origin"], {"system", "catalog"})
        accessibility = [rule for rule in rules if rule["key"] == "cond_acessibilidade_na_secundaria"]
        self.assertEqual(len(accessibility), 1)

    def test_toda_especificidade_do_conjunto_fica_na_primaria(self):
        data = {
            "grupo_codigo": "30",
            "cj_sufixo": "CJ",
            "encosto": "2- RECLINAVEL",
            "fornecedor": "4- STF",
            "linha": "2- LE",
            "cj_layout": "4,2-1,2-1,3",
            "tipo_cinto": "2- 3P",
            "tipo_revestimento": "2- COURVIN",
            "especificidade": ["EXECUTIVO", "MASTER - PME"],
            "cor_do_revestimento": "4- CAPA LE PRETA",
            "tipo_costura": "7- COSTURA ST02 STF",
            "cor_da_linha": "3- COR LINHA BRANCA",
        }

        description = excel_bancos.build_descriptions(self.fields, data, "bancos")
        primaria = (
            "CJ BANCOS REC - STF - LE - 4,2-1,2-1,3 - 3P - "
            "COURVIN PRETO/ST02/LINHA BRANCA - MASTER - PME - EXECUTIVO"
        )
        self.assertEqual(description["primaria"], primaria)
        self.assertEqual(description["secundaria"], primaria)

    def test_conjuntos_le_executivos_sao_recompostos_pelos_campos(self):
        cases = [
            (
                "30200032",
                "3,3",
                "2- CS",
                "3- CAPA LE MARROM",
                "4- CS COSTURA BOOMERANG",
                "6- COR LINHA DOURADO",
                ["11- E/S/J", "EXECUTIVO"],
                "CJ BANCOS REC - CS - LE - 3,3 - 3P - COURVIN "
                "MARROM/BOOMERANG/LINHA DOURADA - E/S/ J - EXECUTIVO",
            ),
            (
                "30200036",
                "3,3",
                "2- CS",
                "5- CAPA LE PRETA/CINZA",
                "3- CS COSTURA DIAMANTE",
                "7- COR LINHA CINZA",
                ["11- E/S/J", "EXECUTIVO"],
                "CJ BANCOS REC - CS - LE - 3,3 - 3P - COURVIN "
                "PRETO/CINZA/DIAMANTE/LINHA CINZA - E/S/ J - EXECUTIVO",
            ),
            (
                "30200039",
                "3,2-1",
                "2- CS",
                "4- CAPA LE PRETA",
                "4- CS COSTURA BOOMERANG",
                "4- COR LINHA PRETA",
                ["11- E/S/J", "EXECUTIVO"],
                "CJ BANCOS REC - CS - LE - 3,2-1 - 3P - COURVIN "
                "PRETO/BOOMERANG/LINHA PRETA - E/S/ J - EXECUTIVO",
            ),
        ]
        for sku, layout, fornecedor, cor, costura, linha, especificidade, esperado in cases:
            with self.subTest(sku=sku):
                data = {
                    "grupo_codigo": "30",
                    "cj_sufixo": "CJ",
                    "encosto": "2- RECLINAVEL",
                    "fornecedor": fornecedor,
                    "linha": "2- LE",
                    "cj_layout": layout,
                    "tipo_cinto": "2- 3P",
                    "tipo_revestimento": "2- COURVIN",
                    "especificidade": especificidade,
                    "cor_do_revestimento": cor,
                    "tipo_costura": costura,
                    "cor_da_linha": linha,
                }
                description = excel_bancos.build_descriptions(self.fields, data, "bancos")
                self.assertEqual(description["primaria"], esperado)
                self.assertEqual(description["secundaria"], esperado)

    def test_edicao_le_chave_legada_sem_duplicar_o_campo(self):
        groups = supabase_store._groups_from_record(
            self.fields,
            {
                "form_values": {"cj_especificidade": ["11- E/S/J", "EXECUTIVO"]},
                "field_values": {},
            },
        )
        self.assertEqual(groups["especificidade"], ["11- E/S/J", "EXECUTIVO"])

    def test_prefixo_cj_e_unico_e_aciona_grupo_conjunto(self):
        prefix_fields = [field for field in self.fields if field["key"] == "pre_fixo"]
        self.assertEqual(len(prefix_fields), 1)
        self.assertNotIn("cj_sufixo", {field["key"] for field in self.fields})
        self.assertIn("8- CJ", prefix_fields[0]["options"])

        data = {"grupo_codigo": "10", "pre_fixo": "8- CJ"}
        self.assertTrue(excel_bancos.is_banco_conjunto(data))
        self.assertEqual(
            excel_bancos.pn_code_prefix(
                {"label": "20 - BANCOS", "sheet_name": "20 - BANCOS"},
                self.fields,
                data,
            ),
            "3020",
        )

    def test_registro_legado_cj_sufixo_abre_com_prefixo_canonico(self):
        groups = supabase_store._groups_from_record(
            self.fields,
            {
                "form_values": {"grupo_codigo": ["30"], "cj_sufixo": ["CJ"]},
                "field_values": {},
            },
        )
        self.assertEqual(groups["pre_fixo"], ["8- CJ"])

    def test_salvamento_do_grupo_conjunto_persiste_prefixo_canonico(self):
        groups = supabase_store._field_groups(
            self.fields,
            {"grupo_codigo": "30", "pre_fixo": ""},
        )
        self.assertEqual(groups["grupo_codigo"], ["30"])
        self.assertEqual(groups["pre_fixo"], ["8- CJ"])

    def test_banco_unitario_preserva_regra_e_descricao_existente(self):
        data = {
            "grupo_codigo": "10",
            "pre_fixo": "1- BCO",
            "encosto": "2- RECLINAVEL",
            "lotacao": "3- 3 L",
            "especificidade": "1- NORMAL",
            "fornecedor": "1- MC",
            "linha": "1- LB",
            "tipo_cinto": "2- 3P",
            "tipo_revestimento": "2- COURVIN",
        }
        description = excel_bancos.build_descriptions(self.fields, data, "bancos")
        self.assertTrue(description["primaria"].startswith("BCO"))
        self.assertNotIn("CJ BANCOS", description["primaria"])
        excel_bancos._validate_banco_dependencies(self.fields, data)

        data["especificidade"] = ["1- NORMAL", "3- BJD"]
        with self.assertRaisesRegex(ValueError, "somente uma ESPECIFICIDADE"):
            excel_bancos._validate_banco_dependencies(self.fields, data)

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
