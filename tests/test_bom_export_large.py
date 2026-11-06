import tempfile
import unittest
import urllib.parse
from pathlib import Path
from unittest.mock import patch

from openpyxl import load_workbook

import supabase_store as store


class LargeBomExportTests(unittest.TestCase):
    def setUp(self):
        self.headers = [
            {'id': i, 'parent_sku': f'3020{i:04}', 'parent_category_label': 'BANCOS'}
            for i in range(1, 1336)
        ]
        self.components = [
            {'id': i * 8 + j, 'bom_id': i, 'parent_sku': f'3020{i:04}',
             'component_sku': f'1022{i:04}', 'quantidade': j, 'ordem': j}
            for i in range(1, 1336) for j in range(1, 9)
        ]
        self.catalog = [
            {'sku': f'{prefix}{i:04}', 'descricao_primaria': f'ITEM {i}',
             'unidade': 'pc', 'ativo': not (prefix == '3020' and i == 1335)}
            for prefix in ('1022', '3020') for i in range(1, 1336)
        ]
        self.requests = []

    def request(self, method, table, query, **kwargs):
        self.assertEqual(method, 'GET')
        self.assertLess(len(urllib.parse.urlencode(query)), 6000)
        self.requests.append((table, dict(query)))
        rows = {store.BOM_HEADERS_TABLE: self.headers,
                store.BOM_COMPONENTS_TABLE: self.components,
                store.REGISTRATIONS_TABLE: self.catalog}[table]
        params = dict(query)
        for key, value in query:
            if value.startswith('in.('):
                codes = set(value[4:-1].split(','))
                rows = [row for row in rows if str(row[key]) in codes]
        start = int(params.get('offset', 0))
        return rows[start:start + int(params.get('limit', 1000))]

    def test_export_batches_large_catalog_and_preserves_over_ten_thousand_components(self):
        with tempfile.TemporaryDirectory() as directory, \
             patch.object(store, '_request', side_effect=self.request), \
             patch.object(store, 'EXPORT_DIR', Path(directory)):
            output = store.export_boms()
            wb = load_workbook(output, data_only=True, read_only=True)
            try:
                rows = list(wb['BOM'].values)
            finally:
                wb.close()
        self.assertEqual(len(rows) - 1, 1334 * 8)
        self.assertEqual(rows[-1][1], '30201334')
        self.assertEqual(rows[-1][-1], 8)
        self.assertNotIn('30201335', {row[1] for row in rows[1:]})
        self.assertTrue(any(q.get('offset') == '1000' for _, q in self.requests))

    def test_component_search_paginates_matches_and_batches_header_filters(self):
        with patch.object(store, '_request', side_effect=self.request):
            rows = store.list_boms(component_query='ITEM', include_inactive=True, limit=5000)
        self.assertEqual(len(rows), 1335)
        self.assertEqual(sum(len(row['components']) for row in rows), 1335 * 8)
        self.assertEqual(rows[-1]['parent_sku'], '30201335')


if __name__ == '__main__':
    unittest.main()

