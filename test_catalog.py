import tempfile, unittest, json
from pathlib import Path
from openpyxl import Workbook
try:
    import catalog
except ModuleNotFoundError:
    catalog = None


class CatalogTests(unittest.TestCase):
    def setUp(self):
        self.assertIsNotNone(catalog, 'Falta implementar el catálogo y la búsqueda segura')
        self.tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self.tmp.cleanup)
        self.root = Path(self.tmp.name)

    def store(self, rows):
        s = catalog.Store(self.root / 'db.sqlite')
        self.addCleanup(s.close)
        s.replace(rows)
        return s

    def row(self, code='001234567890', **kw):
        return dict(code=code, name='Producto de prueba', category='Producto', **kw)

    def test_variable_lengths_and_leading_zero_are_exact(self):
        s = self.store([self.row(), self.row('1234567890'), self.row('123456789012345678901')])
        for code in ('001234567890', '1234567890', '123456789012345678901'):
            self.assertEqual(s.lookup(code)[0]['record']['code'], code)
        self.assertEqual(s.lookup('01234567890'), [])
        self.assertEqual(s.lookup('123456789'), [])

    def test_scanner_enter_and_printed_spaces(self):
        s = self.store([self.row('012345678901')])
        self.assertEqual(len(s.lookup('0 123456 78901\r\n')), 1)
        with self.assertRaises(ValueError):
            s.lookup('012-345678901')
        with self.assertRaises(ValueError):
            s.lookup('7e11')

    def test_shared_code_keeps_all_candidates(self):
        s = self.store([self.row('000987654321'),self.row('000987654321', notes='Otra presentación')])
        self.assertEqual(len(s.lookup('000987654321')), 2)

    def test_itf_and_explicit_dv_and_calculated_dv_are_distinguished(self):
        s = self.store([self.row('012345678901', dv='0', itf14='30123456789012'), self.row('123456789012')])
        self.assertEqual(s.lookup('0123456789010')[0]['kind'], 'DV en Excel')
        self.assertEqual(s.lookup('30123456789012')[0]['kind'], 'ITF14 en Excel')
        self.assertEqual(s.lookup('1234567890128')[0]['kind'], 'DV calculado')
        self.assertEqual(s.lookup('1234567890129'), [])

    def test_derived_collision_cannot_hide_an_exact_record(self):
        s = self.store([self.row('012345678901'), self.row('0123456789012')])
        self.assertEqual(len(s.lookup('0123456789012')), 2)

    def test_crud_survives_restart_and_delete_is_recoverable(self):
        s = self.store([self.row()])
        r = s.list()[0]
        r['name'] = 'Nombre editado'
        s.save(r)
        s.save(self.row('999'))
        s.delete(r['id'])
        self.assertEqual(len(s.list()), 1)
        s.undo_delete()
        self.assertEqual(len(s.list()), 2)
        s.close()
        other = catalog.Store(self.root/'db.sqlite')
        self.addCleanup(other.close)
        self.assertEqual(other.lookup('001234567890')[0]['record']['name'],'Nombre editado')

    def test_invalid_replace_does_not_damage_catalog(self):
        s = self.store([self.row()])
        with self.assertRaises(ValueError):
            s.replace([self.row('750'), self.row('bad')])
        self.assertEqual(len(s.lookup('001234567890')),1)

    def test_backup_portable_restore_and_auto_backup(self):
        s = self.store([self.row()])
        target = self.root/'copy.corrugados'
        s.export_backup(target)
        s.delete(s.list()[0]['id'])
        s.replace(catalog.read_backup(target))
        self.assertEqual(len(s.list()),1)
        self.assertTrue(list((self.root/'Respaldos').glob('*.corrugados')))
        target.write_text('{"format":"wrong","records":[]}',encoding='utf-8')
        with self.assertRaises(ValueError): catalog.read_backup(target)

    def test_excel_two_headers_mixed_types_and_zero_dv(self):
        w=Workbook(); a=w.active; a.title=' Código de Producto '
        a.append(['NUMERO BASE + NUMERO MOVIL','PRODUCTO','FAMILIA'])
        a.append([123456789012, 'Producto A', 'Reposado'])
        a.append(['000123','Con ceros',''])
        b=w.create_sheet(' Código de corrugado')
        b.append(['Texto de referencia']); b.append(['NUMERO BASE + NUMERO MOVIL','DV','PRODUCTO','SKU DE PRODUCTO','ITF14'])
        b.append(['0 123456 78901',0,'Caja litro','SKU-TEST-01','00123456789017'])
        p=self.root/'source.xlsx'; w.save(p); w.close()
        rows,issues=catalog.read_excel(p)
        self.assertEqual(len(rows),3)
        self.assertEqual(rows[0]['code'],'123456789012')
        self.assertEqual(rows[1]['code'],'000123')
        self.assertEqual(rows[2]['dv'],'0')
        self.assertEqual(rows[2]['itf14'],'00123456789017')
        self.assertEqual(rows[2]['source_row'],3)
        self.assertEqual(issues,[])

    def test_excel_bad_rows_reported_without_inventing_data(self):
        w=Workbook(); a=w.active
        a.append(['NUMERO BASE + NUMERO MOVIL','PRODUCTO'])
        a.append(['12X','Incorrecto']); a.append([None,'Sin código']); a.append(['123',None])
        p=self.root/'invalid.xlsx'; w.save(p);w.close()
        rows,issues=catalog.read_excel(p)
        self.assertEqual(len(rows),1)
        self.assertEqual(rows[0]['name'],'')
        self.assertEqual(len(issues),3)

    def test_import_rejects_scientific_text_and_unsafe_numeric_precision(self):
        for value in ('7.501e11', 1234567890123456, 123.5):
            with self.assertRaises(ValueError): catalog.normalize_code(value)

    def test_form_validation_rejects_bad_dv_without_losing_zero(self):
        s=self.store([self.row(dv='0')])
        self.assertEqual(s.list()[0]['dv'],'0')
        with self.assertRaises(ValueError): s.save(self.row('123',dv='10'))
        with self.assertRaises(ValueError): s.save(self.row('123',itf14='x'))

if __name__=='__main__': unittest.main(verbosity=2)
