import unittest,tempfile
from pathlib import Path
try:
    import app
except ModuleNotFoundError:
    app=None
import catalog

class InterfaceTests(unittest.TestCase):
    def setUp(self):
        self.assertIsNotNone(app,'Falta la interfaz del escáner')
        self.tmp=tempfile.TemporaryDirectory();self.addCleanup(self.tmp.cleanup)
        self.ui=app.App(Path(self.tmp.name));self.addCleanup(self.ui.shutdown)
        self.ui.store.replace([{'code':'12345','name':'Caja litro','category':'Corrugado'}, {'code':'99','name':'Botella A','category':'Producto'}, {'code':'99','name':'Botella B','category':'Producto'}])
        self.ui.update()

    def test_scan_unknown_clears_previous_product(self):
        self.ui.scan_var.set('12345');self.ui.perform_scan();self.ui.update()
        self.assertEqual(self.ui.current_record['name'],'Caja litro')
        self.ui.scan_var.set('99999');self.ui.perform_scan();self.ui.update()
        self.assertIsNone(self.ui.current_record)
        self.assertIn('no registrado',self.ui.status_var.get().lower())

    def test_shared_code_requires_selection(self):
        self.ui.scan_var.set('99');self.ui.perform_scan();self.ui.update()
        self.assertIsNone(self.ui.current_record)
        self.assertEqual(len(self.ui.result_tree.get_children()),2)
        self.assertIn('compartido',self.ui.status_var.get().lower())

    def test_enter_scans_and_next_input_clears_result(self):
        self.ui.scan_var.set('12345');self.ui.scan_entry.focus_force();self.ui.update()
        self.ui.scan_entry.event_generate('<Return>');self.ui.update()
        self.assertEqual(self.ui.current_record['code'],'12345')
        self.ui.scan_var.set('9');self.ui.update()
        self.assertIsNone(self.ui.current_record)

    def test_catalog_filter_and_edit_save(self):
        self.ui.show_catalog();self.ui.filter_var.set('Caja');self.ui.refresh_catalog();self.ui.update()
        self.assertEqual(len(self.ui.catalog_tree.get_children()),1)
        row=self.ui.store.lookup('12345')[0]['record']
        edit=app.Editor(self.ui,row);self.ui.update()
        edit.variables['name'].set('Caja modificada');edit.commit();self.ui.update()
        self.assertEqual(self.ui.store.lookup('12345')[0]['record']['name'],'Caja modificada')

    def test_new_record_adds_without_replacing_existing(self):
        edit=app.Editor(self.ui);self.ui.update()
        edit.variables['code'].set('0000123456');edit.variables['name'].set('Nuevo producto');edit.commit();self.ui.update()
        self.assertEqual(len(self.ui.store.list()),4)
        self.assertEqual(self.ui.store.lookup('0000123456')[0]['record']['name'],'Nuevo producto')

    def test_new_installation_starts_with_empty_catalog(self):
        folder=Path(self.tmp.name)/'new-installation'
        second=app.App(folder)
        try:self.assertEqual(second.store.list(), [])
        finally:second.shutdown()

if __name__=='__main__':unittest.main(verbosity=2)
