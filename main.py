from kivymd.app import MDApp
from plyer import filechooser
from datetime import date

from production_bom import production_bom


class MainApp(MDApp):
    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        self.bom_production_path = ''
        self.bom_path = ''
        self.today = date.today().strftime('%d/%m')

    def build(self):
        self.theme_cls.theme_style = 'Dark'
        self.theme_cls.primary_palette = 'Blue'
        pass

    def set_bom_production_path(self):
        self.bom_production_path = filechooser.open_file(title='choose production bom list',
                                                         filters=[('Excel file', '*.xlsx')])[0]
        self.root.ids.lab_prod_path.text = self.bom_production_path

    def set_bom_path(self):
        self.bom_path = filechooser.open_file(title='choose main bom list', filters=[('Excel file', '*.xlsx')])[0]
        self.root.ids.lab_bom_path.text = self.bom_path

    def update_bom(self):
        production_bom(self.bom_production_path,
                       self.bom_path,
                       supplier=self.root.ids.PartSupplier.text,
                       status=self.root.ids.Status.text,
                       besteldatum=self.root.ids.Besteldatum.text,
                       leveringsdatum=self.root.ids.Leveringsdatum.text)
        self.root.ids.run_button.text = 'BOM list updated!'


if __name__ == '__main__':
    MainApp().run()
