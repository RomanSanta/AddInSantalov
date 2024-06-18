# -*- coding: utf-8 -*-
import clr, shutil, os# ДЛЯ РАБОТЫ С ФАЙЛАМИ
clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')
import sys
from pyrevit import revit, forms, DB # ОСНОВНЫЕ БИБЛИОТЕКИ REVIT API
from Autodesk.Revit.UI import RevitCommandId, PostableCommand # БИБЛИОТЕКИ КОМАНД В REVIT
from System.Collections.Generic import List # ДЛЯ ОСОБОГО СПИСКА
from System.Windows.Forms import (Application, Form, Button, ListBox, Label, FolderBrowserDialog, DialogResult,
                                  OpenFileDialog)
from System.Drawing import Size, Point, Font, FontStyle

uidoc = __revit__.ActiveUIDocument
path = 'C:/CollisionsRev/forRev.html' # ПУТЬ К ФАЙЛУ ПО УМОЛЧАНИЮ


def is_window_open(window_type):
    for form in Application.OpenForms:
        if form.GetType() == window_type:
            return True
    return False


def parse_ids(path):
    f = open(path).read()
    li = f.replace('\xd0\xbe\xd0\xb1\xd1\x8a\xd0\xb5\xd0\xba\xd1\x82\xd0\xb0</i>', 'NEED').split()
    odd = 1
    ids = []
    for i in range(len(li)):
        if li[i] == 'NEED':
            if odd % 2 == 1:
                a = li[i + 2].replace('</th>', '')
                ids.append(int(a))
            odd += 1
    return ids


def choose_directory():
    selected_path = forms.pick_file()
    if selected_path:
        source_f = selected_path
        destination_f = path
        shutil.copyfile(source_f, destination_f)

def make_sel_box(el_id):
    collection = List[DB.ElementId]([DB.ElementId(el_id)])
    selection_box = RevitCommandId.LookupPostableCommandId(PostableCommand.SelectionBox)
    with revit.Transaction():
        revit.uidoc.Selection.SetElementIds(collection)
        revit.uidoc.Application.PostCommand(selection_box)


class MyForm(Form):
    def __init__(self):
        self.Text = "Показать срез"
        self.ClientSize = Size(400, 700)

        self.label = Label()
        self.label.Location = Point(50, 20)
        self.label.Size = Size(300, 20)
        self.label.Text = u"Выберите ID объекта из списка:"
        self.label.Font = Font("Times New Roman", 12, FontStyle.Bold)

        self.listbox = ListBox()
        self.listbox.Location = Point(50, 50)
        self.listbox.Size = Size(300, 600)
        self.listbox.Font = Font("Times New Roman", 10)

        self.select_button = Button()
        self.select_button.Location = Point(50, 650)
        self.select_button.Size = Size(100, 30)
        self.select_button.Text = u"Выбрать"
        self.select_button.Click += self.select_button_click

        self.choose_button = Button()
        self.choose_button.Location = Point(200, 650)
        self.choose_button.Size = Size(150, 30)
        self.choose_button.Text = u"Указать путь файла"
        self.choose_button.Click += self.choose_button_click

        self.Controls.Add(self.label)
        self.Controls.Add(self.listbox)
        self.Controls.Add(self.select_button)
        self.Controls.Add(self.choose_button)

    def select_button_click(self, sender, event):
        selected_item = self.listbox.SelectedItem
        if selected_item:
            make_sel_box(selected_item)
            self.Close()

    def choose_button_click(self, sender, event):
        file_dialog = OpenFileDialog()
        result = file_dialog.ShowDialog()
        if result == DialogResult.OK:
            selected_file = file_dialog.FileName
            source_f = selected_file
            destination_f = path
            if os.path.exists(path):
                os.remove(path)
            shutil.copyfile(source_f, destination_f)
            self.Close()


app = MyForm()
if os.path.exists(path):
    items = parse_ids(path)
    for i in items:
        app.listbox.Items.Add(i)

if is_window_open(MyForm):
    sys.exit()
else:
    if __name__ == "__main__":
        Application.Run(app)
    # if os.path.exists(path):
    #     pass
    # else:
    #     choose_directory()
    #
    # selected_item = forms.SelectFromList.show(['CHOOSE A NEW DIRECTORY'] + items, title='Select the object id',
    # multiselect=False)
    #
    # if selected_item == 'CHOOSE A NEW DIRECTORY':
    #     choose_directory()
    # else:
    #     make_sel_box(int(selected_item))
