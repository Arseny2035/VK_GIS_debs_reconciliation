"""
App is performing a check of occurrences of water utility subscribers
from of request list in the list of subscribers with debts which have
open collection cases.
Requests is getting from downloaded GIS (state system) Excel file.
Information about collection cases is getting from VK (enterprise) Excel file,
which filled by lawyers.
Match list is saved in Excel reportFile.
"""

import os.path
import pickle

from kivy.lang import Builder
from kivy.uix.image import Image
from kivy.clock import mainthread

from kivymd.uix.screen import MDScreen
from kivymd.app import MDApp
from kivymd.uix.button import MDFillRoundFlatButton, MDIconButton
from kivymd.uix.textfield import MDTextField
from kivymd.uix.toolbar import MDToolbar
from kivymd.uix.dialog import MDDialog
from kivymd.uix.button import MDFlatButton
from kivymd.uix.label import MDLabel
from kivymd.uix.boxlayout import BoxLayout

import numpy as np
import pandas as pd
import easygui

KV = '''
<Content_city>
    orientation: "vertical"
    spacing: "12dp"
    size_hint_y: None
    height: "120dp"
        
    MDTextField:
        id: city
        hint_text: "Введите название населенного пункта принятое в водоканале"
        
<Content_street>
    orientation: "vertical"
    spacing: "12dp"
    size_hint_y: None
    height: "120dp"
        
    MDTextField:
        id: street
        hint_text: "Введите название улицы принятое в водоканале"

MDFloatLayout:

'''

# Constants
DATA_FILE_NAME = 'Data/data.pkl'
DICT_FILE_NAME = 'Data/dict.xlsx'
IMAGE_FILE_NAME = 'Data/logo.jpg'


class Content_city(BoxLayout):
    pass


class Content_street(BoxLayout):
    pass


class CheckApp(MDApp):

    def __init__(self, cities_dict: dict, streets_dict: dict,
                 GISFileName: str, VKFileName: str, reportFilePath: str, baseLoaded: bool, *kw):
        super(CheckApp, self).__init__()
        self.GIS_FileName = GISFileName  # Path+filename of GIS requests file.
        self.VK_FileName = VKFileName  # Path+filename of VK data file.
        self.report_FilePath = reportFilePath  # Path for report file saving.
        self.cities_dict = cities_dict  # Dict of GIS/VK variants for cities names.
        self.streets_dict = streets_dict  # Dict of GIS/VK variants for streets names.
        self.baseLoaded = baseLoaded  # Showing error (False) if database wasn't loaded.

    def cancelAddCity(self, args):
        """Canceling dialog about adding new city."""
        self.labelCondition.theme_text_color = 'Error'
        self.labelCondition.text = 'Справочник городов не был обновлен корректно.'
        self.dialog.dismiss()
        self.dialog = None

    @mainthread
    def addCityAddress(self, *args):
        """Added new city in cities dicts, close adding dialog and restart threatment."""
        print('Добавляем нас пункт: ', self.dialog.content_cls.ids.city.text, ' вместо ', self.addCity)
        # Check that input textField with new name is not empty
        if len(self.dialog.content_cls.ids.city.text) > 0:
            self.cities_dict[self.addCity] = self.dialog.content_cls.ids.city.text
            self.waitMessageTreatment()
        else:
            self.labelCondition.theme_text_color = 'Error'
            self.labelCondition.text = 'Справочник городов не был обновлен корректно.'
        self.dialog.dismiss()
        self.dialog = None

    def cancelAddStreet(self, args):
        """Canceling dialog about adding new street."""
        self.labelCondition.theme_text_color = 'Error'
        self.labelCondition.text = 'Справочник улиц не был обновлен корректно.'
        self.dialog.dismiss()
        self.dialog = None

    @mainthread
    def addStreetAddress(self, *args):
        """Add new street in streets dicts, close adding dialog and restart treatment."""
        print('Добавляем улицу: ', self.dialog.content_cls.ids.street.text, ' вместо ', self.addStreet)
        # Check that input textField with new name is not empty
        if len(self.dialog.content_cls.ids.street.text) > 0:
            self.streets_dict[self.addStreet] = self.dialog.content_cls.ids.street.text
            self.treatment()
        else:
            self.labelCondition.theme_text_color = 'Error'
            self.labelCondition.text = 'Справочник улиц не был обновлен корректно.'
        self.dialog.dismiss()
        self.dialog = None

    def getGISAddresses(self, *args) -> list[str]:
        """Read GIS file and collect all GIS requests in list of hashed house addresses.
        Then we compare GIS hashes of addresses with VK hashes of addresses."""
        datas = pd.read_excel(self.GIS_FileName, sheet_name='Для поставщика')
        addresses = datas.pop('Адрес дома')
        addresses = list(addresses)

        # Create empty lists
        GIS_cities = [''] * len(addresses)
        GIS_streets = [''] * len(addresses)
        GIS_houses = [''] * len(addresses)

        # Load and treat address rows
        for i in range(len(addresses)):
            # Delete from address first text about zip code and region (same for all addresses),
            # and extract information for address fields
            GIS_cities[i], GIS_streets[i], GIS_houses[i] = getGISAddress(addresses[i])

        # Load and treat flats rows
        flats = datas.pop(' Номер квартиры, комнаты, блока жилого дома')
        flats = list(flats)

        GIS_flats = [''] * len(flats)
        for i in range(len(flats)):
            GIS_flats[i] = getGISFlats(flats[i])

        # Now we convert GIS named variant of cities and streets to VK named variant
        # (use dict from file "data.pkl")
        for i in range(len(addresses)):
            if GIS_cities[i] in self.cities_dict:
                GIS_cities[i] = self.cities_dict[GIS_cities[i]]
            else:
                # If don't find city name in dictionary, then start dialog to add new name
                print('Не найден нас. пункт: ', GIS_cities[i])
                self.addCity = GIS_cities[i]
                self.showDialogAddCity(self.addCity)
                return []

        for i in range(len(addresses)):
            if GIS_streets[i] in self.streets_dict:
                GIS_streets[i] = self.streets_dict[GIS_streets[i]]
            else:
                # If don't find street name in dictionary, then start dialog to add new name
                print('Не найдена улица: ', GIS_streets[i])
                self.addStreet = GIS_streets[i]
                self.showDialogAddStreet(self.addStreet)
                return []

        # Create as list of hashed list of address fiends.
        GIS_addresses = [''] * len(addresses)
        for i in range(len(addresses)):
            GIS_addresses[i] = hash(list[GIS_cities[i], GIS_streets[i], GIS_houses[i], GIS_flats[i]])

        print('getGISAddresses done')
        return GIS_addresses

    def waitMessageTreatment(self, *args):
        """Show wait panel and then starting treatment process (may take some time).
        After - hide wait panel."""
        if not self.dialog:
            self.dialog = MDDialog(
                title="Идет поиск совпадений, ожидайте...",
                type='custom',
                size_hint=(0.5, 0.5),
                pos_hint={'center_x': 0.5, 'center_y': 0.5},
            )
            self.dialog.open()

            self.treatment()

            # Close wait panel after treating
            self.dialog.dismiss()
            self.dialog = None

    @mainthread
    def treatment(self, *args):
        """Launch treatment process.
        Make list of matching GIS and VK hashes and saved it into report file."""
        # Get GIS and VK hashes lists
        self.GIS_addresses = self.getGISAddresses()
        self.VK_addresses = getVKAddresses(self.VK_FileNameField.text)

        # Starting treatment only if GIS_addresses not empty - we have not empty request
        if len(self.GIS_addresses) == 0:
            self.labelCondition.theme_text_color = "Error"
            self.labelCondition.text = "Запрос ГИС не обработан: данные не найдены."
        else:
            # Starting treatment only if VK_addresses is not empty - we have few data for compare
            if len(self.VK_addresses) == 0:
                self.labelCondition.theme_text_color = "Error"
                self.labelCondition.text = "Файл взысканий не содержит записей."
            else:
                # Make reconciliation between GIS and VK addresses
                self.found = reconciliation(self.GIS_addresses, self.VK_addresses)

                # If we found some reconciliations -> save a report with them in report file.
                # Reflect statement of reconciliations result on main page
                # Saving report file in any case, even if empty
                if len(self.found) > 0:
                    df = pd.DataFrame({'Cовпали': self.found})
                    self.report_FileName = os.path.join(self.report_FilePath, 'Найдены совпадения.xlsx')
                    df.to_excel(self.report_FileName, sheet_name='Отчет ГИС', index=False)
                    os.system(r"explorer.exe " + self.report_FilePathField.text)
                    self.labelCondition.theme_text_color = "Primary"
                    self.labelCondition.text = 'Совпадения найдены. Смотрите файл отчета ' \
                                               '"Найдены совпадения".'
                else:
                    df = pd.DataFrame({})
                    self.report_FileName = os.path.join(self.report_FilePath, 'Найдены совпадения.xlsx')
                    df.to_excel(self.report_FileName, sheet_name='Отчет ГИС', index=False)
                    self.labelCondition.theme_text_color = "Primary"
                    self.labelCondition.text = "Совпадения не найдены."

                # Updating dicts after work done (in case of adding new streets or cities)
                saveData(self.cities_dict, self.streets_dict,
                         self.GIS_FileNameField.text, self.VK_FileNameField.text,
                         self.report_FilePathField.text)

        return True

    def checkDataBaseCreateButtonCanBeEnabled(self):
        """If file with database is missing, or we want to reload dicts from excel so
        we deleted datafile, than we need the possibility to create datafile, but
        before this we need to ensure than the ways to files and report dir are registered
        in the appropriate fields.
        If registered and database from datafile is not already loaded - than button mast to
        be enabled."""
        if self.GIS_FileNameField.text != '' and self.VK_FileNameField.text != '' \
                and self.report_FilePathField.text != '' and self.baseLoaded == False:
            self.dataBaseCreate_Button.disabled = False
        else:
            self.dataBaseCreate_Button.disabled = True

    def changeGISFileName(self, args):
        """Change used GIS file by the fileOpenBox dialog."""
        filePathString = easygui.fileopenbox()
        if filePathString:
            self.GIS_FileName = filePathString
            self.GIS_FileNameField.text = filePathString
            self.checkDataBaseCreateButtonCanBeEnabled()

    def changeVKFileName(self, args):
        """Change used VK file by the fileOpenBox dialog."""
        filePathString = easygui.fileopenbox()
        if filePathString:
            self.VK_FileName = filePathString
            self.VK_FileNameField.text = filePathString
            self.checkDataBaseCreateButtonCanBeEnabled()

    def changeReportFilePath(self, args):
        """Change path to save reportFile by the dirOpenBox dialog."""
        filePathString = easygui.diropenbox()
        if filePathString:
            self.report_FilePath = filePathString
            self.report_FilePathField.text = filePathString
            self.checkDataBaseCreateButtonCanBeEnabled()

    def build(self):
        """Building the App interface."""

        self.theme_cls.primary_palette = 'Blue'
        screen = MDScreen()

        # Check - if database doesn't load, then demonstrate an error
        if self.baseLoaded:
            self.reportButtonDisabled = False
            self.conditionColor = 'Primary'
            self.conditionMessage = ''
        else:
            self.reportButtonDisabled = True
            self.conditionColor = 'Error'
            self.conditionMessage = 'База данных не найдена'

        self.dialog = None

        # TOOLBAR Header
        self.toolbar = MDToolbar(title="Сверка запросов ГИС по должникам (взыскания).")
        self.toolbar.pos_hint = {'top': 1}
        screen.add_widget(self.toolbar)

        # IMAGE Logo
        screen.add_widget(Image(
            source=IMAGE_FILE_NAME,
            pos_hint={'center_x': 0.15, 'center_y': 0.7},
            size_hint=(0.3, 0.3),
        ))

        # LABEL Instructions
        screen.add_widget(MDLabel(
            pos_hint={'center_x': 0.75, 'center_y': 0.7},
            # padding_y = 20,
            padding_x=50,
            text="1. Проверьте пути к файлам и отчету.\n\n"
                 "2. Нажмите «Провести сверку».\n\n"
                 "3. Требуемая информация находится в сгенерированном файле\n "
                 "excel «Найдены совпадения».\n",
            theme_text_color='Secondary',
        ))

        # LABEL Condition
        self.labelCondition = MDLabel(
            pos_hint={'center_x': 0.75, 'center_y': 0.55},
            # padding_y = 20,
            padding_x=50,
            text=self.conditionMessage,
            theme_text_color=self.conditionColor
        )
        screen.add_widget(self.labelCondition)

        # TEXT_FIELD GIS file address setup
        self.GIS_FileNameField = MDTextField(
            hint_text="Файл c запросом ГИС",
            text=self.GIS_FileName,
            halign='left',
            size_hint=(0.9, 1),
            pos_hint={'center_x': 0.48, 'center_y': 0.4},
            font_size=12
        )
        screen.add_widget(self.GIS_FileNameField)

        # ICON-BUTTON VK file
        screen.add_widget(MDIconButton(
            icon='folder-open',
            pos_hint={"center_x": 0.95, "center_y": 0.4},
            on_press=self.changeGISFileName
        ))

        # TEXT_FIELD VK file address setup
        self.VK_FileNameField = MDTextField(
            hint_text="Файл c данными водоканала",
            text=self.VK_FileName,
            halign='left',
            size_hint=(0.9, 1),
            pos_hint={'center_x': 0.48, 'center_y': 0.3},
            font_size=12
        )
        screen.add_widget(self.VK_FileNameField)

        # ICON-BUTTON VK file
        screen.add_widget(MDIconButton(
            icon='folder-open',
            pos_hint={"center_x": 0.95, "center_y": 0.3},
            on_press=self.changeVKFileName
        ))

        # TEXT_FIELD Report file save address setup
        self.report_FilePathField = MDTextField(
            hint_text="Место для сохранения отчета",
            text=self.report_FilePath,
            halign='left',
            size_hint=(0.9, 1),
            pos_hint={'center_x': 0.48, 'center_y': 0.2},
            font_size=12
        )
        screen.add_widget(self.report_FilePathField)

        # ICON-BUTTON Folder report
        screen.add_widget(MDIconButton(
            icon='folder-open',
            pos_hint={"center_x": 0.95, "center_y": 0.2},
            on_press=self.changeReportFilePath
        ))

        # BUTTON Report
        self.report_Button = MDFillRoundFlatButton(
            text="Провести сверку",
            font_size=17,
            pos_hint={'center_x': 0.15, 'center_y': 0.1},
            on_press=self.waitMessageTreatment,
            disabled=self.reportButtonDisabled
        )
        screen.add_widget(self.report_Button)

        # BUTTON Create DataBase
        self.dataBaseCreate_Button = MDFillRoundFlatButton(
            text="Создать словари",
            font_size=17,
            pos_hint={'center_x': 0.40, 'center_y': 0.1},
            on_press=self.mustCreateDataBase,
            # disabled=True
        )
        screen.add_widget(self.dataBaseCreate_Button)

        # BUTTON Exit
        screen.add_widget(MDFillRoundFlatButton(
            text="Выход",
            font_size=17,
            pos_hint={'center_x': 0.85, 'center_y': 0.1},
            on_press=self.showDialogExit
        ))

        # Loading KV file (with context to inputText dialogs)
        screen.add_widget(Builder.load_string(KV))

        # Run check function, and, maybe, enabling dataBaseCreate_Button
        self.checkDataBaseCreateButtonCanBeEnabled()

        return screen

    def dialogClose(self, *args):
        """Close dialog (most likely for Exit dialog)."""
        self.dialog.dismiss()

    def saveAndExit(self, *args):
        """If everything going well, and we're closing the App, then we must save possible changes."""
        if self.baseLoaded:
            # If reportButton is enabled, then all dict are loaded well, and we can save
            # information in database
            saveData(self.cities_dict, self.streets_dict,
                     self.GIS_FileNameField.text, self.VK_FileNameField.text,
                     self.report_FilePathField.text)
        self.stop()

    def mustCreateDataBase(self, args):
        """Starting processes of creating database from excel dicts and save in with current paths."""
        print('Начинаем создание словарей')
        if os.path.exists(DICT_FILE_NAME):
            self.cities_dict, self.streets_dict, dictsLoaded = createDataBase(self.GIS_FileName, self.VK_FileName,
                                                                              self.report_FilePath)
            if dictsLoaded:
                Application.labelCondition.theme_text_color = 'Primary'
                Application.labelCondition.text = 'Словари успешно созданы!'
            else:
                Application.labelCondition.theme_text_color = 'Error'
                Application.labelCondition.text = 'Ошибка формирования словарей!'
            self.dataBaseCreate_Button.disabled = True  # If data is loaded, button isn't needs
            self.report_Button.disabled = False  # If data is loaded, we can start generate report
        else:
            # If file Excel with dicts wasn't found
            Application.labelCondition.theme_text_color = 'Error'
            Application.labelCondition.text = 'Файл со словарями "', DICT_FILE_NAME, '" не найден!'

    @mainthread
    def showDialogAddCity(self, city):
        """DIALOG To add new city. Show in case when next city (from GIS request) not in dicts."""
        if not self.dialog:
            self.dialog = MDDialog(
                title="Добавить город: " + city,
                type='custom',
                size_hint=(0.5, 0.5),
                pos_hint={'center_x': 0.5, 'center_y': 0.5},
                content_cls=Content_city(),
                buttons=[
                    MDFlatButton(
                        text="Добавить город ",
                        theme_text_color="Custom",
                        text_color=self.theme_cls.primary_color,
                        on_press=self.addCityAddress
                    ),
                    MDFlatButton(
                        text="Отмена",
                        theme_text_color="Custom",
                        text_color=self.theme_cls.primary_color,
                        on_press=self.cancelAddCity
                    )
                ]
            )
            self.dialog.open()

    def showDialogAddStreet(self, street):
        """DIALOG To add new street. Show in case when next street (from GIS request) not in dicts."""
        if not self.dialog:
            self.dialog = MDDialog(
                title="Добавить улицу: " + street,
                type='custom',
                size_hint=(0.5, 0.5),
                pos_hint={'center_x': 0.5, 'center_y': 0.5},
                content_cls=Content_street(),
                buttons=[
                    MDFlatButton(
                        text="Добавить улицу",
                        theme_text_color="Custom",
                        text_color=self.theme_cls.primary_color,
                        on_press=self.addStreetAddress
                    ),
                    MDFlatButton(
                        text="Отмена",
                        theme_text_color="Custom",
                        text_color=self.theme_cls.primary_color,
                        on_press=self.cancelAddStreet
                    )
                ]
            )
            self.dialog.open()

    def showDialogExit(self, *args):
        """DIALOG Show whe buttonExit pressed."""
        if not self.dialog:
            self.dialog = MDDialog(
                text="Выйти из приложения?",
                buttons=[
                    MDFlatButton(
                        text="Выйти",
                        theme_text_color="Custom",
                        text_color=self.theme_cls.primary_color,
                        on_press=self.saveAndExit  # Save data and exit
                    ),
                    MDFlatButton(
                        text="Отмена",
                        theme_text_color="Custom",
                        text_color=self.theme_cls.primary_color,
                        on_press=self.dialogClose
                    ),
                ],
            )
        self.dialog.open()


def saveData(cities_dict: dict, streets_dict: dict, GISFileName: str, VKFileName: str, reportFilePath: str):
    """ Save data (filenames, path, dicts) in pickle dataFile."""
    addresses = [cities_dict, streets_dict, [GISFileName, VKFileName, reportFilePath]]
    dict_file = open(DATA_FILE_NAME, 'wb')
    pickle.dump(addresses, dict_file)
    dict_file.close()
    print('Данные сохранены.')


def loadData(filename: str) -> [dict, dict, str, str, str, bool]:
    """Loading data from dataFile. It contains: cities dict, streets dict, """
    if os.path.exists(filename):
        dataFile = open(filename, 'rb')
        # Load info from previously saved dataFile
        data = pickle.load(dataFile)
    else:
        print('Файл базы данных не найден!')
        cities_dict = []
        streets_dict = []
        GISFileName = ''
        VKFileName = ''
        reportFilePath = ''
        baseLoaded = False
        return cities_dict, streets_dict, GISFileName, VKFileName, reportFilePath, baseLoaded

    # Dicts of cities and streets it is dict witch include ['GIS_name_variant' : 'VK_name_variant'].
    # We want to check coincidences between GIS and VK lists, so first we need to
    # convert GIS named variant addresses to VK named variant addresses.
    # Load dict of cities names (GIS-VK)
    cities_dict = data[0]
    # Load dict of streets names (GIS-VK)
    streets_dict = data[1]

    # Load paths and filenames to files with requests, and to dir with generated report
    GISFileName, VKFileName, reportFilePath = data[2]
    baseLoaded = True
    return cities_dict, streets_dict, GISFileName, VKFileName, reportFilePath, baseLoaded


def reconciliation(GISRequestAddresses: list, VKaddresses: dict) -> list:
    """Mare a list of results of comparing GIS and VK hashes."""
    found = []
    for check in GISRequestAddresses:
        if VKaddresses.get(check):
            print('Найден: ', VKaddresses[check])
            found.append(VKaddresses[check])
    print('reconciliation done')
    found = sorted(found)
    return found


def getGISHouseNumber(houseNumber: str) -> str:
    """Convert house number "47, к. 5" (example) to "47/5".
    Convert house number "47a" (example) to "47/A"."""
    houseNumber = houseNumber[4:].replace(' к. ', '/').upper()
    if ('/' not in houseNumber) and (not houseNumber.isdigit()):
        for i in range(1, len(houseNumber)):
            if not houseNumber[i].isdigit():
                houseNumber = houseNumber[:i] + '/' + houseNumber[i:]
                break
    return houseNumber


def getGISAddress(address: str) -> list[str, str, str]:
    """Convert string "352690, край Краснодарский, р-н Апшеронский, ст Тверская, ул Центральная, д. 6а"
    to [GIS_city, GIS_street, GIS_house] = [ст Тверская, ул Центральная, 6/А]."""
    address = address[45:]
    parts = address.split(',')
    GIS_city = parts[0].strip()
    GIS_street = parts[1].strip()
    GIS_house = getGISHouseNumber(''.join(parts[2:]))
    return [GIS_city, GIS_street, GIS_house]


def getGISFlats(flat: str) -> str:
    """Convert 'кв. 3а' to '3A'. Decided to don't use slash (/) in flats numbers."""
    if (len(flat) > 0) and (flat[:3] == 'кв.'):
        flat = flat[3:].upper()
    else:
        flat = ''
    return flat


def putAwayDotZeros(tempList: list):
    """Convert '26.0' to '26' after converting to string."""
    for i in range(len(tempList)):
        try:
            tempList[i] = str(tempList[i])
            if (tempList[i][-2:]) == '.0':
                tempList[i] = tempList[i][:(len(tempList[i]) - 2)]
        except:
            pass
    return tempList


def getVKAddresses(filename: str) -> dict:
    """Read VK file and collect all VK data lines from matched sheets to generate a list
    of hashed addresses. Then we would compare GIS hashes of addresses with VK hashes of addresses."""
    VKfile = pd.read_excel(filename, None)  # Read VK file

    VK_cities = []
    VK_streets = []
    VK_houses = []
    VK_flats = []
    VK_enable = []
    VK_sheet = []

    for sheet in VKfile.keys():
        if ('приказы' in sheet) or ('иски' in sheet):  # Choose only corrected sheets
            datas = pd.read_excel(filename, sheet_name=sheet, header=0)
            temporary = datas.pop('Населенный пункт').replace(np.nan, '').tolist()

            # Get the length from city cols and crop others be same length = main_length
            main_length = len(temporary)

            VK_cities.append(temporary)

            temporary = datas.pop('Улица').replace(np.nan, '').tolist()[:main_length]
            VK_streets.append(temporary)

            temporary = []
            col1 = datas.pop('Дом').replace(np.nan, '').tolist()[:main_length]
            col1 = putAwayDotZeros(col1)
            col2 = datas.pop('Литер / дробь').replace(np.nan, '').tolist()[:main_length]
            col2 = putAwayDotZeros(col2)
            # If col2 (litter for house number) != 0, then concatenate both cols in one house number
            # with slash inside (for example: col1: "122", col2: "A", house number: "122/A")
            for i in range(len(col1)):
                if col2[i] != '':
                    temporary.append(col1[i] + '/' + col2[i])
                else:
                    temporary.append(col1[i])
            VK_houses.append(temporary)

            temporary = datas.pop('Кварт.').replace(np.nan, '').tolist()[:main_length]
            VK_flats.append(temporary)

            # If lawyers want us to skip some rows with address, then they marked this row by "Нет"
            # in col "Включать в сверку ГИС-АБО". Get this row too before checking.
            temporary = datas.pop('Включать в сверку ГИС-АБО').replace(np.nan, '').tolist()[:main_length]
            VK_enable.append(temporary)

            # Adding current sheet name to each row getted from this sheet.
            temporary = [sheet] * main_length
            VK_sheet.append(temporary)

    # Concatenate all nested lists in one. For example: [[3, 4], [5, 6]] -> [3, 4, 5, 6]
    VK_cities = sum(VK_cities, [])
    VK_streets = sum(VK_streets, [])
    VK_houses = sum(VK_houses, [])
    VK_flats = sum(VK_flats, [])
    VK_enable = sum(VK_enable, [])
    VK_sheet = sum(VK_sheet, [])

    # Convert "20.0" to "20" for all cases, including str, int, float and etc.
    VK_flats = putAwayDotZeros(VK_flats)

    for i in range(len(VK_houses)):
        VK_houses[i] = VK_houses[i].upper()  # Making all non digits in uppercase

    for i in range(len(VK_flats)):
        VK_flats[i] = VK_flats[i].upper()  # Making all non digits in uppercase

    # Create empty dict for finally addresses {'hash number' : 'full address in string'}
    VK_addresses = dict()

    for i in range(len(VK_cities)):
        # Check that city name is not empty, and skip cell dosn't have 'Нет' value.
        if (VK_cities[i] != '') and (VK_enable[i] != 'Нет'):
            # If flat number is empty, than we don't concatenate all addresses with flats
            if VK_flats[i] != '':
                address_string = 'Лист "{}": {}, {}, д. {}, кв. {}'. \
                    format(VK_sheet[i], VK_cities[i], VK_streets[i], VK_houses[i], VK_flats[i])
            else:
                address_string = 'Лист "{}": {}, {}, д. {}'. \
                    format(VK_sheet[i], VK_cities[i], VK_streets[i], VK_houses[i])
            # Generate hash value for current address and added new record in dict
            VK_addresses[hash(list[VK_cities[i], VK_streets[i], VK_houses[i], VK_flats[i]])] = address_string

    print('getVKaddresses done')
    return VK_addresses


def loadDict(filename: str, dictType: str) -> dict:
    """Create dict of dictType for GIS-VK variants via uploading these variants from Excel dict file ."""
    datas = pd.read_excel(filename, sheet_name=dictType, header=0)
    GIS_variant = datas.pop(0)
    VK_variant = datas.pop(1)
    variants = {}
    for i in range(len(GIS_variant)):
        variants[GIS_variant[i]] = VK_variant[i]
    return variants


def loadCitiesAndStreetsDicts(filename: str) -> [dict, dict, bool]:
    """Initialize dicts of cities and streets loading."""
    try:
        cities_dict = loadDict(filename, 'cities')
        print('Словарь Cities загружен. Длина ' + str(len(cities_dict)))
        streets_dict = loadDict(filename, 'streets')
        print('Словарь Streets загружен. Длина ' + str(len(streets_dict)))
        return [cities_dict, streets_dict, True]
    except:
        return [{}, {}, False]


def createDataBase(GISFileName: str, VKFileName: str, reportFilePath: str) -> [dict, dict, bool]:
    """Initialize dataBase creating from loaded dicts and paths from textInputs."""
    print('Пытаемся загрузить cities и streets')
    cities_dict, streets_dict, dicts_loaded = loadCitiesAndStreetsDicts(DICT_FILE_NAME)
    if dicts_loaded:
        print('Начинаем сохранять данные.')
        # Saving loaded dicts and paths.
        saveData(cities_dict, streets_dict, GISFileName, VKFileName, reportFilePath)
        return [cities_dict, streets_dict, True]
    # Return empty dicts and False if dicts doesn't load.
    return [{}, {}, False]


# Load dicts of cities and streets from file.
# Dicts is about ['GIS_name_variant' : 'VK_name_variant'].
_cities_dict, _streets_dict, _GisFile, _VKFile, _reportFileDir, _baseLoaded = loadData(DATA_FILE_NAME)

if __name__ == "__main__":
    Application = CheckApp(_cities_dict, _streets_dict, _GisFile, _VKFile, _reportFileDir, _baseLoaded)
    Application.run()
