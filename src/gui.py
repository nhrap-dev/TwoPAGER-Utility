import ctypes
import sys
import pyodbc as py
import os
import tkinter as tk
import pandas as pd
from tkinter import messagebox
from tkinter import filedialog
from tkinter import OptionMenu
from tkinter import StringVar
from tkinter import ttk
from tkinter import PhotoImage
from tkinter import Label
from tkinter import Canvas
from tkinter.ttk import Progressbar
from tkinter import TOP, RIGHT, LEFT, BOTTOM
from tkinter import N, S, E, W
from PIL import ImageTk, Image
from time import time, sleep
import json


class App():
    def __init__(self):
        # Create app
        self.root = tk.Tk()
        self.root.grid_propagate(0)

        # global styles
        config = json.loads(open('src/config.json').read())
        themeId = config['activeThemeId']
        theme = list(filter(
            lambda x: config['themes'][x]['themeId'] == themeId, config['themes']))[0]
        self.globalStyles = config['themes'][theme]['style']
        self.backgroundColor = self.globalStyles['backgroundColor']
        self.foregroundColor = self.globalStyles['foregroundColor']
        self.hoverColor = self.globalStyles['hoverColor']
        self.fontColor = self.globalStyles['fontColor']
        self.textEntryColor = self.globalStyles['textEntryColor']
        self.starColor = self.globalStyles['starColor']
        self.padl = 15
        # tk styles
        self.textBorderColor = self.globalStyles['textBorderColor']
        self.textHighlightColor = self.globalStyles['textHighlightColor']

        # ttk styles classes
        self.style = ttk.Style()
        self.style.configure("BW.TCheckbutton", foreground=self.fontColor,
                             background=self.backgroundColor, bordercolor=self.backgroundColor, side='LEFT')
        self.style.configure('TCombobox', background=self.backgroundColor, bordercolor=self.backgroundColor, relief='flat',
                             lightcolor=self.backgroundColor, darkcolor=self.backgroundColor, borderwidth=4, foreground=self.foregroundColor)

        # App parameters
        self.root.title('TwoPAGER Tool')
        self.root_h = 220
        self.root_w = 330
        self.root.geometry(str(self.root_w) + 'x' + str(self.root_h))
        windowWidth = self.root.winfo_reqwidth()
        windowHeight = self.root.winfo_reqheight()
        # Gets both half the screen width/height and window width/height
        positionRight = int(self.root.winfo_screenwidth()/2 - windowWidth)
        positionDown = int(self.root.winfo_screenheight()/3 - windowHeight)
        # Positions the window in the center of the page.
        self.root.geometry("+{}+{}".format(positionRight, positionDown))
        self.root.resizable(0, 0)
        self.root.configure(background=self.backgroundColor,
                            highlightcolor='#fff')

        # App images
        self.root.wm_iconbitmap('src/assets/images/HazusHIcon.ico')
        self.img_data = ImageTk.PhotoImage(Image.open(
            "src/assets/images/data_blue.png").resize((20, 20), Image.BICUBIC))
        self.img_edit = ImageTk.PhotoImage(Image.open(
            "src/assets/images/edit_blue.png").resize((20, 20), Image.BICUBIC))
        self.img_folder = ImageTk.PhotoImage(Image.open(
            "src/assets/images/folder_icon.png").resize((20, 20), Image.BICUBIC))

        # Init dynamic row
        self.row = 0

    def getStudyRegions(self):
        """Gets all study region names imported into your local Hazus install

        Returns:
            studyRegions: list -- study region names
        """
        comp_name = os.environ['COMPUTERNAME']
        conn = py.connect('Driver=ODBC Driver 11 for SQL Server;SERVER=' +
                          comp_name + '\HAZUSPLUSSRVR; UID=SA;PWD=Gohazusplus_02')
        exclusionRows = ['master', 'tempdb', 'model',
                         'msdb', 'syHazus', 'CDMS', 'flTmpDB']
        cursor = conn.cursor()
        cursor.execute('SELECT [StateID] FROM [syHazus].[dbo].[syState]')
        for state in cursor:
            exclusionRows.append(state[0])
        cursor = conn.cursor()
        cursor.execute('SELECT * FROM sys.databases')
        studyRegions = []
        for row in cursor:
            if row[0] not in exclusionRows:
                studyRegions.append(row[0])
        studyRegions.sort(key=lambda x: x.lower())
        return studyRegions

    def setup(self, scenario_name, folder_path):
        # Make a folder in folder_path for scenario_name
        if not os.path.exists(folder_path + '\\' + scenario_name):
            os.makedirs(folder_path + '\\' + scenario_name)
        # Connect to Hazus SQL Server database where scenario results are stored
        comp_name = os.environ['COMPUTERNAME']
        cnxn = py.connect('Driver=ODBC Driver 11 for SQL Server;SERVER=' + comp_name +
                          '\HAZUSPLUSSRVR;DATABASE=' + scenario_name + ';UID=hazuspuser;PWD=Gohazusplus_02')
        return comp_name, cnxn

    def read_sql(self, comp_name, cnxn, scenario_name):
        # Select Hazus results from SQL Server scenario database
        sql_econ_loss = """SELECT Tract, SUM(ISNULL(BldgLoss, 0) + ISNULL(ContentLoss, 0) + ISNULL(InvLoss, 0) + ISNULL(RelocLoss, 0) +
        ISNULL(IncLoss, 0) + ISNULL(RentLoss, 0) + ISNULL(WageLoss, 0)) AS EconLoss
        FROM %s.dbo.[eqTractEconLoss] GROUP BY [eqTractEconLoss].Tract""" % scenario_name
        sql_county_fips = """SELECT Tract, CountyFips FROM %s.dbo.[hzTract]""" % scenario_name
        sql_demographics = """SELECT Tract, Population, Households FROM
        %s.dbo.[hzDemographicsT]""" % scenario_name
        sql_impact = """SELECT Tract, DebrisW, DebrisS, DisplacedHouseholds AS DisplHouse, ShortTermShelter AS Shelter
        FROM %s.dbo.[eqTract]""" % scenario_name
        sql_injury = """SELECT Tract, SUM(ISNULL(Level1Injury, 0)) AS Level1Injury, SUM(ISNULL(Level2Injury, 0)) AS Level2Injury,
        SUM(ISNULL(Level3Injury, 0)) AS Level3Injury, SUM(ISNULL(Level1Injury, 0) + ISNULL(Level2Injury, 0) + ISNULL(Level3Injury, 0)) AS NonFatal5p
        FROM %s.dbo.[eqTractCasOccup] WHERE CasTime = 'C' AND InOutTot = 'Tot' GROUP BY Tract""" % scenario_name
        sql_building_damage = """SELECT Tract, SUM(ISNULL(PDsNoneBC, 0)) As NoDamage, SUM(ISNULL(PDsSlightBC, 0) + ISNULL(PDsModerateBC, 0)) AS GreenTag,
        SUM(ISNULL(PDsExtensiveBC, 0)) AS YellowTag, SUM(ISNULL(PDsCompleteBC, 0)) AS RedTag FROM %s.dbo.[eqTractDmg] WHERE DmgMechType = 'STR' GROUP BY Tract""" % scenario_name
        sql_building_damage_occup = """SELECT Occupancy, SUM(ISNULL(PDsNoneBC, 0)) As NoDamage, SUM(ISNULL(PDsSlightBC, 0) + ISNULL(PDsModerateBC, 0)) AS GreenTag,
        SUM(ISNULL(PDsExtensiveBC, 0)) AS YellowTag, SUM(ISNULL(PDsCompleteBC, 0)) AS RedTag FROM %s.dbo.[eqTractDmg] WHERE DmgMechType = 'STR' GROUP BY Occupancy""" % scenario_name
        sql_building_damage_bldg_type = """SELECT eqBldgType, SUM(ISNULL(PDsNoneBC, 0)) As NoDamage, SUM(ISNULL(PDsSlightBC, 0) + ISNULL(PDsModerateBC, 0)) AS GreenTag,
        SUM(ISNULL(PDsExtensiveBC, 0)) AS YellowTag, SUM(ISNULL(PDsCompleteBC, 0)) AS RedTag FROM %s.dbo.[eqTractDmg] WHERE DmgMechType = 'STR' GROUP BY eqBldgType""" % scenario_name
        sql_tract_spatial = """SELECT Tract, Shape.STAsText() AS Shape FROM %s.dbo.[hzTract]""" % scenario_name
        # Group tables and queries into iterable
        hazus_results = {'econ_loss': sql_econ_loss, 'county_fips': sql_county_fips, 'demographics': sql_demographics, 'impact': sql_impact, 'injury': sql_injury,
                         'building_damage': sql_building_damage, 'building_damage_occup': sql_building_damage_occup, 'building_damage_bldg_type': sql_building_damage_bldg_type,
                         'tract_spatial': sql_tract_spatial}
        # Read Hazus result tables from SQL Server into dataframes with Tract ID as index
        hazus_results_df = {table: pd.read_sql(
            query, cnxn) for table, query in hazus_results.items()}
        for name, dataframe in hazus_results_df.items():
            if (name != 'building_damage_occup') and (name != 'building_damage_bldg_type'):
                dataframe.set_index('Tract', append=False, inplace=True)
        return hazus_results_df

    def results(self, hazus_results_df):
        # Join and group Hazus result dataframes into final TwoPAGER dataframes
        tract_results = hazus_results_df['county_fips'].join([hazus_results_df['econ_loss'], hazus_results_df['demographics'], hazus_results_df['impact'],
                                                              hazus_results_df['injury'], hazus_results_df['building_damage']])
        county_results = tract_results.groupby('CountyFips').sum()
        return tract_results, county_results

    def to_csv(self, hazus_results_df, tract_results, county_results, folder_path, scenario_name):
        # Export TwoPAGER dataframes to text files
        two_pager_df = {'building_damage_occup': hazus_results_df['building_damage_occup'], 'building_damage_bldg_type': hazus_results_df['building_damage_bldg_type'],
                        'tract_results': tract_results, 'county_results': county_results}
        for name, dataframe in two_pager_df.items():
            path = folder_path + '\\' + scenario_name + '\\' + name + '.csv'
            dataframe.to_csv(path)

    def two_pager(self, scenario_name, folder_path, ftp=False):
        comp_name, cnxn = self.setup(scenario_name, folder_path)
        hazus_results_df = self.read_sql(comp_name, cnxn, scenario_name)
        tract_results, county_results = self.results(hazus_results_df)
        self.to_csv(hazus_results_df, tract_results,
                    county_results, folder_path, scenario_name)
        # to_shp(folder_path, scenario_name, hazus_results_df, tract_results)
        # if ftp == True:
        #     credentials = to_ftp(folder_path, scenario_name)
        #     print('Hazus earthquake results available for download at: https://' +
        #           credentials['host'] + credentials['dir'] + '/' + scenario_name)
        print('Hazus earthquake results available locally at: ' +
              folder_path + '\\' + scenario_name)

    def browsefunc(self):
        self.output_directory = filedialog.askdirectory()
        self.input_studyRegion = self.dropdownMenu.get()
        self.text_outputDir.delete("1.0", 'end-1c')
        if len(self.input_studyRegion) > 0:
            self.text_outputDir.insert(
                "1.0", self.output_directory + '/' + self.input_studyRegion)
            self.root.update_idletasks()
        else:
            self.text_outputDir.insert("1.0", self.output_directory)
            self.root.update_idletasks()

    def on_field_change(self, index, value, op):
        try:
            self.input_studyRegion = self.dropdownMenu.get()
            self.output_directory = str(
                self.text_outputDir.get("1.0", 'end-1c'))
            check = self.input_studyRegion in self.output_directory
            if (len(self.output_directory) > 0) and (not check):
                self.output_directory = '/'.join(
                    self.output_directory.split('/')[0:-1])
                self.text_outputDir.delete('1.0', tk.END)
                self.text_outputDir.insert(
                    "1.0", self.output_directory + '/' + self.input_studyRegion)
            self.root.update_idletasks()
        except:
            pass

    def getTextFields(self):
        dict = {
            'output_directory': '/'.join(self.text_outputDir.get("1.0", 'end-1c').split('/')[0:-1])
        }
        return dict

    def focus_next_widget(self, event):
        event.widget.tk_focusNext().focus()
        return("break")

    def on_enter_dir(self, e):
        self.button_outputDir['background'] = self.hoverColor

    def on_leave_dir(self, e):
        self.button_outputDir['background'] = self.backgroundColor

    def on_enter_run(self, e):
        self.button_run['background'] = '#006b96'

    def on_leave_run(self, e):
        self.button_run['background'] = '#0078a9'

    def run(self):
        try:
            if (len(self.dropdownMenu.get()) > 0) and (len(self.text_outputDir.get("1.0", 'end-1c')) > 0):
                self.root.update()
                sleep(1)
                func_row = self.row
                t0 = time()
                ### RUN FUNC HERE ###
                self.two_pager(self.input_studyRegion, self.output_directory)
                print('Total elasped time: ' + str(time() - t0))
                tk.messagebox.showinfo("Hazus", "Success! Output files can be found at: " +
                                       self.output_directory + '/' + self.input_studyRegion)
            else:
                tk.messagebox.showwarning(
                    'Hazus', 'Make sure a study region and output directory have been selected')
        except:
            ctypes.windll.user32.MessageBoxW(
                None, u"Unable to open correctly: " + str(sys.exc_info()[1]), u'Hazus - Message', 0)

    def build_gui(self):
        # App body
        # Required label
        self.label_required1 = tk.Label(
            self.root, text='*', font='Helvetica 14 bold', background=self.backgroundColor, fg=self.starColor)
        self.label_required1.grid(row=self.row, column=0, padx=(
            self.padl, 0), pady=(20, 5), sticky=W)
        # Scenario name label
        self.label_scenarioName = tk.Label(
            self.root, text='Study Region', font='Helvetica 10 bold', background=self.backgroundColor, fg=self.fontColor)
        self.label_scenarioName.grid(
            row=self.row, column=1, padx=0, pady=(20, 5), sticky=W)
        self.row += 1

        # Get options for dropdown
        options = self.getStudyRegions()
        self.variable = StringVar(self.root)
        self.variable.set(options[0])  # default value

        # Scenario name dropdown menu
        self.v = StringVar()
        self.v.trace(W, self.on_field_change)
        self.dropdownMenu = ttk.Combobox(
            self.root, textvar=self.v, values=options, width=40, style='H.TCombobox')
        self.dropdownMenu.grid(row=self.row, column=1,
                               padx=(0, 0), pady=(0, 0), sticky=W)
        self.dropdownMenu.bind("<Tab>", self.focus_next_widget)

        # Scenario icon
        self.img_scenarioName = tk.Label(
            self.root, image=self.img_data, background=self.backgroundColor)
        self.img_scenarioName.grid(
            row=self.row, column=2, padx=0, pady=(0, 0), sticky=W)
        self.row += 1

        # Required label
        self.label_required3 = tk.Label(
            self.root, text='*', font='Helvetica 14 bold', background=self.backgroundColor, fg=self.starColor)
        self.label_required3.grid(row=self.row, column=0, padx=(
            self.padl, 0), pady=(10, 0), sticky=W)
        # Output directory label
        self.label_outputDir = tk.Label(self.root, text='Output Directory',
                                        font='Helvetica 10 bold', background=self.backgroundColor, fg=self.fontColor)
        self.label_outputDir.grid(
            row=self.row, column=1, padx=0, pady=(10, 0), sticky=W)
        self.row += 1

        # Output directory text form
        self.text_outputDir = tk.Text(self.root, height=1, width=37, font='Helvetica 10', background=self.textEntryColor,
                                      relief='flat', highlightbackground=self.textBorderColor, highlightthickness=1, highlightcolor=self.textHighlightColor)
        self.text_outputDir.grid(
            row=self.row, column=1, padx=(0, 0), pady=(0, 0), sticky=W)
        self.text_outputDir.bind("<Tab>", self.focus_next_widget)

        # Output directory browse button
        self.button_outputDir = tk.Button(self.root, text="", image=self.img_folder, command=self.browsefunc,
                                          relief='flat', background=self.backgroundColor, fg='#dfe8e8', cursor="hand2", font='Helvetica 8 bold')
        self.button_outputDir.grid(
            row=self.row, column=2, padx=0, pady=(0, 0), sticky=W)
        self.button_outputDir.bind("<Tab>", self.focus_next_widget)
        self.button_outputDir.bind("<Enter>", self.on_enter_dir)
        self.button_outputDir.bind("<Leave>", self.on_leave_dir)
        self.row += 1

        # Run button
        self.button_run = tk.Button(self.root, text='Run', width=5, command=self.run,
                                    background='#0078a9', fg='#fff', cursor="hand2", font='Helvetica 8 bold', relief='flat')
        self.button_run.grid(row=self.row, column=1, columnspan=1,
                             sticky='nsew', padx=50, pady=(30, 20))
        self.button_run.bind("<Enter>", self.on_enter_run)
        self.button_run.bind("<Leave>", self.on_leave_run)
        self.row += 1

    # Run app
    def start(self):
        self.build_gui()
        self.root.mainloop()


# Start the app
app = App()
app.start()
