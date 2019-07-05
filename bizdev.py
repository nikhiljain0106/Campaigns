#!/usr/bin/env python

#libraries for data manipulation and analysis
import pandas as pd
import numpy as np

#libraries to access and move files in the computer
import glob,os
import shutil

#library to access current date and time
import datetime

#library to design GUI (Graphical user interace)
import tkinter as tk
from tkinter import messagebox
import pandastable as pt
from tkinter import ttk
from firstnames import names_dict

dirname = os.path.dirname(__file__)

#--------------------------------------
#         CODE STARTS HERE
#--------------------------------------

class MainWindow:

    dirname = os.path.dirname(__file__)

    #----------------------------------
    #       Main window
    #----------------------------------

    # Class MainWindow is used to describe the first window that opens. There will never be a problem in this part of the code.

    def __init__(self, master):

        self.master = master
        self.FRAME = tk.Frame(self.master)
        
        HEIGHT = 750
        WIDTH = 500

        self.canvas = tk.Canvas(self.master, height = HEIGHT, width = WIDTH)
        self.canvas.pack()
        self.FRAME = tk.Frame(self.master, bd = 5)


        #----------------------------------
        #       Header Image
        #----------------------------------

        my_path = dirname + "/logo_keravalon.png"

        img = tk.PhotoImage(file=my_path)
        self.LABEL = tk.Label(self.FRAME, image = img)
        self.LABEL.image = img
        self.LABEL.place(relx= 0.5, rely = 0, relheight = 0.1, anchor = 'n')


        self.LABEL = tk.Label(self.FRAME, text="Keravalon BizDev", font = "Helvetica 20 bold")
        self.LABEL.place(relx = 0.5, rely = 0.2, anchor = 'n')

        

        self.B = tk.Button(self.FRAME, text ="Campaigns", command = self.new_window_campaign)
        self.C = tk.Button(self.FRAME, text ="Analysis", command = self.new_window_analysis)
        self.N = tk.Button(self.FRAME, text ="Newsletter", command = self.new_window_newsletter)
        self.M = tk.Button(self.FRAME, text ="Manual Contact", command = self.new_window_mancontact)
        
        self.B.place(relx = 0.5, rely = 0.35, relwidth = 0.4, anchor = 'n')
        self.C.place(relx = 0.5, rely = 0.45, relwidth = 0.4, anchor = 'n')
        self.N.place(relx = 0.5, rely = 0.55, relwidth = 0.4, anchor = 'n')
        self.M.place(relx = 0.5, rely = 0.65, relwidth = 0.4, anchor = 'n')

        self.FRAME.place(relheight = 1, relwidth = 1)

    def new_window_campaign(self):
        self.newWindow = tk.Toplevel(self.master)
        self.app = Campaigns(self.newWindow)

    def new_window_newsletter(self):
        self.newWindow2 = tk.Toplevel(self.master)
        self.app = Newsletter(self.newWindow2)

    def new_window_mancontact(self):
        self.newWindow3 = tk.Toplevel(self.master)
        self.app = ManualContact(self.newWindow3)
        
    def new_window_analysis(self):
        self.newWindow4 = tk.Toplevel(self.master)
        self.app = Analysis(self.newWindow4)

class Campaigns:

    dirname = os.path.dirname(__file__)

    # init function is used to describe the GUI (Graphical user interface)

    def __init__(self, master):

        #----------------------------------
        #      GUI
        #----------------------------------

        # This function is used to describe the entire graphical user interface. There will most likely never be a problem in this part of the code.

        self.master = master
        HEIGHT = 750
        WIDTH = 500

        self.canvas = tk.Canvas(self.master, height = HEIGHT, width = WIDTH)
        self.canvas.pack()
        self.FRAME = tk.Frame(self.master)
        self.FRAME.place(relwidth=1,relheight=1)

        my_path = dirname + "/logo_keravalon.png"

        img = tk.PhotoImage(file=my_path)
        self.LABEL = tk.Label(self.FRAME, image = img)
        self.LABEL.image = img
        self.LABEL.place(relx= 0.5, rely = 0, relheight = 0.1, anchor = 'n')

        #-------------------------------
        # UPDATE BUTTONS definition
        #-------------------------------        
        
        self.Init = tk.Button(self.FRAME, text ="Initialize", command = lambda: self.initialize())
        self.Skr = tk.Button(self.FRAME, text ="Update Skrapp", command = lambda: self.update_skrapp(self.main_df))
        self.Hub = tk.Button(self.FRAME, text ="Update HubSpot", command = lambda: self.update_hubspot(self.main_df))
        self.Wood = tk.Button(self.FRAME, text ="Update Woodpecker", command = lambda: self.update_woodpecker(self.main_df))

        #-----------------------------------
        # Checkbox, radio etc. declaration
        #-----------------------------------

        self.region = tk.Label(self.FRAME, text = "Select Regions:", font = "bold")
        self.last_contact = tk.Label(self.FRAME, text = "Last Contact Date:", font = "bold")

        self.contacts = tk.Label(self.FRAME, text = "No. Of Contacts found:", font = "bold")
        self.num_contacts_NB = tk.Label(self.FRAME, text = "0", bg = '#f4f1f1')
        self.num_contacts_WP = tk.Label(self.FRAME, text = "0", bg = '#f4f1f1')

        self.label_contacts_NB = tk.Label(self.FRAME, text = "NB")
        self.label_contacts_WP = tk.Label(self.FRAME, text = "WP")

        var1 = tk.IntVar()
        self.c1 = tk.Checkbutton(master, text="Spanish South America", variable=var1)
        var2 = tk.IntVar()
        self.c2 = tk.Checkbutton(master, text="Non Spanish South America", variable=var2)
        var3 = tk.IntVar()
        self.c3 = tk.Checkbutton(master, text="North America", variable=var3)
        var4 = tk.IntVar()
        self.c4 = tk.Checkbutton(master, text="France & Belgium", variable=var4)
        var5 = tk.IntVar()
        self.c5 = tk.Checkbutton(master, text="DACH", variable=var5)
        var6 = tk.IntVar()
        self.c6 = tk.Checkbutton(master, text="Spain", variable=var6)
        var7 = tk.IntVar()
        self.c7 = tk.Checkbutton(master, text="Italy", variable=var7)
        var8 = tk.IntVar()
        self.c8 = tk.Checkbutton(master, text="Other Western Europe")
        var9 = tk.IntVar()
        self.c9 = tk.Checkbutton(master, text="Rest of the world")

        var_contactdate = tk.IntVar()

        self.R1 = tk.Radiobutton(self.FRAME, text="Never", variable=var_contactdate, value=1)

        self.R2 = tk.Radiobutton(self.FRAME, text="Last ___ months", variable=var_contactdate, value=2)

        self.month_var = tk.Entry(self.FRAME)

        self.click_ok = tk.Button(self.FRAME, text ="OK", command = lambda: self.no_of_contacts(self.main_df, [var1.get(),var2.get(),var3.get(),var4.get(),var5.get(),var6.get(),var7.get(),var8.get(),var9.get()], self.find_months(var_contactdate.get(), self.month_var.get()), self.domlist))

        self.Ex_NB = tk.Button(self.FRAME, text ="Export for NeverBounce", command = lambda: self.export_neverbounce(self.main_df, [var1.get(),var2.get(),var3.get(),var4.get(),var5.get(),var6.get(),var7.get(),var8.get(),var9.get()], self.find_months(var_contactdate.get(), self.month_var.get()), self.domlist))
        self.NB = tk.Button(self.FRAME, text ="Update NeverBounce", command = lambda: self.update_neverbounce(self.main_df))
        self.Ex_WP = tk.Button(self.FRAME, text ="Export for Woodpeceker", command = lambda: self.export_woodpecker(self.main_df, [var1.get(),var2.get(),var3.get(),var4.get(),var5.get(),var6.get(),var7.get(),var8.get(),var9.get()], self.find_months(var_contactdate.get(), self.month_var.get()), self.domlist))


        #-------------------------------
        # UPDATE BUTTONS placement
        #-------------------------------
        self.Init.place(relwidth = 0.8, relx= 0.5, rely = 0.16, anchor = 'n')
        self.Skr.place(relwidth = 0.8, relx= 0.5, rely = 0.2, anchor = 'n')
        self.Hub.place(relwidth = 0.8, relx= 0.5, rely = 0.24, anchor = 'n')
        self.Wood.place(relwidth = 0.8, relx= 0.5, rely = 0.28, anchor = 'n')
        self.region.place(relx= 0, rely = 0.37)

        self.last_contact.place(relx= 0.6, rely = 0.37)

        #-------------------------------
        # REGION CHECKBUTTONS placement
        #-------------------------------
        self.c1.place(relx= 0.01, rely = 0.42)
        self.c2.place(relx= 0.01, rely = 0.45)
        self.c3.place(relx= 0.01, rely = 0.48)
        self.c4.place(relx= 0.01, rely = 0.51)
        self.c5.place(relx= 0.01, rely = 0.54)
        self.c6.place(relx= 0.01, rely = 0.57)
        self.c7.place(relx= 0.01, rely = 0.60)
        self.c8.place(relx= 0.01, rely = 0.63)
        self.c9.place(relx= 0.01, rely = 0.66)

        #-------------------------------------
        # Contact date RadioButtons placement
        #-------------------------------------
        self.R1.place(relx= 0.6, rely = 0.42)
        self.R2.place(relx= 0.6, rely = 0.45)
        self.month_var.place(relx = 0.6, rely = 0.48, relwidth= 0.1)
        self.click_ok.place(relx = 0.75, rely = 0.48, relwidth = 0.1)

        self.contacts.place(relx = 0.6, rely= 0.52)

        self.num_contacts_NB.place(relx = 0.7, rely = 0.57, relwidth = 0.2)
        self.num_contacts_WP.place(relx = 0.7, rely = 0.62, relwidth = 0.2)

        self.label_contacts_NB.place(relx = 0.6, rely = 0.57, relwidth = 0.1)
        self.label_contacts_WP.place(relx = 0.6, rely = 0.62, relwidth = 0.1)

        #-------------------------------
        # EXPORT CHECKBUTTONS placement
        #-------------------------------

        self.Ex_NB.place(relwidth = 0.8, relx= 0.5, rely = 0.72, anchor = 'n')
        self.NB.place(relwidth = 0.8, relx= 0.5, rely = 0.76, anchor = 'n')
        self.Ex_WP.place(relwidth = 0.8, relx= 0.5, rely = 0.80, anchor = 'n')

    def initialize(self):
    
        #------------------------------------------------
        #        Initialization step
        #------------------------------------------------

        # Performs the following steps:
        # 1. Imports the main file
        # 2. Does some modifications to the date formats to make them uniform
        # 3. Moves the file to the _old folder

        #import main file into pandas dataframe
        
        try:
            my_path = dirname + "/Main Database/*.xlsx"
            self.filename = glob.glob(my_path)[0]
            # self.filename = glob.glob("/Users/nikhiljain/Desktop/PYTHON/Jupyter/CAMPAIGNS/Main Database/*.xlsx")[0]
            main = pd.read_excel(self.filename, sheet_name = 'dataframe')

            main.replace(np.nan, pd.NaT)

            main['Last Campaign date'] = main['Last Campaign date'].dt.strftime('%d-%m-%Y')
            main['Neverbounce last update'] = main['Neverbounce last update'].dt.strftime('%d-%m-%Y')
            main['Date added'] = main['Date added'].dt.strftime('%d-%m-%Y')


            main['Last Campaign date'] = main['Last Campaign date'].astype(str)
            main['Neverbounce last update'] = main['Neverbounce last update'].astype(str)
            main['Date added'] = main['Date added'].astype(str)

            
            self.main_df = main

            self.move_to_old(self.filename)
            messagebox.showinfo( "Initialization done", "Initialization done")

        except:

            messagebox.showinfo( "ERROR", "Please ensure the main database has been saved in the appropriate folder")

    def update_skrapp(self, maindf):

        #------------------------------------------------
        #        Update main with new skrapped contacts
        #------------------------------------------------

        # Performs the following steps:
        # 1. Import the Skrapp CSV
        # 2. Delete irrelevant columns if any and rename columns as per code
        # 3. Add new skrapp contacts to the main file
        # 4. Delete duplicate emails (in case we skrapped anyone who was already in the database)
        # 5. Move skrapp CSV to old (uses an alternate function)
        # 6. Update the global main database
        
        try:

            my_path = dirname + "/Skrapp Files/*.csv"
            self.filename = glob.glob(my_path)[0]
            skrapp = pd.read_csv(self.filename)
            
            #Get country and region
            r_mapper = pd.read_csv('/Users/nikhiljain/Desktop/PYTHON/Jupyter/CAMPAIGNS/codefiles/region_mapper.csv', index_col = 0, squeeze = True).to_dict()
            c_mapper = pd.read_csv('/Users/nikhiljain/Desktop/PYTHON/Jupyter/CAMPAIGNS/codefiles/skrapp_to_country.csv', index_col=0, squeeze=True).to_dict()
            
            skrapp['Country'] = skrapp['Location'].map(c_mapper)
            skrapp =skrapp[['Email','First name', 'Last name', 'Title','Country','Company']]
            
            date_today = datetime.datetime.today()
            skrapp['Date added'] = date_today.strftime('%d-%m-%Y')

            skrapp['Neverbounced?'] = "NO"
            skrapp['Is in HubSpot'] = "FALSE"
            skrapp['Hubspot Blacklisted domains'] = "FALSE"
            
            
            final_df = maindf.append(skrapp)
            final_df = final_df[['Email', 'First name', 'Last name', 'Title', 'Country', 'Company','Neverbounce last update', 'Neverbounce Validity', 'Last Campaign date','Last Campaign status', 'Is in HubSpot', 'Hubspot Blacklisted domains', 'Neverbounced?', 'Date added', 'Region']]
            final_df['Region'] = final_df['Country'].map(r_mapper)
            final_df['Region']=np.where(len(final_df['Region'])<2,
                                    "Rest of the world",final_df['Region'])
            final_df = final_df.drop_duplicates(subset='Email', keep='first')
            
            self.move_to_old(self.filename)
            
            self.main_df = final_df   

            messagebox.showinfo( "Updated Skrapp", "Updated Skrapp") 

        except:

            messagebox.showinfo( "ERROR", "Please ensure the skrapp file has been saved in the folder Skrapp Files")

    def update_woodpecker(self, maindf):

        #------------------------------------------------
        #        Update main with woodpecker file
        #------------------------------------------------

        # Performs the following steps:
        # 1. Import the WP CSV delete irrelevant columns if any; rename columns as per code
        # 2. Delete irrelevant columns if any and rename columns as per code
        # 3. Merge main CSV and WP CSV
        # 4. Update information of main columns with WP columns
        # 5. Delete newly created neverbounce columns from MAIN
        # 6. Move WP CSV to old (uses an alternate function)
        # 7. Update the global main database
    
        #Import woodpecker file and set it up

        try:


            my_path = dirname + "/Woodpecker Files/*.csv"
            self.filename = glob.glob(my_path)[0]
            
            woodpecker = pd.read_csv(self.filename)
            woodpecker = woodpecker[['email','status','last contacted']]
            woodpecker.rename(columns={'status':'status_wp'},inplace=True)
            woodpecker.rename(columns={'last contacted':'last_contacted_wp'},inplace=True)
            woodpecker.rename(columns={'email':'Email'},inplace=True)
            
            woodpecker['date_wp'] = pd.to_datetime(woodpecker['last_contacted_wp'], format='%Y.%m.%d %H:%M:%S')
            woodpecker['date_wp'] = woodpecker['date_wp'].dt.strftime('%d-%m-%Y')
            
            
            #Import woodpecker column to the main database
            final_df = pd.merge(maindf,woodpecker,on=['Email'],how='left')

            final_df['Last Campaign status']=np.where(final_df['status_wp'].isnull(),
                                    final_df['Last Campaign status'],final_df['status_wp'])
            
            final_df['Last Campaign date']=np.where(final_df['date_wp'].isnull(),
                                    final_df['Last Campaign date'],final_df['date_wp'])
            
            del final_df['status_wp']
            del final_df['date_wp']
            del final_df['last_contacted_wp']

            #Move the woodpecker file to old
            self.move_to_old(self.filename)
            
            self.main_df = final_df

            messagebox.showinfo( "Updated Woodpecker", "Updated WoodPecker")

        except:

            messagebox.showinfo( "ERROR", "Please ensure the woodpecker has been saved in the folder Woodpecker Files")

    def update_hubspot(self, maindf):
        
        #------------------------------------------------
        #        Update main with hubspot file (HS)
        #------------------------------------------------

        # Performs the following steps:
        # 1. Import the HS CSV delete irrelevant columns if any; rename columns as per code
        # 2. Delete irrelevant columns if any and rename columns as per code
        # 3. Merge main CSV and HS CSV
        # 4. Update information of main columns with HS columns
        # 5. Delete newly created HS columns from MAIN
        # 6. Move HS CSV to old (uses an alternate function)
        # 7. Update the global main database

        try:

            #Import hubspot file and set it up
            my_path = dirname + "/Hubspot Files/*.csv"
            self.filename = glob.glob(my_path)[0]
            
            hubspot = pd.read_csv(self.filename)
            # hubspot = hubspot[['email','status']]
            
            #filter useful columns
            hubspot = hubspot[['Email','Email Domain', 'Relationship', 'Interest', 'Newsletter and Caution']]
            hubspot['blacklist_hs'] = (hubspot['Newsletter and Caution'].str.find('caution')>=0) | (hubspot['Interest']=='Real interest') | (hubspot['Relationship']=='True love')
            
            # code for removing rows with blacklisted domains (yahoo, gmail etc.)
            doms = hubspot['Email Domain']
            doms.drop_duplicates(keep='first', inplace=False)
            for i,x in doms.iteritems():
                if 'yahoo' in x or 'gmail' in x or 'hotmail' in x or 'orange' in x or 'outlook' in x:
                    doms = doms.drop(labels = i)

            doms = doms.str.split('.').str[0]
            domlist = doms.tolist()
            
            hubspot = hubspot[['Email','Email Domain', 'blacklist_hs']]    
            final_df = pd.merge(maindf,hubspot,on=['Email'],how='left')
            
            final_df['Is in HubSpot']=np.where(final_df['Email Domain'].isnull(),
                                    final_df['Is in HubSpot'],'TRUE')
            
            final_df['Hubspot Blacklisted domains']=np.where(final_df['blacklist_hs'].isnull(),
                                    final_df['Hubspot Blacklisted domains'],final_df['blacklist_hs'])
            
            
            del final_df['blacklist_hs']
            del final_df['Email Domain']

            #Move the woodpecker file to old
            self.move_to_old(self.filename)
            
            self.main_df = final_df
            self.domlist = domlist

            messagebox.showinfo( "Updated Hubspot", "Updated HubSpot")

        except:

            messagebox.showinfo( "ERROR", "Please ensure the hubspot file has been saved in the folder Hubspot Files")

    def update_neverbounce(self, maindf):

        #------------------------------------------------
        #        Update main with neverbounce file
        #------------------------------------------------

        # Performs the following steps:
        # 1. Import the neverbounce CSV delete irrelevant columns if any; rename columns as per code
        # 2. Delete irrelevant columns if any and rename columns as per code
        # 3. Merge main CSV and neverbounce CSV
        # 4. Update information of main columns with neverbounce columns
        # 5. Delete newly created neverbounce columns from MAIN
        # 6. Move neverbounce CSV to old (uses an alternate function)
        # 7. Update the global main database
        # 8. Export new main file
        
        # try:
        #Import woodpecker file and set it up
        my_path = dirname + "/NeverBounce Files/*.csv"
        self.filename = glob.glob(my_path)[0]

        neverbounce = pd.read_csv(self.filename)
        neverbounce = neverbounce[['Email','email_status']]
        neverbounce.columns = ['Email', 'status_nb']
        
        date_today = datetime.datetime.today()
        neverbounce['date_nb'] = date_today.strftime('%d-%m-%Y')
        
        #Import woodpecker column to the main database
        final_df = pd.merge(maindf,neverbounce,on=['Email'],how='left')
        
        final_df['Neverbounce Validity']=np.where(final_df['status_nb'].isnull(),
                                final_df['Neverbounce Validity'],final_df['status_nb'])
        
        final_df['Neverbounce last update']=np.where(final_df['date_nb'].isnull(),
                                final_df['Neverbounce last update'],final_df['date_nb'])
        final_df['Neverbounced?']=np.where(final_df['date_nb'].isnull(),
                                final_df['Neverbounced?'],'YES')
        
        del final_df['status_nb']
        del final_df['date_nb']
        
        #Move the woodpecker file to old
        self.move_to_old(self.filename)
        
        self.main_df = final_df

        messagebox.showinfo( "Updated Neverbounce", "Update Neverbounce")

        date_today = date_today.strftime('%d_%m_%Y')

        destination = dirname + "/Main Database/all_updated_main_"+date_today+".xlsx"

        final_df.replace("NaT", "")

        final_df.to_excel(destination,sheet_name = "dataframe", index=False)

        messagebox.showinfo( "New database exported", "New database exported")

        # except:

            # messagebox.showinfo( "ERROR", "Please ensure the NeverBounce file has been saved in the folder NeverBounce files")

    def export_neverbounce(self, maindf, r_selector, months_since_contact, domlist):
        
        #------------------------------------------------
        #        Export file for neverbounce
        #------------------------------------------------

        # Applies the following filters one by one:
        # 1. Regions selected
        # 2. Date filter (never contacted or last contact in x months)
        # 3. Woodpecker filters
        # 4. Neverbounce filters
        # 5. Hubspot blacklist filters
        # 6. Keep only the columns required
        # 7. Export CSV

        #region filter
        my_path = dirname + "/NeverBounce EXPORT Files/nb_export.csv"

        regions = ['Spanish South America',
                'Non Spanish South america',
                'North America',
                'France & Belgium',
                'DACH',
                'Spain',
                'Italy',
                'Other Western Europe',
                'Rest of the world']
        
        regions_code = []
        
        for x in range(0,9):
            if r_selector[x]==1:
                regions_code.append(regions[x])
        
        
        date_today = datetime.datetime.today()

        nb_export = maindf[maindf['Region'].isin(regions_code)]

        print(len(nb_export))
        
        months_since_contact = int(months_since_contact)

        if months_since_contact>0:

            #Woodpecker Date criteria (Today - last campaign date > 30*no. of months)
            nb_export['Last Campaign date'] = nb_export['Last Campaign date'].apply(pd.Timestamp)
            nb_export['since_last_campaign']=(date_today - nb_export['Last Campaign date']).dt.days
            nb_export = nb_export[nb_export['since_last_campaign'] > 30*months_since_contact]

        else:

            nb_export = nb_export[nb_export['Last Campaign date'] == 'NaT']
        
        print(len(nb_export))

        #woodpecker filters
        nb_export = nb_export[nb_export['Last Campaign status'] != 'BLACKLIST']
        nb_export = nb_export[nb_export['Last Campaign status'] != 'BOUNCED']
        nb_export = nb_export[nb_export['Last Campaign status'] != 'INVALID']
        
        print(len(nb_export))
        #neverbounce filters
        nb_export = nb_export[nb_export['Neverbounce Validity'] != 'invalid']
        nb_export = nb_export[nb_export['Neverbounce Validity'] != 'Invalid']
        
        print(len(nb_export))
        #keep only required columns
        nb_export = nb_export[['Email']]

        #hubspot filters – removes emails where domain is present in domlist
        nb_export['domain'] = (nb_export['Email'].str.split('@').str[1]).str.split('.').str[0]

        nb_export = nb_export[~nb_export['domain'].isin(domlist)]
        nb_export = nb_export[['Email']]
        
        print(len(nb_export))
        nb_export.to_csv(my_path, index = False)

        messagebox.showinfo( "NeverBounce file exported", "NeverBounce file exported")

    def export_woodpecker(self, maindf, r_selector, months_since_contact, domlist):
    
        #------------------------------------------------
        #        Export file for woodpecker
        #------------------------------------------------

        # Applies the following filters one by one:
        # 1. Regions selected
        # 2. Date filter (never contacted or last contact in x months)
        # 3. Woodpecker filters
        # 4. Neverbounce filters
        # 5. Hubspot blacklist filters
        # 6. Keep only the columns required
        # 7. Export CSV

        #region filter
        my_path = dirname + "/WoodPecker EXPORT Files/"
        
        regions = ['Spanish South America',
                'Non Spanish South america',
                'North America',
                'France & Belgium',
                'DACH',
                'Spain',
                'Italy',
                'Other Western Europe',
                'Rest of the world']
        
        regions_code = []

        wp_export = maindf
        #woodpecker filters
        wp_export = wp_export[wp_export['Last Campaign status'] != 'BLACKLIST']
        wp_export = wp_export[wp_export['Last Campaign status'] != 'BOUNCED']
        wp_export = wp_export[wp_export['Last Campaign status'] != 'INVALID']
        
        #neverbounce filters
        wp_export = wp_export[wp_export['Neverbounce Validity'] != 'invalid']
        wp_export = wp_export[wp_export['Neverbounce Validity'] != 'Invalid']
        wp_export = wp_export[wp_export['Neverbounce Validity'].notnull()]
        
        #keep only required columns
        
        #hubspot filters – removes emails where domain is present in domlist
        wp_export['domain'] = (wp_export['Email'].str.split('@').str[1]).str.split('.').str[0]

        wp_export = wp_export[~wp_export['domain'].isin(domlist)]
        
        
        date_today = datetime.datetime.today()
        months_since_contact = int(months_since_contact)

        if months_since_contact>0:

            wp_export = wp_export[wp_export['Last Campaign date'] != 'NaT']
            wp_export = wp_export[wp_export['Last Campaign date'] != '']
            wp_export = wp_export[wp_export['Last Campaign date'].notnull()]
            #Woodpecker Date criteria (Today - last campaign date > 30*no. of months)
            wp_export['Last Campaign date'] = wp_export['Last Campaign date'].apply(pd.Timestamp)
            wp_export['since_last_campaign']=(date_today - wp_export['Last Campaign date']).dt.days
            wp_export = wp_export[wp_export['since_last_campaign']>30*months_since_contact]

        else:
            wp_export_temp = wp_export[wp_export['Last Campaign date']== 'NaT']
            wp_export_temp.append(wp_export[wp_export['Last Campaign date']== ''])
            wp_export_temp.append(wp_export[wp_export['Last Campaign date'].isnull()])
            wp_export = wp_export_temp


        # CODE TO ADD SALUTATIONS:

        dict_fr = {"M":"Monsieur", "F":"Madame", "MF":"Bonjour"}
        dict_it = {"M":"Gentile Signore", "F":"Gentile Signora", "MF":"Buongiorno"}
        dict_es = {"M":"Estimado Senor", "F":"Estimada Senora", "MF":"Buenos Dias"}

        wp_export["Gender"] = wp_export["First name"].map(names_dict)
        wp_export['Gender']=wp_export['Gender'].replace(np.nan,"MF")

        wp_export["Salutation_fr"] = wp_export["Gender"].map(dict_fr)
        wp_export["Salutation_fr"] = np.where(wp_export["Salutation_fr"] == "Bonjour", "Bonjour " + wp_export['First name'], wp_export["Salutation_fr"])

        wp_export['Salutation_it'] = wp_export["Gender"].map(dict_it)

        wp_export['Salutation_es'] = wp_export["Gender"].map(dict_es)

        wp_export = wp_export[['Email', 'First name', 'Last name', 'Title', 'Country', 'Company','Region','Neverbounce Validity', 'Salutation_fr','Salutation_es','Salutation_it']]


        for x in range(0,9):
            if r_selector[x]==1:
                regions_code.append(regions[x])
        
        date_today = date_today.strftime('%d_%m_%Y')
        for region in regions_code:
            wp_export_temp = wp_export[wp_export['Region'] == region]
            dest = my_path + 'wp_export_'+region+'_'+date_today+'.csv'
            wp_export_temp.to_csv(dest, index = False)
        
        wp_export = wp_export[wp_export['Region'].isin(regions_code)]

        destination = my_path + 'wp_export_all_'+date_today+'.csv'
        wp_export.to_csv(destination, index = False)

        messagebox.showinfo( "WoodPecker file exported", "WoodPecker file exported")

    def move_to_old(self, filename):

        #------------------------------------------------
        #        Move the file to _old folder
        #------------------------------------------------

        # This code takes a file and moves it to the old folder in the same file

        destination = self.filename[0:self.filename.rfind('/')+1] + '_old/' + self.filename[self.filename.rfind('/')+1:len(self.filename)] 
        shutil.move(self.filename, destination)

    def find_months(self, rad_resp, month_var):

        # This code is used to differentiate between 'Never contacted' & 'Contacted x months ago'

        if rad_resp == 2:
            return month_var
        else:
            return -1

    def no_of_contacts(self, maindf, r_selector, months_since_contact, domlist):

        #------------------------------------------------
        #        Shows no. of contacts ready for export
        #------------------------------------------------

        # Applies the following filters one by one:
        # 1. Regions selected
        # 2. Date filter (never contacted or last contact in x months)
        # 3. Woodpecker filters
        # 4. Neverbounce filters
        # 5. Hubspot blacklist filters
        # 6. Finally displays no. of contacts to be exported on screen

        #region filter
        
        regions = ['Spanish South America',
                'Non Spanish South America',
                'North America',
                'France & Belgium',
                'DACH',
                'Spain',
                'Italy',
                'Other Western Europe',
                'Rest of the world']
        
        regions_code = []
        
        for x in range(0,9):
            if r_selector[x]==1:
                regions_code.append(regions[x])
        
        date_today = datetime.datetime.today()
        
        wp_export = maindf[maindf['Region'].isin(regions_code)]

                #woodpecker filters
        wp_export = wp_export[wp_export['Last Campaign status'] != 'BLACKLIST']
        wp_export = wp_export[wp_export['Last Campaign status'] != 'BOUNCED']
        wp_export = wp_export[wp_export['Last Campaign status'] != 'INVALID']
        
        #neverbounce filters
        wp_export = wp_export[wp_export['Neverbounce Validity'] != 'invalid']
        wp_export = wp_export[wp_export['Neverbounce Validity'] != 'Invalid']
        

        #hubspot filters – removes emails where domain is present in domlist
        wp_export['domain'] = (wp_export['Email'].str.split('@').str[1]).str.split('.').str[0]

        wp_export = wp_export[~wp_export['domain'].isin(domlist)]

        months_since_contact = int(months_since_contact)

        if months_since_contact>0:

            wp_export = wp_export[wp_export['Last Campaign date'] != 'NaT']
            #Woodpecker Date criteria (Today - last campaign date > 30*no. of months)
            wp_export['Last Campaign date'] = wp_export['Last Campaign date'].apply(pd.Timestamp)
            wp_export['since_last_campaign']=(date_today - wp_export['Last Campaign date']).dt.days
            wp_export = wp_export[wp_export['since_last_campaign']>30*months_since_contact]

        else:

            wp_export = wp_export[wp_export['Last Campaign date']== 'NaT']
        

        nb_export = wp_export

        wp_export = wp_export[wp_export['Neverbounce Validity'].notnull()]
        
        self.num_contacts_NB['text'] = len(nb_export)
        self.num_contacts_WP['text'] = len(wp_export)

class Analysis:

    dirname = os.path.dirname(__file__)

    def __init__(self, master):

        #----------------------------------
        #      GUI
        #----------------------------------

        # This function is used to describe the entire graphical user interface. There will most likely never be a problem in this part of the code.

        self.master = master
        HEIGHT = 1000
        WIDTH = 1200

        self.canvas = tk.Canvas(self.master, height = HEIGHT, width = WIDTH)
        self.canvas.pack()
        self.FRAME = tk.Frame(self.master)
        self.FRAME.place(relwidth=1,relheight=0.45)

        my_path = dirname + "/logo_keravalon.png"

        img = tk.PhotoImage(file=my_path)
        self.LABEL = tk.Label(self.FRAME, image = img)
        self.LABEL.image = img
        self.LABEL.place(relx= 0.5, rely = 0, relheight = 0.2, anchor = 'n')
        self.News = tk.Button(self.FRAME, text ="Generate Pivot", command = lambda: self.create_pivot())
        self.News.place(relwidth = 0.1, relheight = 0.05, relx= 0.5, rely = 0.23, anchor = 'n')


        #--------------------------------------------------------------
        #          Analysis summary code
        #--------------------------------------------------------------

        self.label_acc = tk.Label(self.FRAME, text = "Enter no. of accounts", justify = "left")
        self.entry_acc = tk.Entry(self.FRAME)

        self.label_mail = tk.Label(self.FRAME, text = "Enter Mails/account", justify = "left")
        self.entry_mail = tk.Entry(self.FRAME)

        self.label_newskrapp = tk.Label(self.FRAME, text = "New skrapps until Dec '19", justify = "left")
        self.entry_newskrapp = tk.Entry(self.FRAME)

        self.label_propvalid = tk.Label(self.FRAME, text = "Proportion valid (hint: 0.7)", justify = "left")
        self.entry_propvalid = tk.Entry(self.FRAME)

        self.button_acc = tk.Button(self.FRAME, text = "OK", command = lambda: self.build_analysis(self.entry_acc.get(), self.entry_newskrapp.get(), self.entry_mail.get(), self.entry_propvalid.get() ,self.main_pivot))

        self.label_acc.place(relx = 0.01, rely=0.3, relwidth = 0.16)
        self.entry_acc.place(relx = 0.18, rely=0.3, relwidth = 0.08)

        self.label_mail.place(relx = 0.01, rely=0.38, relwidth = 0.16)
        self.entry_mail.place(relx = 0.18, rely=0.38, relwidth = 0.08)

        self.label_newskrapp.place(relx = 0.01, rely=0.46, relwidth = 0.16)
        self.entry_newskrapp.place(relx = 0.18, rely=0.46, relwidth = 0.08)

        self.label_propvalid.place(relx = 0.01, rely=0.54, relwidth = 0.16)
        self.entry_propvalid.place(relx = 0.18, rely=0.54, relwidth = 0.08)

        self.button_acc.place(relx = 0.18, rely = 0.61, relwidth = 0.08)

        self.label_assumption = tk.Label(self.FRAME, text = "Assuming all valid*", justify = "left")
        self.label_never = tk.Label(self.FRAME, text = "–> Never contacted", justify = "left")
        self.label_not19 = tk.Label(self.FRAME, text = "–> Not contacted '19", justify = "left")
        self.label_wo_skrapp = tk.Label(self.FRAME, text = "W/o new skraps:", justify = "left")
        self.label_w_skrapp = tk.Label(self.FRAME, text = "With new skraps:", justify = "left")

        self.label_assumption.place(relx =0.01, rely = 0.7, relwidth = 0.15)
        self.label_wo_skrapp.place(relx = 0.2, rely = 0.7, relwidth = 0.1)
        self.label_w_skrapp.place(relx = 0.3, rely = 0.7, relwidth = 0.1)
        self.label_never.place(relx =0.01, rely = 0.8, relwidth = 0.15)
        self.label_not19.place(relx =0.01, rely = 0.85, relwidth = 0.15)

        self.label_assumption2 = tk.Label(self.FRAME, text = "Assuming 70% valid*", justify = "left")
        self.label_wo_skrapp2 = tk.Label(self.FRAME, text = "W/o new skraps:", justify = "left")
        self.label_w_skrapp2 = tk.Label(self.FRAME, text = "With new skraps:", justify = "left")
        self.label_never2 = tk.Label(self.FRAME, text = "–> Never contacted", justify = "left")
        self.label_not19_2 = tk.Label(self.FRAME, text = "–> Not contacted '19", justify = "left")

        self.new_skrapp = tk.Entry(self.FRAME)

        self.label_assumption2.place(relx =0.51, rely = 0.7, relwidth = 0.15)
        self.label_wo_skrapp2.place(relx = 0.7, rely = 0.7, relwidth = 0.1)
        self.label_w_skrapp2.place(relx = 0.8, rely = 0.7, relwidth = 0.1)
        self.label_never2.place(relx =0.51, rely = 0.8, relwidth = 0.15)
        self.label_not19_2.place(relx =0.51, rely = 0.85, relwidth = 0.15)

        self.date11 = tk.Label(self.FRAME, text = "", justify = "left")
        self.date12 = tk.Label(self.FRAME, text = "", justify = "left")
        self.date13 = tk.Label(self.FRAME, text = "", justify = "left")
        self.date14 = tk.Label(self.FRAME, text = "", justify = "left")

        self.date11.place(relx = 0.2, rely= 0.8, relwidth = 0.1)
        self.date12.place(relx = 0.3, rely= 0.8, relwidth = 0.1)
        self.date13.place(relx = 0.7, rely= 0.8, relwidth = 0.1)
        self.date14.place(relx = 0.8, rely= 0.8, relwidth = 0.1)

        self.date21 = tk.Label(self.FRAME, text = "", justify = "left")
        self.date22 = tk.Label(self.FRAME, text = "", justify = "left")
        self.date23 = tk.Label(self.FRAME, text = "", justify = "left")
        self.date24 = tk.Label(self.FRAME, text = "", justify = "left")

        self.date21.place(relx = 0.2, rely= 0.85, relwidth = 0.1)
        self.date22.place(relx = 0.3, rely= 0.85, relwidth = 0.1)
        self.date23.place(relx = 0.7, rely= 0.85, relwidth = 0.1)
        self.date24.place(relx = 0.8, rely= 0.85, relwidth = 0.1)

        self.FRAME2 = tk.Frame(self.master)
        self.FRAME2.place(relwidth = 1, relheight = 0.5, relx = 0, rely= 0.5)

    def create_pivot(self):

        try:

            my_path = dirname + "/Main Database/*.xlsx"
            self.filename = glob.glob(my_path)[0]
            main = pd.read_excel(self.filename, sheet_name = "dataframe")

            main['Last Campaign date'] = main['Last Campaign date'].apply(pd.Timestamp)
            main['Neverbounce last update'] = main['Neverbounce last update'].apply(pd.Timestamp)
            main['Date added'] = main['Date added'].apply(pd.Timestamp)

            main['Date Added Year'] = main["Date added"].dt.year
            main['Last Campaign Year'] = main["Last Campaign date"].dt.year

            main['Last Campaign Year']=np.where(main['Last Campaign Year'].notnull(),
                                    main['Last Campaign Year'],'Never')

            main_pivot = main.pivot_table(index = ["Date Added Year","Last Campaign Year","Neverbounced?"], 
                columns = "Region", values = "Email", aggfunc = "count", margins = True)

            pivot_analysis = main_pivot

            pivot_analysis = pivot_analysis.reset_index()
            pivot_analysis['Date Added Year'] = pivot_analysis['Date Added Year'].drop_duplicates(keep = 'first')
            pivot_analysis = pivot_analysis.replace(np.nan, '0')
            pivot_analysis['Date Added Year'] = np.where(pivot_analysis['Date Added Year'] == '0', '', pivot_analysis['Date Added Year'])

            for i in range(1,len(pivot_analysis)-1):
                if pivot_analysis['Last Campaign Year'][i-1] == pivot_analysis['Last Campaign Year'][i]:
                    pivot_analysis['Last Campaign Year'][i] = ''

            df = pivot_analysis
            self.main_pivot = df
            tree = ttk.Treeview(self.FRAME2)


            df_col = df.columns.values
            tree["columns"]=(df_col)
            counter = len(df)
            tree.column("#0", width=0)

            for x in range(len(df_col)):
                tree.column(x, width= 80 )
                tree.heading(x, text=df_col[x])
                
            for i in range(counter):
                mylist = []
                for x in range(len(df_col)):
                    mylist.append(df[df_col[x]][i])
                tree.insert('', "end", values=mylist)

            tree.place(relx = 0, rely = 0, relwidth = 1, relheight =1)

        except:

            messagebox.showinfo( "ERROR", "Please ensure the main database has been saved in the appropriate folder")

    def build_analysis(self,no_acc,newskrapp,mailperacc,propvalid,main_pivot):

        try:
            pivot_analysis = self.main_pivot 
            number_acc = int(no_acc)
            new_skrapp = int(newskrapp)
            mailperacc = int(mailperacc)
            propvalid = float(propvalid)

        except:
            messagebox.showinfo( "ERROR", "Please enter appropriate values")

        for i in range(1,len(pivot_analysis)-1):
            if (pivot_analysis['Date Added Year'][i-1]!="" and pivot_analysis['Date Added Year'][i]==""):
                pivot_analysis['Date Added Year'][i] = pivot_analysis['Date Added Year'][i-1]

        for i in range(1,len(pivot_analysis)-1):
            if (pivot_analysis['Last Campaign Year'][i-1]!="" and pivot_analysis['Last Campaign Year'][i]==""):
                pivot_analysis['Last Campaign Year'][i] = pivot_analysis['Last Campaign Year'][i-1]

        df = pivot_analysis
        # new_skrapp = int(self.new_skrapp.get())

        date_today = datetime.datetime.now()

        add_days = int(((df.loc[df['Last Campaign Year'] == "Never", 'All'].sum()))/(mailperacc*number_acc))
        date_11 = date_today + datetime.timedelta(days = add_days)
        add_days = int(add_days*propvalid)
        date_13 = date_today + datetime.timedelta(days = add_days)
        add_days = int(((df.loc[df['Last Campaign Year'] == "Never", 'All'].sum()) + new_skrapp)/(mailperacc*number_acc))
        date_12 = date_today + datetime.timedelta(days = add_days)
        add_days = int(add_days*propvalid)
        date_14 = date_today + datetime.timedelta(days = add_days)

        self.date11['text'] = date_11.strftime('%d/%m/%Y')
        self.date12['text'] = date_12.strftime('%d/%m/%Y')
        self.date13['text'] = date_13.strftime('%d/%m/%Y')
        self.date14['text'] = date_14.strftime('%d/%m/%Y')

        no_not_19 = df.loc[df['Last Campaign Year']!= "2019.0", 'All'].sum() - df['All'][len(df)-1]

        add_days = int(no_not_19/(mailperacc*number_acc))
        date_21 = date_today + datetime.timedelta(days = add_days)
        add_days = int(add_days*propvalid)
        date_23 = date_today + datetime.timedelta(days = add_days)
        add_days = int((no_not_19 + new_skrapp)/(mailperacc*number_acc))
        date_22 = date_today + datetime.timedelta(days = add_days)
        add_days = int(add_days*propvalid)
        date_24 = date_today + datetime.timedelta(days = add_days)

        self.date21['text'] = date_21.strftime('%d/%m/%Y')
        self.date22['text'] = date_22.strftime('%d/%m/%Y')
        self.date23['text'] = date_23.strftime('%d/%m/%Y')
        self.date24['text'] = date_24.strftime('%d/%m/%Y')



        # print(date_11.strftime('%d_%m_%Y'))
        # print(date_12.strftime('%d_%m_%Y'))

class Newsletter:

    dirname = os.path.dirname(__file__)

    #----------------------------------------------
    #   Export a list of contacts for newsletter
    #----------------------------------------------

    def __init__(self, master):

        #----------------------------------
        #      GUI
        #----------------------------------

        # This function is used to describe the entire graphical user interface. There will most likely never be a problem in this part of the code.

        self.master = master
        HEIGHT = 750
        WIDTH = 500

        self.canvas = tk.Canvas(self.master, height = HEIGHT, width = WIDTH)
        self.canvas.pack()
        self.FRAME = tk.Frame(self.master)
        self.FRAME.place(relwidth=1,relheight=1)

        my_path = dirname + "/logo_keravalon.png"

        img = tk.PhotoImage(file=my_path)
        self.LABEL = tk.Label(self.FRAME, image = img)
        self.LABEL.image = img
        self.LABEL.place(relx= 0.5, rely = 0, relheight = 0.1, anchor = 'n')    
        
        self.News = tk.Button(self.FRAME, text ="Export newsletter contacts", command = lambda: self.news_export())
        self.News.place(relwidth = 0.8, relx= 0.5, rely = 0.4, relheight = 0.2, anchor = 'n')

    def news_export(self):

        try:

            my_path = dirname + "/Hubspot Files/*.csv"
            self.filename = glob.glob(my_path)[0]
        
            df = pd.read_csv(self.filename)
            df = df[['Email','First Name', 'Last Name', 'Newsletter and Caution']]
            ndf = df[df['Newsletter and Caution']=="Newsletter"]
            ndf = ndf.append(df[df['Newsletter and Caution']=="Newsletter; Treat with caution"])
            ndf = ndf.append(df[df['Newsletter and Caution']=="Treat with caution; Newsletter"])

            date_today = datetime.datetime.today()
            date_today = date_today.strftime('%d_%m_%Y')

            destination = dirname + '/File for newsletter/Newletter_' + date_today + '.csv'

            ndf.to_csv(destination, index = False)

            messagebox.showinfo( "Newsletter file exported", "Newsletter file exported")

        except:
            messagebox.showinfo( "ERROR", "Please ensure hubspot CSV has been saved in the folder Hubspot Files")

class ManualContact:

    dirname = os.path.dirname(__file__)

    #----------------------------------------------
    #   Export a list of contacts for newsletter
    #----------------------------------------------

    def __init__(self, master):

        #----------------------------------
        #      GUI
        #----------------------------------

        # This function is used to describe the entire graphical user interface. There will most likely never be a problem in this part of the code.

        self.master = master
        HEIGHT = 750
        WIDTH = 500

        self.canvas = tk.Canvas(self.master, height = HEIGHT, width = WIDTH)
        self.canvas.pack()
        self.FRAME = tk.Frame(self.master)
        self.FRAME.place(relwidth=1,relheight=1)

        my_path = dirname + "/logo_keravalon.png"

        img = tk.PhotoImage(file=my_path)
        self.LABEL = tk.Label(self.FRAME, image = img)
        self.LABEL.image = img
        self.LABEL.place(relx= 0.5, rely = 0, relheight = 0.1, anchor = 'n')    
        
        self.Cont = tk.Button(self.FRAME, text ="Export contacts", command = lambda: self.contact_export())
        self.Cont.place(relwidth = 0.8, relx= 0.5, rely = 0.4, relheight = 0.2, anchor = 'n')

    def contact_export(self):
        
        try:
            my_path = dirname + "/Hubspot Files/*.csv"
            self.filename = glob.glob(my_path)[0]
            df = pd.read_csv(self.filename)

            date_today = datetime.datetime.today()

            df = df[['Email','First Name', 'Last Name', 'Newsletter and Caution', 'Interest','Last Contacted','Relationship']]

            df['Last Contacted'] = pd.to_datetime(df['Last Contacted'], format='%Y.%m.%d %H:%M:%S')

            df['Last Contacted'] = df['Last Contacted'].dt.strftime('%d-%m-%Y')

            df['Last Contacted'] = df['Last Contacted'].apply(pd.Timestamp)

            df['Days since contact'] = (date_today - df['Last Contacted']).dt.days


            ndf = df[df['Interest'] == 'Little interest']
            ndf2 = ndf[ndf['Relationship'] == 'Neutral']
            ndf2 = ndf2.append(ndf[ndf['Relationship']=='Friendship'])
            ndf2 = ndf2.append(ndf[ndf['Relationship']=='True love'])


            ndf2 = ndf2[ndf2['Days since contact']>270]

            date_today = date_today.strftime('%d_%m_%Y')
            

            destination = dirname + '/File for newsletter/Manual_Contact_' + date_today + '.csv'

            ndf2.to_csv(destination, index = False)

            messagebox.showinfo( "Manual Contact file exported", "Manual Contact file exported")

        except:

            messagebox.showinfo( "ERROR", "Please ensure hubspot CSV has been saved in the folder Hubspot Files")


#-----------------------------------
#    App launcher
#-----------------------------------

# this code launches the app. There will never be a bug here

def main():

    #-----------------------------------
    #    App launcher
    #-----------------------------------

    # this code launches the app. There will never be a bug here
    root = tk.Tk()
    app = MainWindow(root)
    root.mainloop()

if __name__ == '__main__':

    #-----------------------------------
    #    App launcher
    #-----------------------------------

    # this code launches the app. There will never be a bug here

    main()
