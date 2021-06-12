import xlrd
import xlwt
from datetime import datetime
import calendar
import os
import uuid
from .settings import Settings
import shutil
import sys


class ETLSource:
    # ETLSource class creates the source XL
    # for content generation

    def __init__(self):
        self.uuid_str = str(uuid.uuid4())
        self.target_file = ""
        self.source_file = ""
        self.fmt_sheet1 = ""
        self.fmt_sheet2 = ""
        self.fmt_sheet3 = ""
        self.fmt_sheet4 = ""
        self.fmt_sheet5 = ""
        self.col_width_sheet1 = ""
        self.col_width_sheet2 = ""
        self.col_width_sheet3 = ""
        self.col_width_sheet4 = ""
        self.col_width_sheet5 = ""
        self.pt_date = ""
        self.bulletin_title_dict = {}
        self.bulletin_qnum = {}
        self.bulletin_details = {}
        self.patch_details = {}
        self.products = {}
        self.files = {}
        self.registry = {}
        self.start_row = 0

    def get_pt_date(self):
        now = datetime.now()
        month = now.month
        year = now.year
        cal = calendar.monthcalendar(year,month)
        pt_week_1 = cal[0]
        pt_week_2 = cal[1]
        pt_week_3 = cal[2]
        pt_day = ""
        
        # If a Saturday presents in the first week, the second Saturday
        # is in the second week.  Otherwise, the second Saturday must 
        # be in the third week.
            
        if pt_week_1[calendar.TUESDAY]:
            pt_day = pt_week_2[calendar.TUESDAY]
        else:
            pt_day = pt_week_3[calendar.TUESDAY]

        self.pt_date = str(pt_day) + "-" + calendar.month_abbr[month] + "-" + str(year)

    def get_column_names(self, sheet_number):
        if sheet_number == 1:
            return(self.fmt_sheet1[self.fmt_sheet1.find("|") + 1:].split(';'))
        elif sheet_number == 2:
            return(self.fmt_sheet2[self.fmt_sheet2.find("|") + 1:].split(';'))
        elif sheet_number == 3:
            return(self.fmt_sheet3[self.fmt_sheet3.find("|") + 1:].split(';'))
        elif sheet_number == 4:
            return(self.fmt_sheet4[self.fmt_sheet4.find("|") + 1:].split(';'))
        elif sheet_number == 5:
            return(self.fmt_sheet5[self.fmt_sheet5.find("|") + 1:].split(';'))

    def get_column_width(self, sheet_number):
        if sheet_number == 1:
            return(self.col_width_sheet1.split(';'))
        elif sheet_number == 2:
            return(self.col_width_sheet2.split(';'))
        elif sheet_number == 3:
            return(self.col_width_sheet3.split(';'))
        elif sheet_number == 4:
            return(self.col_width_sheet4.split(';'))
        elif sheet_number == 5:
            return(self.col_width_sheet5.split(';'))

    def get_bulletin_qnums(self):
        # read the source file
        source_wb = xlrd.open_workbook(self.source_file) 
        source_ws = source_wb.sheet_by_index(0) 
        rows = source_ws.nrows

        bulletin = ""
        kb = "" # the numeric value
        qnum = ""

        final_qnum = "" 
        already_read_bulletin = source_ws.cell_value(0, 19)
        already_read_kb = ""
        
        for i in range(rows):
            bulletin = (
                source_ws.cell_value(i, 19).rstrip(os.linesep)
            )
            kb = source_ws.cell_value(i, 8).rstrip(
                os.linesep).replace("KB","").replace("kb","")
            if (bulletin != "Bulletin"):
                if (bulletin != already_read_bulletin):
                    final_qnum = qnum
                    if (already_read_bulletin != ""):
                        self.bulletin_qnum[already_read_bulletin] = final_qnum
                    qnum = kb
                else:
                    if (already_read_kb != kb):
                        qnum = qnum + "," + kb
                already_read_kb = kb 
                already_read_bulletin = bulletin   

    def get_bulletin_title(self):
        
        # read the source file
        source_wb = xlrd.open_workbook(self.source_file) 
        ws = source_wb.sheet_by_index(0) 
        rows = ws.nrows

        bulletin = ""      
        category = ""
        prod_sp = ""
        build = ""
        prod_sp_list_for_bulletin = []
        title = ""
    
        #title logic for bulletin summary
        for i in range(rows):
            build = ""
            if (ws.cell_value(i, 21).rstrip(os.linesep) == "YES" 
                or  ws.cell_value(i, 21).rstrip(os.linesep) == "SSU"
                or  ws.cell_value(i, 21).rstrip(os.linesep) == "SEC"):
                if (i != 0):
                    if ws.cell_value(i, 0) != "":
                        category = ws.cell_value(i, 0) 
                    
                    if ws.cell_value(i, 19).rstrip(os.linesep) != bulletin:
                        bulletin = ws.cell_value(i, 19).rstrip(os.linesep) 
                        prod_sp_list_for_bulletin = []

                    if category =="Office" or (
                        category == "SQL Server") or (
                        category == ".Net Framework") or (
                        category == "Exchange"):
                        prod_sp = bulletin
                    elif "Windows Security" in category:
                        prod_sp =  ws.cell_value(i, 2).rstrip(os.linesep)
                    elif category =="Browsers":
                        prod_sp =  "Internet Explorer"
                    elif "Server" in ws.cell_value(i, 2).rstrip(os.linesep):
                        prod_sp =  ws.cell_value(i, 2).rstrip(os.linesep)
                    elif "RTM" in  str(ws.cell_value(i, 4)):
                        prod_sp =  ws.cell_value(i, 2).rstrip(os.linesep)

                    else: 
                        build = str(ws.cell_value(i, 4))
                        build = build[:-2]
                        if build != "":
                            prod_sp = ws.cell_value(i, 2).rstrip(
                                os.linesep) + " build " + build 
                        else:
                            prod_sp = ws.cell_value(i, 2).rstrip(os.linesep) 

                    if prod_sp == "Windows 10":
                        build = "RTM" 
                        prod_sp = ws.cell_value(i, 2).rstrip(
                            os.linesep) + " " + build 

                    if prod_sp not in prod_sp_list_for_bulletin:
                        prod_sp_list_for_bulletin.append(prod_sp)

                    if ws.cell_value(i+1, 19).rstrip(os.linesep) != bulletin:
                        if "Rel" in bulletin:
                            title = "Cumulative update for " + bulletin
                        else:
                            for psp in prod_sp_list_for_bulletin:
                                title = title + psp + ', ' 
                            title = title [:-2]
                        self.bulletin_title_dict[bulletin] = title
                        title = ""

    def get_bulletin_details(self):
        data_row = ""
        category = ""

        title = ""
        bulletin_location = ""
        bulletin = ""
        URL = "http://support.microsoft.com/kb/"
        i = 0
        # read the source file
        source_wb = xlrd.open_workbook(self.source_file) 
        source_ws = source_wb.sheet_by_index(0) 
        rows = source_ws.nrows

        for i in range(self.start_row, rows): 
            data_row = ""
            if (source_ws.cell_value(i, 21).rstrip(os.linesep) == "YES" 
                or  source_ws.cell_value(i, 21).rstrip(os.linesep) == "SSU"
                or  source_ws.cell_value(i, 21).rstrip(os.linesep) == "SEC"):
                if (i != 0):
                    # bulletin
                    bulletin = source_ws.cell_value(i, 19).rstrip(os.linesep)
                    data_row = data_row + bulletin
                    data_row = data_row + "|"    
                    
                    # bulletin_location    
                    bulletin_location = URL + bulletin.replace("KB", "")
                    data_row = data_row + bulletin_location
                    data_row = data_row + "|"    
                    
                    # faq_location
                    data_row = data_row + bulletin_location
                    data_row = data_row + "|" 
                    
                    # faq_page_name
                    data_row = data_row + "FQ" + bulletin
                    data_row = data_row + "|"     
                    
                    if source_ws.cell_value(i, 0) != "":
                        category = source_ws.cell_value(i, 0) 
                    
                    if category =="Office" or (
                        category == "SQL Server") or (
                        category == ".Net Framework")  or (
                        category == "Exchange"):
                        title = "Cumulative Update for " + bulletin
                    elif category =="SSU":
                        title = "Servicing Stack update for " + (
                            self.bulletin_title_dict[bulletin] 
                        )
                    elif "Windows Security" in category:
                        title = "Security only update for " + (
                            self.bulletin_title_dict[bulletin] 
                        )
                    elif category =="Browsers":
                        title =  "Cumulative Security update for Internet Explorer"
                    else:
                        #print(source_ws.cell_value(i, 19))
                        #if "Rel" in source_ws.cell_value(i, 19):
                        #    title =  "Cumulative Security update for " +source_ws.cell_value(i, 19)
                        #else:
                        title = "Cumulative Update for "  + self.bulletin_title_dict[bulletin] 

                    data_row = data_row + title     
                    data_row = data_row + "|"
                    
                    # date_posted
                    data_row = data_row + self.pt_date     
                    data_row = data_row + "|"
                    
                    # date_revised
                    data_row = data_row + self.pt_date     
                    data_row = data_row + "|"
                    
                    # supported
                    data_row = data_row + "Yes"     
                    data_row = data_row + "|"
                    
                    # summary
                    data_row = data_row + ""     
                    data_row = data_row + "|"
                    
                    # issue
                    data_row = data_row + ""     
                    data_row = data_row + "|"
                    
                    # impact_severity_id
                    data_row = data_row + "0"     
                    data_row = data_row + "|"
                    
                    # pre_req_severity_id
                    data_row = data_row + "0"     
                    data_row = data_row + "|"
                    
                    # mitigation_severity_id
                    data_row = data_row + "0"     
                    data_row = data_row + "|"
                    
                    # popularity_severity_id
                    data_row = data_row + source_ws.cell_value(i, 20)     
                    data_row = data_row + "|"
                    
                    # bulletin_comments
                    data_row = data_row + ""    
                    data_row = data_row + "|"
                    
                    # qnumbers
                    data_row = data_row + self.bulletin_qnum[bulletin]     
                    data_row = data_row + "|"
                    
                    # superceded_bulletins 
                    data_row = data_row + source_ws.cell_value(i, 9)       
                    data_row = data_row + "|"
                    
                    self.bulletin_details[i] = data_row
                    
    def get_patch_details(self):
        data_row = ""

        bulletin = ""
        patch_location = ""	
        sq_number = ""	
        i = 1

        # read the source file
        source_wb = xlrd.open_workbook(self.source_file) 
        source_ws = source_wb.sheet_by_index(0) 
        rows = source_ws.nrows

        for i in range(self.start_row, rows): 
            if (source_ws.cell_value(i, 21).rstrip(os.linesep) == "YES" or(
                source_ws.cell_value(i, 21).rstrip(os.linesep) == "SSU") or(
                source_ws.cell_value(i, 21).rstrip(os.linesep) == "SEC")):
                bulletin = source_ws.cell_value(i, 19)
                data_row = data_row + bulletin
                data_row = data_row + "|"
                if source_ws.cell_value(i, 0) != "":
                    category = source_ws.cell_value(i, 0) 

                # patch_file
                data_row = data_row + source_ws.cell_value(i, 16)
                data_row = data_row + "|"

                patch_location = source_ws.cell_value(i, 17)
                data_row = data_row + patch_location
                data_row = data_row + "|"
                
                # patch_comments
                data_row = data_row + ""
                data_row = data_row + "|"

                # absolute_path
                data_row = data_row + ""
                data_row = data_row + "|"	
                
                # patchdetails_scan_before_deploy
                data_row = data_row + "0"
                data_row = data_row + "|"

                # patchdetails_scan_after_deploy
                data_row = data_row + "0"
                data_row = data_row + "|"

                # patchdetails_allow_greater_version 
                data_row = data_row + "1"
                data_row = data_row + "|"

                # patchdetails_file_checksum 
                data_row = data_row + "NULL"
                data_row = data_row + "|"	
                
                # patchdetails_last_modified
                data_row = data_row + self.pt_date
                data_row = data_row + "|"	

                # patchfiles_file_name
                data_row = data_row + "%systemroot%\\regedit.exe"	
                data_row = data_row + "|"

                # patchfiles_file_version 
                data_row = data_row + "NULL"	
                data_row = data_row + "|"

                # patchfiles_file_size
                data_row = data_row + "NULL"	
                data_row = data_row + "|"

                # patchfiles_file_checksum 
                data_row = data_row + "NULL"	
                data_row = data_row + "|"

                # patchfiles_last_modified
                data_row = data_row + self.pt_date
                data_row = data_row + "|"	

                # patchswitch_type = "1"
                data_row = data_row + "1"
                data_row = data_row + "|"

                if category =="Office":
                    patchswitch_switches = "/quiet /passive /norestart"
                elif category ==".Net Framework":
                    patchswitch_switches = "/quiet /norestart"
                elif category == "SQL Server":
                    patchswitch_switches = "/quiet /allinstances /IAcceptSQLServerLicenseTerms"
                elif category == "Exchange":
                    patchswitch_switches = ""
                elif category == "Browsers":
                    patchswitch_switches = "/quiet /norestart"
                elif "Windows Security" in category:
                    patchswitch_switches =  "/quiet /norestart"
                else:
                    patchswitch_switches= "/quiet /norestart"
                
                # patchswitch_switches
                data_row = data_row + patchswitch_switches
                data_row = data_row + "|"
                
                if (bulletin.find('Rel') == -1):
                    sq_number = bulletin.replace("KB", "").replace("kb","")		
                else: 
                    sq_number = source_ws.cell_value(i, 8).rstrip(os.linesep).replace("KB","").replace("kb","")

                # no_reboot
                data_row = data_row + "0"	
                data_row = data_row + "|"

                # sq_number
                data_row = data_row + sq_number
                data_row = data_row + "|"


                # superceded_bulletin
                data_row = data_row + source_ws.cell_value(i, 9)	
                data_row = data_row + "|"

                self.patch_details[i] = data_row
                data_row = ""

    def get_patch_products(self):
        data_row = ""
        
        i = 0
        flag = ""

        # read the source file
        source_wb = xlrd.open_workbook(self.source_file) 
        source_ws = source_wb.sheet_by_index(0) 
        rows = source_ws.nrows

        for i in range(self.start_row, rows): 
            flag = source_ws.cell_value(i, 21).rstrip(os.linesep)

            if ( flag == "YES" or flag == "SSU" or flag == "SEC"):
                # bulletin
                data_row = data_row + source_ws.cell_value(i, 19)
                data_row = data_row + "|"

                # patch_file
                data_row = data_row + source_ws.cell_value(i, 16)
                data_row = data_row + "|"
                
                # product
                data_row = data_row + source_ws.cell_value(i, 5)
                data_row = data_row + "|"

                # service_pack
                data_row = data_row + source_ws.cell_value(i, 6)
                data_row = data_row + "|"

                # fixed_in_sp
                data_row = data_row + ""
                data_row = data_row + "|"

                # url
                data_row = data_row + ""
                data_row = data_row + "|"

                # release_date
                data_row = data_row + ""
                data_row = data_row + "|"

                self.products[i] = data_row
                data_row = ""
        
    def get_patch_files(self):
        data_row = ""
        
        i=0
        flag = ""

        # read the source file
        source_wb = xlrd.open_workbook(self.source_file) 
        source_ws = source_wb.sheet_by_index(0) 
        rows = source_ws.nrows

        for i in range(self.start_row, rows): 
            flag = source_ws.cell_value(i, 21).rstrip(os.linesep)

            if ( flag == "YES" or flag == "SSU" or flag == "SEC"):

                # bulletin
                data_row = data_row + source_ws.cell_value(i, 19)
                data_row = data_row + "|"
                
                # patch_file
                patch_file = source_ws.cell_value(i, 16)
                data_row = data_row + patch_file
                data_row = data_row + "|"

                # file_name
                file_name = source_ws.cell_value(i, 10)
                data_row = data_row + file_name
                data_row = data_row + "|"

                # version
                data_row = data_row + source_ws.cell_value(i, 11)
                data_row = data_row + "|"

                # location
                data_row = data_row + source_ws.cell_value(i, 13)
                data_row = data_row + "|"

                # file size
                data_row = data_row + ""
                data_row = data_row + "|"

                # command id
                data_row = data_row + ""
                data_row = data_row + "|"

                # checksum
                data_row = data_row + ""
                data_row = data_row + "|"

                self.files[i] = data_row
                data_row = ""

    def get_registry(self):
        data_row = ""
        
        i=0
        flag = ""

        # read the source file
        source_wb = xlrd.open_workbook(self.source_file) 
        source_ws = source_wb.sheet_by_index(0) 
        rows = source_ws.nrows

        for i in range(self.start_row, rows): 
        
            flag = source_ws.cell_value(i, 21).rstrip(os.linesep)

            if ( flag == "YES" or flag == "SSU" or flag == "SEC"):
                if len(source_ws.cell_value(i, 10).rstrip(os.linesep)) == 0:
                    
                    # bulletin
                    data_row = data_row + source_ws.cell_value(i, 19)
                    data_row = data_row + "|"

                    # patch_file
                    patch_file = source_ws.cell_value(i, 16)
                    data_row = data_row + patch_file
                    data_row = data_row + "|"

                    # registry_path
                    data_row = data_row + "Installed Hot Fixes"
                    data_row = data_row + "|"

                    # absolute_path
                    data_row = data_row + ""
                    data_row = data_row + "|"

                    # key
                    data_row = data_row + "Type"
                    data_row = data_row + "|"

                    # value 
                    data_row = data_row + source_ws.cell_value(i, 8)
                    data_row = data_row + "|"

                    # type
                    data_row = data_row + "0"
                    data_row = data_row + "|"

                    self.registry[i] = data_row
                    data_row = ""
            
    def create(self):
        print("START Creating the source XL for content generation")
        s = Settings()
        s.get_settings()

        self.source_file = s.XL_processed
        self.target_file = s.ETL_xl
        self.fmt_sheet1 = s.ETL_xl_format_sheet1
        self.fmt_sheet2 = s.ETL_xl_format_sheet2
        self.fmt_sheet3 = s.ETL_xl_format_sheet3
        self.fmt_sheet4 = s.ETL_xl_format_sheet4
        self.fmt_sheet5 = s.ETL_xl_format_sheet5
        self.col_width_sheet1 = s.ETL_xl_col_width_sheet1
        self.col_width_sheet2 = s.ETL_xl_col_width_sheet2
        self.col_width_sheet3 = s.ETL_xl_col_width_sheet3
        self.col_width_sheet4 = s.ETL_xl_col_width_sheet4
        self.col_width_sheet5 = s.ETL_xl_col_width_sheet5
        self.start_row = s.XL_start_row

        if not os.path.isfile (self.source_file):
            print(" Missing file: " + self.source_file)
            sys.exit(1)

        titles = []
        col_width = []
        tmp_name = self.target_file[: len(self.target_file)-4] + (
             "_" + str(uuid.uuid4())  +".xls")

        if os.path.isfile(self.target_file):
           shutil.move(self.target_file, 
                      s.FOLDER_tmp + "\\" + 
                      os.path.basename(tmp_name))

        font = xlwt.Font() # Create the Font
        font.name = s.ETL_xl_text_font
        font.height = int(s.ETL_xl_text_font_size)
        style = xlwt.XFStyle() # Create the Style
        style.font = font # Apply the fone

        font_title = xlwt.Font() # Create the Font
        font_title.name = s.ETL_xl_title_text_font
        font_title.height = int(s.ETL_xl_title_text_font_size)
        font_title.bold = s.ETL_xl_title_text_font_bold
        style_title = xlwt.XFStyle() # Create the Style
        style_title.font = font_title # Apply the font

        self.get_pt_date()
        
        destination_wb = xlwt.Workbook(encoding = 'ascii')
        
        # read the source file
        source_wb = xlrd.open_workbook(self.source_file) 
        source_ws = source_wb.sheet_by_index(0) 
        rows = source_ws.nrows

        #Bulletins Detail Page
        sheet_name = self.fmt_sheet1[0: self.fmt_sheet1.find("|")]
        ws = destination_wb.add_sheet(sheet_name)

        col_width = self.get_column_width(1)
        index = 0
        length = len(col_width)

        while index < length:
            ws.col(index).width = (
                int(col_width[index])) 
            index = index + 1 

        titles = self.get_column_names(1)
        index = 0
        length = len(titles)

        while index < length:
            ws.write(0, index, 
                label = titles[index] ,style = style_title)
            index = index + 1

        self.get_bulletin_qnums()
        self.get_bulletin_title()
        self.get_bulletin_details()
        
        bulletins_written = []
        dest_row = 1
        data = []    
        rows = []
        rows = self.bulletin_details.keys()

        print("        writing Bulletin details page")
        for i in rows:
            data = self.bulletin_details[i].split("|")
            data
            if data[0] not in bulletins_written:
                if (data[0].strip() != "" and data[0].strip() != "" ):
                    ws.write(dest_row, 0, label = data[0] ,style = style)
                    ws.write(dest_row, 1, label = data[1] ,style = style)
                    ws.write(dest_row, 2, label = data[2] ,style = style)
                    ws.write(dest_row, 3, label = data[3] ,style = style)
                    ws.write(dest_row, 4, label = data[4] ,style = style)
                    ws.write(dest_row, 5, label = data[5] ,style = style)
                    ws.write(dest_row, 6, label = data[6] ,style = style)
                    ws.write(dest_row, 7, label = data[7] ,style = style)
                    ws.write(dest_row, 8, label = data[8] ,style = style)
                    ws.write(dest_row, 9, label = data[9] ,style = style)
                    ws.write(dest_row, 10, label = data[10] ,style = style)
                    ws.write(dest_row, 11, label = data[11] ,style = style)
                    ws.write(dest_row, 12, label = data[12] ,style = style)
                    ws.write(dest_row, 13, label = data[13] ,style = style)
                    ws.write(dest_row, 14, label = data[14] ,style = style)
                    ws.write(dest_row, 15, label = data[15] ,style = style)
                    ws.write(dest_row, 16, label = data[16] ,style = style)
                
                    bulletins_written.append(data[0] )
                    dest_row = dest_row + 1

        #Patch Details Page
        sheet_name = self.fmt_sheet2[0: self.fmt_sheet2.find("|")]
        ws = destination_wb.add_sheet(sheet_name)

        col_width = self.get_column_width(2)
        index = 0
        length = len(col_width)

        while index < length:
            ws.col(index).width =  int(col_width[index]) 
            index = index + 1 

        titles = self.get_column_names(2)
        index = 0
        length = len(titles)

        while index < length:
            ws.write(0, index, 
                label = titles[index] ,style = style_title)
            index = index + 1 

        self.get_patch_details()
        data = []
        rows = []
        rows = self.patch_details.keys()

        dest_row = 1
        patches_written = []
        
        print("        writing Patch details page")
        for i in rows:
            data = self.patch_details[i].split("|")
            if data[0] + data[1] not in patches_written:
                ws.write(dest_row, 0, label = data[0] ,style = style)
                ws.write(dest_row, 1, label = data[1] ,style = style)
                ws.write(dest_row, 2, label = data[2] ,style = style)
                ws.write(dest_row, 3, label = data[3] ,style = style)
                ws.write(dest_row, 4, label = data[4] ,style = style)
                ws.write(dest_row, 5, label = data[5] ,style = style)
                ws.write(dest_row, 6, label = data[6] ,style = style)
                ws.write(dest_row, 7, label = data[7] ,style = style)
                ws.write(dest_row, 8, label = data[8] ,style = style)
                ws.write(dest_row, 9, label = data[9] ,style = style)
                ws.write(dest_row, 10, label = data[10] ,style = style)
                ws.write(dest_row, 11, label = data[11] ,style = style)
                ws.write(dest_row, 12, label = data[12] ,style = style)
                ws.write(dest_row, 13, label = data[13] ,style = style)
                ws.write(dest_row, 14, label = data[14] ,style = style)
                ws.write(dest_row, 15, label = data[15] ,style = style)
                ws.write(dest_row, 16, label = data[16]  ,style = style)
                ws.write(dest_row, 17, label = data[17]  ,style = style)
                ws.write(dest_row, 18, label = data[18]  ,style = style)
                ws.write(dest_row, 19, label = data[19]  ,style = style)
            
                patches_written.append(data[0] + data[1] )
                dest_row = dest_row + 1 

        #Product
        sheet_name = self.fmt_sheet3[0: self.fmt_sheet3.find("|")]
        ws = destination_wb.add_sheet(sheet_name)

        col_width = self.get_column_width(3)
        index = 0
        length = len(col_width)

        while index < length:
            ws.col(index).width = int(col_width[index]) 
            index = index + 1 

        titles = self.get_column_names(3)
        index = 0
        length = len(titles)

        while index < length:
            ws.write(0, index, 
                label = titles[index] ,style = style_title)
            index = index + 1 

        self.get_patch_products()
        
        rows = []
        rows = self.products.keys()

        data = []    
        
        dest_row = 1
        products_written = []
        i = 0
        product = ""
        service_pack = ""
        bulletin = ""
        patch_file = ""
        patch_product = ""

        print("        writing Products page")
        for i in rows:
            data = self.products[i].split("|")
            
            product = data[2]
            service_pack = data[3]
            bulletin = data[0]
            patch_file = data[1]
            patch_product = patch_file + product
            
            if patch_product not in products_written:

                if (product.strip() != "" and service_pack.strip() != "" and bulletin.strip() != "" and patch_file.strip() != "" ):
                    ws.write(dest_row, 0, label = bulletin,style = style)
                    ws.write(dest_row, 1, label = patch_file ,style = style)
                    ws.write(dest_row, 2, label = product ,style = style)
                    ws.write(dest_row, 3, label = service_pack ,style = style)
                    ws.write(dest_row, 4, label = data[4] ,style = style)
                    ws.write(dest_row, 5, label = data[5] ,style = style)
                    ws.write(dest_row, 6, label = data[6] ,style = style)
                    
                    products_written.append(patch_product)
                    dest_row= dest_row + 1 

        #Files
        sheet_name = self.fmt_sheet4[0: self.fmt_sheet4.find("|")]
        ws = destination_wb.add_sheet(sheet_name)

        col_width = self.get_column_width(4)
        index = 0
        length = len(col_width)

        while index < length:
            ws.col(index).width = int(col_width[index]) 
            index = index + 1 

        titles = self.get_column_names(4)
        index = 0
        length = len(titles)

        while index < length:
            ws.write(0, index, 
                label = titles[index] ,style = style_title)
            index = index + 1 

        self.get_patch_files()
        rows = []
        rows = self.files.keys()

        data = []
        dest_row = 1
        patch_file_written = []

        i = 0
        patch_file = ""
        file_name = ""
        bulletin =""

        print("        writing Files page")
        for i in rows:
            if i not in self.files.keys():
                continue
            data = self.files[i].split("|")
            
            bulletin = data[0]
            patch_file = data[1]
            file_name = data[2]
            patchfile_filename = patch_file + file_name
            
            if patchfile_filename not in patch_file_written:

                if (file_name.strip() != "" and bulletin.strip() != "" and patch_file.strip() != "" ):
                    ws.write(dest_row, 0, label = bulletin,style = style)
                    ws.write(dest_row, 1, label = patch_file ,style = style)
                    ws.write(dest_row, 2, label = file_name ,style = style)
                    ws.write(dest_row, 3, label = data[3] ,style = style)
                    ws.write(dest_row, 4, label = data[4] ,style = style)
                    ws.write(dest_row, 5, label = data[5] ,style = style)
                    ws.write(dest_row, 6, label = data[6] ,style = style)
                    ws.write(dest_row, 7, label = data[7] ,style = style)

                    patch_file_written.append(patchfile_filename)
                    dest_row= dest_row + 1 

        print("        writing Registry page")
        #Registry
        sheet_name = self.fmt_sheet5[0: self.fmt_sheet5.find("|")]
        ws = destination_wb.add_sheet(sheet_name)

        col_width = self.get_column_width(5)
        index = 0
        length = len(col_width)

        while index < length:
            ws.col(index).width = int(col_width[index]) 
            index = index + 1 

        titles = self.get_column_names(5)
        index = 0
        length = len(titles)

        while index < length:
            ws.write(0, index, 
                label = titles[index] ,style = style_title)
            index = index + 1 

        self.get_registry()
        
        registry_written = []

        data = []
        dest_row = 1

        patch_file = ""
        bulletin =""

        rows = []
        rows = self.registry.keys()

        for i in rows:
            data = self.registry[i].split("|")

            bulletin = data[0]
            patch_file = data[1]
            registry = bulletin + patch_file

            if registry not in registry_written and (
               data[0] != ""     
            ):
                if (bulletin.strip() != "" and patch_file.strip() != "" ):
                    ws.write(dest_row, 0, label = bulletin,style = style)
                    ws.write(dest_row, 1, label = patch_file ,style = style)
                    ws.write(dest_row, 2, label = data[2] ,style = style)
                    ws.write(dest_row, 3, label = data[3] ,style = style)
                    ws.write(dest_row, 4, label = data[4] ,style = style)
                    ws.write(dest_row, 5, label = data[5] ,style = style)
                    ws.write(dest_row, 6, label = data[6] ,style = style)

                    registry_written.append(registry)
                    dest_row= dest_row + 1 

        #Save the file
        destination_wb.save(self.target_file)
        print("Created the file " + self.target_file)