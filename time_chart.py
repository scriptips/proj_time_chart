import pandas as pd
from pandas import DataFrame as df
import xlwings as xw
from datetime import datetime as dt, timedelta
import numpy as np
import pywintypes
import re
from xlsxwriter.utility import xl_range, xl_col_to_name
from os.path import expanduser
import datetime
import itertools
import xlsxwriter
from dateutil.relativedelta import relativedelta
import locale
locale.setlocale(locale.LC_ALL, 'latvian')

while True:
    try:
        # Connect to the range 'choose_date' in the sheet 'Countries' to get the date from the user.
        default_cutoff_date = xw.books.active.sheets['Countries'].range('choose_date').value

        print('\nChart writer in progress, please wait...\n\n')

        # Sort Countries in alphabetic order, so them as well to appear in the sheet "List".
        ws = xw.sheets['Countries']
        ctry_iso_Table = ws.api.ListObjects('ctry_iso_Table')
        ctry_iso_Table.Sort.Apply()

        # Sort Employees in alphabetic order, so them as well to appear in the sheet "List".
        ws = xw.sheets['Employees']
        empl_alp_list = ws.api.ListObjects('empl_list_Table')
        empl_alp_list.Sort.Apply()

        class Source:
            def __init__(self, my_book, my_sheet, my_data_frame):
                self.my_book = my_book
                self.my_sheet = my_sheet
                self.my_data_frame = my_data_frame

            def clean_up_df(self):
                self.my_data_frame.rename(columns=self.my_data_frame.iloc[0], inplace=True) # renames header;
                self.my_data_frame.drop(self.my_data_frame.index[0], inplace=True)  # drops first duplicate;
                self.my_data_frame.dropna(subset=['Nm', 'Fitter', 'Country'], inplace=True)  # drops rows where None in the columns;
                self.my_data_frame.drop([x for x in self.my_data_frame.columns if re.search('Column*', x)], axis='columns', inplace=True)  # drops empty columns;
                return self.my_data_frame

        class Comment:
            def __init__(self, fitter, iso, wbs, local_wo, network, begin, compl, invoiced, ken, po, site):
                self.fitter = fitter
                self.iso = iso
                self.wbs = wbs
                self.local_wo = local_wo
                self.network = network
                self.begin = begin
                self.compl = compl
                self.invoiced = invoiced
                self.ken = ken
                self.po = po
                self.site = site

            def insert_comment(self):
                """Creates line-by-line comments\n
                of the preselected fields from\n
                the "List" excel sheet. """
                return f"Fitter: {self.fitter}\nStarted: {self.begin}\nSite: {self.site}\nWO: {self.local_wo}\nNetw: {self.network}\nKEN: {self.ken}\nWBS: {self.wbs}\nPO: {self.po}\nInvoiced: {self.invoiced}"
            def get_cbox_dims(self):
                obj_attr_list = [x.replace(' ','') for x in [str(self.fitter), str(self.begin), str(self.site), str(self.local_wo), str(self.network), str(self.ken), str(self.wbs), str(self.po), str(self.invoiced)]]
                w = len(max(obj_attr_list, key=len)) * 10.667
                h = len(obj_attr_list) * 17
                return w, h

            def make_iso(self, country_name):
                ctry_iso = df(xw.books.active.sheets['Countries'].range('ctry_iso').value)
                countries = ctry_iso[0].tolist()
                isos = ctry_iso[1].tolist()
                dicts = {countries: isos for countries, isos in zip(countries, isos)}
                return dicts.get(country_name) if self.invoiced == 'n/a' else dicts.get(country_name) + ' '

        # Manually instantiate below four objects.
        fdates = Source(xw.books.active, xw.sheets['List'], df(xw.books.active.sheets['List'].range('ListTable[[#All]]').value))
        fnames = Source(xw.books.active, xw.sheets['List'], df(xw.books.active.sheets['List'].range('ListTable[[#All]]').value))
        fcountr = Source(xw.books.active, xw.sheets['List'], df(xw.books.active.sheets['List'].range('ListTable[[#All]]').value))
        cfields = Source(xw.books.active, xw.sheets['List'], df(xw.books.active.sheets['List'].range('ListTable[[#All]]').value))

        # Call the method via list compreh. to clean-up the dfs (empty columns, first row, rename header).
        [c.clean_up_df() for c in [fdates, fnames, fcountr, cfields]]

        # Format comment field to show whole numbers, dates in friendly format, and n/a in empty cells.
        na_list = ['WBS Element', 'Local WO', 'Network', 'PO Amount, EUR', 'Equip.Nr.', 'PO Number', 'Site']
        for na in na_list:
            if na in cfields.my_data_frame.head():
                cfields.my_data_frame[na] = cfields.my_data_frame[na].fillna('n/a')
                cfields.my_data_frame[na] = cfields.my_data_frame[na].apply(lambda c: int(c) if type(c) == float else str(c))
            else: continue
                # Jump to the next iteration of the loop, if nothing is found.

        ts_list = ['Invoiced']
        for ts in ts_list:
            cfields.my_data_frame[ts] = cfields.my_data_frame[ts].fillna('n/a')
            cfields.my_data_frame[ts] = cfields.my_data_frame[ts].apply(lambda d: d.strftime('%d.%m.%Y') if type(d) != str else str(d))

        if xw.books.active.sheets['Countries'].range('choose_date').value is None:
            report_cutoff_date = datetime.datetime.date(min(fdates.myDataFrame['Start'], default=datetime.datetime(2000, 1, 1)))
        else: report_cutoff_date = xw.books.active.sheets['Countries'].range('choose_date').value

        this_day = datetime.datetime(int(dt.today().date().strftime('%Y')), int(dt.today().date().strftime('%m')), int(dt.today().date().strftime('%d')))

        # Determine the works that were started before the cutoff. It is needed to add comments to them later as well.
        cfields.my_data_frame['End'] = cfields.my_data_frame['End'].apply(lambda e: this_day if e is pd.NaT else e)

        cfields.my_data_frame['Start'] = cfields.my_data_frame['Start'].fillna('n/a')

        # Format start date to show only the date, not the time
        cfields.my_data_frame['Start'] = cfields.my_data_frame['Start'].apply(lambda bd: bd.strftime('%d.%m.%Y') if type(bd) != str else bd)
        cfields.my_data_frame = cfields.my_data_frame[cfields.my_data_frame['Start'] != '']

        # Instantiate objects in the list.
        comm_objs = []
        for x in cfields.my_data_frame.iterrows():
            ''' <=... lai iekļautu komentārus, kas attiecas uz iesāktajiem, bet vēl nepabeigtajiem darbiem'''
            if report_cutoff_date <= x[1][8]:
                comm_objs.append(Comment(x[1][1], x[1][2], x[1][3], x[1][4], x[1][5], x[1][7], x[1][8], x[1][9], x[1][10], x[1][11], x[1][12]))

        ### Control of the rolling 182 days starts here.###
        now = datetime.datetime.now()
        now = now.replace(hour=0, minute=0, second=0, microsecond=0)  # drop hh:mm:ss to zero.
        fdates.my_data_frame['End'] = fdates.my_data_frame['End'].fillna(now)
        fdates.my_data_frame[10] = None
        d182 = []

        for df_rw in fdates.my_data_frame.iterrows():
            start = df_rw[1]['Start']
            end = df_rw[1]['End']
            one_year = relativedelta(days=365)
            one_day = relativedelta(days=1)
            # in case if the start and the end dates are in the future, the rolling 182 days will be zero.
            if start > now and end > now:  
                df_rw[1][10] = 0
            else:
                if now > end:
                    if start > (now - one_year):
                        df_rw[1][10] = (end + one_day) - start
                    else:
                        if start <= (now - one_year) and end >= (now - one_year):
                            df_rw[1][10] = end - (now - one_year)
                        else:
                            if (now - one_year) > end:
                                df_rw[1][10] = 0
                else:
                    if start > (now - one_year):
                        df_rw[1][10] = (now + one_day) - start
                    else:
                        if start <= (now - one_year) and end >= (now - one_year):
                            df_rw[1][10] = now - (now - one_year)
                        else:
                            if (now - one_year) > end:
                                df_rw[1][10] = 0
            if type(df_rw[1][10]) == int:
                pass
            else: df_rw[1][10] = df_rw[1][10].days
            d182.append(df_rw[1][10])
        fnames.my_data_frame['182 spent'] = d182
        ### Control of the rolling 182 days ends here.###


        # To avoid "str" and "float" in the same column, because the column starts with the header, then only dates.
        stRow_in_Dates_df = fdates.my_data_frame.iloc[0:]

        # Add a condition, if the field is empty, then the report will be generated from all available data from the earliest registered "Start" date.
        tagad = pd.Timestamp(now.date())
        if xw.books.active.sheets['Countries'].range('choose_date').value == None:
            minD = stRow_in_Dates_df['Start'].min(axis=0)
        else: minD = xw.books.active.sheets['Countries'].range('choose_date').value
        maxD = stRow_in_Dates_df['End'].max(axis=0) 
        if tagad > maxD:
            maxD = tagad

        # Define the date field for the time scale (column headers) in pd.
        dateRange = pd.date_range(start=minD, end=maxD + timedelta(days=1) , freq='D')

        dateArray = df(columns=dateRange)

        # Define data field for the country list in pd.
        fCountrRange = pd.DataFrame(fnames.my_data_frame['Country'])
        fCountrRange = df.drop_duplicates(fCountrRange, keep='first')

        # Convert the named field in xlsx as a py dictionary.
        ctry_iso_dict = xw.books.active.sheets['Countries'].range('ctry_iso').options(dict).value

        fCountrRange = pd.DataFrame(ctry_iso_dict.values())
        fCountrRangeT = fCountrRange.T
        fCountrRangeT.columns = fCountrRangeT.iloc[0]
        fCountrRangeT.drop(fCountrRange,inplace=True)

        # Merge the two dataframes.
        allTable = pd.concat([fnames.my_data_frame, fCountrRangeT, dateArray], axis=1)

        allTable['Country'] = allTable['Country'].map(ctry_iso_dict)

        temp_ctry_rng = allTable.iloc[:, (len(allTable.columns) - len([x.to_pydatetime().date() for x in allTable.columns.values if type(x) == pd.Timestamp]) - len(fCountrRange)):
                                        (len(allTable.columns) - len([x.to_pydatetime().date() for x in allTable.columns.values if type(x) == pd.Timestamp]))]

        allTable.loc[temp_ctry_rng.index.values, temp_ctry_rng.columns.values] = temp_ctry_rng.columns.values

        # To avoid job numbers (networks) to be printed as floats.
        iso_netw_switch = allTable['Country'].astype(str)
        iso_netw_switch_cwidth = 3.5

        z = len(allTable.columns) - len(dateRange)
        q = z - len(fCountrRange)

        # Fill the 182 day control field in the report.
        for h, item in enumerate(temp_ctry_rng.columns.values):
            conditions = [
                (allTable['Country'] == temp_ctry_rng.columns[h]),
                (allTable['Country'] != temp_ctry_rng.columns[h])]
            choices = [allTable['182 spent'],0]
            allTable[allTable.columns[h+q]] = np.select(conditions, choices, default=0)

        # select and fill ISO codes in pd df.
        for x, item in enumerate(dateRange):
            conditions = [
                (allTable['Start'] <= dateRange[x]) & (allTable['End'] > dateRange[x]) & (allTable['Invoiced'] >= dateRange[x]),
                (allTable['Start'] <= dateRange[x]) & (allTable['End'] > dateRange[x]) & (allTable['Invoiced'] < dateRange[x]),
                (allTable['Start'] <= dateRange[x]) & (allTable['End'] > dateRange[x]),
                (allTable['Start'] > dateRange[x]) & (allTable['End'] < dateRange[x])]
            choices = [allTable['Country'] + ' ', allTable['Country'] + ' ', allTable['Country'], '']
            allTable[allTable.columns[z+x]] = np.select(conditions, choices, default='')

        # Drop redundant columns as no longer needed.
        # There might be different headers of different source files, therefore try to drop all possible headers.
        drop_head = ['Nm', ' ',  'Country', 'WBS Element', 'Local WO', 'Network', 'PO Amount, EUR', 'Start', 'End', 'Invoiced', 'Equip.Nr.', 'PO Number', 'Site Address', 'Comm', '182 spent']
        spec_drop_head = []
        for item in drop_head:
            if item not in allTable.columns:
                continue
            spec_drop_head.append(item)
        allTable.drop(columns = spec_drop_head , inplace=True)

        group_it = allTable.groupby('Fitter').agg(lambda w : w.sum() if w.dtype != 'str' else ' '.join(w))

        # Sort employee names in alphabetical order.
        # Create list from group_it index (which is fitter names).
        group_it_suplement_list = [x for x in group_it.index]
        # Sort list in LV alphabetical order.
        group_it_suplement_list.sort(key=locale.strxfrm)

        # Below .loc will accept duplicates and even then would do the sorting.
        group_it = group_it.loc[group_it_suplement_list]  
        # Drop the so-far index.
        group_it.reset_index(inplace = True)

        # Counts employees.
        fitterCount = group_it.iloc[:,1].count()+1 
        # Creates a field of employee count.
        fitterCount = range(1,fitterCount,1)
        # Adds column 'Nr.' upfront the data field.
        group_it.insert(0,'Nr.',fitterCount, True) 
        # Sets 'Nr.' as index.
        group_it = group_it.set_index(['Nr.']) 

        laiks = dt.now().strftime("%H.%M.%S")

        # Dfine the number of columns for the date-cols.
        dRange_in_group_it = len(group_it.columns)-len(fCountrRange)
        home = expanduser("~")
        nm = f"{dt.now().date().strftime('%d-%m-%Y')} ({laiks})"
        temp_wb_fpath = f'{{}}\\OneDrive - KONE Corporation\\Desktop\\Time Chart {nm}.xlsx'

        with pd.ExcelWriter(temp_wb_fpath.format(home),
                                engine='xlsxwriter',
                                datetime_format='dd.mm.yyyy',
                                date_format='dd.mm.yyyy') as writer:
            try:
                # Reset the index and store it as a separate column.
                group_it_with_index = group_it.reset_index()

                # Write the modified DataFrame to Excel without formatting the index.
                group_it_with_index.to_excel(writer, sheet_name=nm, startrow=2, index=False, header=True)

                xrt_wb  = writer.book
                xrt_ws = writer.sheets[nm]

                other_but_date_headers = len(group_it.columns.values)-len([x.to_pydatetime().date().strftime('%d.%m.%Y') for x in group_it.columns.values if type(x) != str])

                for co in comm_objs:                
                    date_header = [x.date() for x in pd.date_range(minD, maxD+timedelta(days=1))].index((co.compl).date())
                    fitter_index = group_it.index[group_it['Fitter'] == co.fitter]
                    write_end_cell_comment = (f"${str(xlsxwriter.utility.xl_col_to_name(other_but_date_headers + date_header + 1))}${fitter_index[0] + 3}")
                    xrt_ws.write(write_end_cell_comment, co.make_iso(co.iso))
                    xrt_ws.write_comment(write_end_cell_comment, co.insert_comment(),
                                        {'width': co.get_cbox_dims()[0],
                                        'height': co.get_cbox_dims()[1]})
                # A variable to avoid absolute references.
                cond_form_rng = xl_range(2, 0, len(group_it.index) + 2, len(group_it.columns))
                cond_form_rng_iso_only = xl_range(2, 2 + len(fCountrRange), len(group_it.index) + 2, len(group_it.columns))

                temp = datetime.datetime(1899, 12, 30) # Note, not 31st Dec, but rather 30th.
                delta = dt.now() - temp
                excel_date = (now - datetime.datetime(1899, 12, 30)).days + (now - datetime.datetime(1899, 12, 30)).seconds / 86400 + 18/24

                # Format the imported df.
                header_format_dates = xrt_wb.add_format({
                    'italic': True,
                    'valign': 'vcenter',
                    'align' : 'center',
                    'num_format':'dd/mm/yy',
                    'rotation': 90})

                for col_num, value in enumerate(group_it.columns):
                    xrt_ws.write(2, col_num + 1 , value, header_format_dates)

                # Format df until date-columns, as there is the 90deg rotation.
                until_dates_rng = xl_range(2, 0, 2, len(fCountrRange))

                until_dates_format = xrt_wb.add_format({'italic': False,
                                                        'align' : 'left',
                                                        'rotation': 0})

                xrt_ws.conditional_format(until_dates_rng, {'type': 'formula',
                                                            'criteria':'=A$3<>""',
                                                            'format':until_dates_format})

                for col_num, value in enumerate (itertools.islice(group_it.columns, len(fCountrRange) + 1)):
                    xrt_ws.write(2, col_num + 1 , value, until_dates_format)

                # Red color vertical line indicating today's date.
                red_vertic_form = xrt_wb.add_format()
                red_vertic_form.set_right(5)
                red_vertic_form.set_right_color('#FF0000')
                xrt_ws.conditional_format(cond_form_rng,{'type': 'formula',
                                                    'criteria':f'=A$3={excel_date}',
                                                    'format': red_vertic_form})

                # Prepare new dicts, where capture excel cell properties and convert to RGB -> HEX.
                my_d = {}
                for x in xw.books.active.sheets['Countries'].range('iso'):
                    my_d[x] = '#{:02x}{:02x}{:02x}'.format(x.color[0],x.color[1],x.color[2])

                # Dict with ISO codes.
                my_d2 = {}
                for y in xw.books.active.sheets['Countries'].range('iso'):
                    my_d2[y] = y.value

                # Zip xl iso country region with the newly created dict where HEX colour codes are in.
                iso_color_format_dict = dict(zip(my_d2.values(), my_d.values()))

                # Paint the ISO codes in the newly created spreadsheet by xlsxwriter.
                for key, value in iso_color_format_dict.items():
                    format1 = xrt_wb.add_format({'bg_color': value,
                                                'font_color': '#000000'})
                    xrt_ws.conditional_format(cond_form_rng,{'type':'cell',
                                                            'criteria':'=',
                                                            'value': f'"{key}"',
                                                            'format':format1})
                    format2 = xrt_wb.add_format({'bg_color': 'D9D9D9',
                                                'font_color': '#ffffff'})
                    xrt_ws.conditional_format(cond_form_rng,{'type':'cell',
                                                            'criteria':'=',
                                                            'value': f'"{key} "',
                                                            'format':format2})
                # Paint weekend cells in pink.
                pink_weekends_form = xrt_wb.add_format({'bg_color': '#fce4d6',
                                                        'font_color': '#000000',})

                pink_wk = datetime.datetime.today().weekday()+1+1
                xrt_ws.conditional_format(cond_form_rng,{'type': 'formula',
                                                    'criteria':f'=weekday(A$3,2)>=6',
                                                    'format': pink_weekends_form})

                # Determine the longest string in the Employee column.
                fitt_name_str_max = group_it['Fitter'].str.len().max()

                xrt_ws.set_column("A:"+xl_col_to_name(len(group_it.columns)),iso_netw_switch_cwidth)
                xrt_ws.set_column(xl_col_to_name(1)+":"+xl_col_to_name(1), fitt_name_str_max)
                xrt_ws.freeze_panes(3, 2 + len(fCountrRange))

                # 182 day control Summary area.
                spent_182_print_area = xl_range(3, 2, (len(fitterCount) + 2), len(fCountrRange) + 1)

                # Remove zeroes from the 182 day control area.
                remove_zeroes = xrt_wb.add_format({'num_format': '#'})
                xrt_ws.conditional_format(spent_182_print_area, {'type': 'cell',
                                                                'criteria':'equal to',
                                                                'value': 0,
                                                                'format':remove_zeroes})
                # Mark cells with 182d+ with red color.
                paint_over_151ds = xrt_wb.add_format({'num_format': '#',
                                                    # 'bg_color': '#fce4d6',
                                                    'font_color': '#FF0000',
                                                    'bold': True})
                xrt_ws.conditional_format(spent_182_print_area, {'type': 'cell',
                                                                'criteria':'>',
                                                                'value': 151,
                                                                'format':paint_over_151ds})
                # Set filter on the Employees column.
                xrt_ws.autofilter('B3:B3')

                format_tchart_header_timestamp = xrt_wb.add_format({'num_format': '#',
                                                    'bg_color': '#fce4d6',
                                                    'font_color': '#FF0000'})

                format_tchart_header_timestamp = xrt_wb.add_format()
                format_tchart_header_timestamp.set_font_color('#FF0000')
                format_tchart_header_timestamp.set_italic()
                tchart_created = f"Time Chart prepared on: {dt.now().date().strftime('%d.%m.%y')} (at: {dt.now().strftime('%H:%M')})"
                xrt_ws.write('B1', tchart_created, format_tchart_header_timestamp)

            except FileNotFoundError as e:
                print(f'\n{e}\n\n')

        default_cutoff_date = tagad - datetime.timedelta(3 * 365 / 12)
        xw.books.active.sheets['Countries'].range('choose_date').value = default_cutoff_date

        targ_wb = xw.Book(temp_wb_fpath.format(home))
        # Convert from notes to threaded comments (new style enhanced comments)
        targ_wb.api.ConvertComments() 
        targ_wb.save()
        break
    except (AttributeError, xw.XlwingsError):
        input('\nERROR: Source File may not be open, try again...\n')
    except  pywintypes.com_error:
        input('\nERROR: Click in the Source File to make the workbook active and try again...\n')