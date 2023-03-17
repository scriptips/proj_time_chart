from pandas import pandas as pd, DataFrame as df
import xlwings as xw
from datetime import datetime as dt
import numpy as np
import pywintypes
import re
from xlsxwriter.utility import xl_range, xl_col_to_name
from os.path import expanduser
import datetime
import itertools
import xlsxwriter
from dateutil.relativedelta import relativedelta
import locale  # LV listes zortēšanai

pd.set_option('display.max_columns', None)

locale.setlocale(locale.LC_ALL, 'latvian')

while True:
    try:
        default_cutoff_date = xw.Range('choose_date').value
        break
    except AttributeError:
        print('')
        input('KĻŪDA: Vispirms jāatver "Darbs Ārzemēs" fails. Kad atvērts, spied "Enter" lai mēģinātu vēlreiz...')
    except  pywintypes.com_error:
        print('')
        input('KĻŪDA: Ieklikšķināt faila "Darbs Ārzemēs ..."  lapā lai to aktivizētu, un spiest "Enter" lai turpinātu.')

print('')
print('Mirkli uzgaidīt, tiek gatavots grafiks...')
print('')


# Sortē VALSTIS alfabētiski, lai dropdown lapā "List" arī būtu alf.kārtībā.
ws = xw.sheets['Countries']
ctry_iso_Table = ws.api.ListObjects('ctry_iso_Table')
ctry_iso_Table.Sort.Apply()

# Zortē DARBINIEKUS alfabētiski, lai dropdown lapā "List" arī būtu alf.kārtībā.
ws = xw.sheets['Employees']
empl_alp_list = ws.api.ListObjects('empl_list_Table')
empl_alp_list.Sort.Apply()


class Source:
    def __init__(self, my_book, my_sheet, my_data_frame):
        self.my_book = my_book
        self.my_sheet = my_sheet
        self.my_data_frame = my_data_frame

    def clean_up_df(self):
        self.my_data_frame.rename(columns=self.my_data_frame.iloc[0], inplace=True) #pārsauc galveni
        self.my_data_frame.drop(self.my_data_frame.index[0], inplace=True)  # nomet pirmo duplikātu
        self.my_data_frame.dropna(subset=['Nm', 'Fitter', 'Country'], inplace=True)  # izmet rindas ja kolonnās None
        self.my_data_frame.drop([x for x in self.my_data_frame.columns if re.search('Column*', x)], axis='columns', inplace=True)  # izmet tukšās kolonnas
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

    def insert_comment(self) -> str('Select comment fields to be later included in the comment boxes in "Time Chart" sheet'):
        """Creates line-by-line comments of the preselected fields from the "List" excel sheet. """
        return f"Fitter: {self.fitter}\nStarted: {self.begin}\nSite: {self.site}\nWO: {self.local_wo}\nNetw: {self.network}\nKEN: {self.ken}\nWBS: {self.wbs}\nPO: {self.po}\nInvoiced: {self.invoiced}"
    def get_cbox_dims(self):
        obj_attr_list = [x.replace(' ','') for x in [str(self.fitter), str(self.begin), str(self.site), str(self.local_wo), str(self.network), str(self.ken), str(self.wbs), str(self.po), str(self.invoiced)]]
        w = len(max(obj_attr_list, key=len)) * 10.667
        h = len(obj_attr_list) * 17
        return w, h

    def make_iso(self, country_name):
        ctry_iso = df(xw.Range('ctry_iso').value)
        countries = ctry_iso[0].tolist()
        isos = ctry_iso[1].tolist()
        dicts = {countries: isos for countries, isos in zip(countries, isos)}
        # Note that Python supports a syntax that allows you to use only one return statement in your case
        return dicts.get(country_name) if self.invoiced == 'n/a' else dicts.get(country_name) + ' '


# Izveido jaunus objektus.

fDates = Source(xw.books.active, xw.sheets['List'], df(xw.Range('ListTable[[#All]]').value))   #xw.books.active
fNames = Source(xw.books.active, xw.sheets['List'], df(xw.Range('ListTable[[#All]]').value))
fCountr = Source(xw.books.active, xw.sheets['List'], df(xw.Range('ListTable[[#All]]').value))
cFields = Source(xw.books.active, xw.sheets['List'], df(xw.Range('ListTable[[#All]]').value))

fDates.clean_up_df()
fNames.clean_up_df()
fCountr.clean_up_df()
cFields.clean_up_df()



# Formatē komentāru lauku lai cipari rādītos veselos skaitļos, datumi draudzīgā formātā, un tukšumos n/a....
na_list = ['WBS Element', 'Local WO', 'KLA Network', 'PO Amount, EUR', 'KEN', 'PO Number', 'Site']
for na in na_list:
    if na in cFields.my_data_frame.head():
        cFields.my_data_frame[na] = cFields.my_data_frame[na].fillna('n/a')
        cFields.my_data_frame[na] = cFields.my_data_frame[na].apply(lambda c: int(c) if type(c) == float else str(c))
    else: continue
        # lec pāri listes vienumam, ja netiek atrasts


ts_list = ['Invoiced']
for ts in ts_list:
    cFields.my_data_frame[ts] = cFields.my_data_frame[ts].fillna('n/a') # svarīgi, šī rinda saistīta
    # ar klasses metodi make_iso(self, country_name)
    cFields.my_data_frame[ts] = cFields.my_data_frame[ts].apply(lambda d: d.strftime('%d.%m.%Y') if type(d) != str else str(d))


if xw.Range('choose_date').value is None:
    report_cutoff_date = datetime.datetime.date(min(fDates.myDataFrame['Start'], default=datetime.datetime(2000, 1, 1)))
else: report_cutoff_date = xw.Range('choose_date').value


this_day = datetime.datetime(int(dt.today().date().strftime('%Y')), int(dt.today().date().strftime('%m')), int(dt.today().date().strftime('%d')))
# nosaka tos darbus, kuri uz cutoff datumu bijuši atvērti. Tas vajadzīgs lai vēlāk pievienotu komentārus arī tiem...
cFields.my_data_frame['End'] = cFields.my_data_frame['End'].apply(lambda e: this_day if e is pd.NaT else e)


cFields.my_data_frame['Start'] = cFields.my_data_frame['Start'].fillna('n/a')
# formatē sākuma datumu, lai vēlāk celles komentārā draudzīgi parādītos..
cFields.my_data_frame['Start'] = cFields.my_data_frame['Start'].apply(lambda bd: bd.strftime('%d.%m.%Y') if type(bd) != str else bd)
cFields.my_data_frame = cFields.my_data_frame[cFields.my_data_frame['Start'] != '']


# instantiate objects through the list
comm_objs = []
for x in cFields.my_data_frame.iterrows():
    ''' <=... lai iekļautu komentārus, kas attiecas uz iesāktajiem, bet vēl nepabeigtajiem darbiem'''
    if report_cutoff_date <= x[1][8]:
        comm_objs.append(Comment(x[1][1], x[1][2], x[1][3], x[1][4], x[1][5], x[1][7], x[1][8], x[1][9], x[1][10], x[1][11], x[1][12]))


# #####################################################################################################################
# 182 dienu kalkulācija pd.df. Šī sadaļa izņemta no excel formulām, lai novērstu kalkulāciajs nobrukšanu netīšas excel formulu izjaukšanas rezultātā
now = datetime.datetime.now()
now = now.replace(hour=0, minute=0, second=0, microsecond=0)  # nomet laiku pa nullēm..
fDates.my_data_frame['End'] = fDates.my_data_frame['End'].fillna(now)
fDates.my_data_frame[10] = None
d182 = []

for df_rw in fDates.my_data_frame.iterrows():
    start = df_rw[1]['Start']
    end = df_rw[1]['End']
    one_year = relativedelta(days=365)
    one_day = relativedelta(days=1)
    if start > now and end > now:  # ja darbus plānots sākt & beigt nākotnē
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
fNames.my_data_frame['182 spent'] = d182
#  Šeit noslēdzas 182 dienu limita kalkulācija python, kas iepriekš bija excel formulās failā "Darbs ārzemēs"
# ###########################################################################################################


# Lai izvairītos no tā ka "str" un "float" vienā kolonnā, jo kolonna sākas ar virsrakstu, pēc tam tikai datumi...
stRow_in_Dates_df = fDates.my_data_frame.iloc[0:]

# Pievieno nosacījumu, ja atstāj tukšu lauku, tad atskaite ģenerēsies no visiem pieejamajiem datiem no senākā reģistrētā "Start" datuma.
tagad = dt.today().date()
if xw.Range('choose_date').value == None:
    minD = stRow_in_Dates_df['Start'].min(axis=0)
else: minD = xw.Range('choose_date').value
maxD = stRow_in_Dates_df['End'].max(axis=0)   #+datetime.timedelta(1*365/12)
if tagad > maxD:
    maxD = tagad

# Definē datu lauku laika skalai (kolonnu galvenēm) Pandās.
dateRange = pd.date_range(start=minD, end=maxD, freq='D')
dateArray = df(columns=dateRange)




# Definē datu lauku valstu listei Pandās.
fCountrRange = pd.DataFrame(fNames.my_data_frame['Country'])
fCountrRange = df.drop_duplicates(fCountrRange, keep='first')

# XlWings konvertē excel nosaukto lauku kā Python vārdnīcu.
ctry_iso_dict = xw.Range('ctry_iso').options(dict).value


fCountrRange = pd.DataFrame(ctry_iso_dict.values())
fCountrRangeT = fCountrRange.T
fCountrRangeT.columns = fCountrRangeT.iloc[0]
fCountrRangeT.drop(fCountrRange,inplace=True)


# !! Apvieno izveidotos datu laukus vienā Pandā.
allTable = pd.concat([fNames.my_data_frame, fCountrRangeT, dateArray], axis=1)

# Svarīgi: junkcija "map". šeit definē jaunu allTable, nevis pieškirt jaunu property..., ja nē, nerāda tabula pareizi.
allTable['Country'] = allTable['Country'].map(ctry_iso_dict)

# 31.01.2022 - dinamiski nosaka valstu kolonnu atrašanās vietu tabulā
temp_ctry_rng = allTable.iloc[:, (len(allTable.columns) - len([x.to_pydatetime().date() for x in allTable.columns.values if type(x) == pd.Timestamp]) - len(fCountrRange)):
                                 (len(allTable.columns) - len([x.to_pydatetime().date() for x in allTable.columns.values if type(x) == pd.Timestamp]))]

allTable.at[temp_ctry_rng.index.values, temp_ctry_rng.columns.values] = temp_ctry_rng.columns.values


# Savādāk rādīsies kā : ..777.0 kā float.. Varbūt ne tik aktuāli, ja rāda tikai valstu kodus (un ne Networkus), bet atstāts turpmākai zināšanai.
iso_netw_switch = allTable['Country'].astype(str)
iso_netw_switch_cwidth = 3.5

z = len(allTable.columns) - len(dateRange)
q = z - len(fCountrRange)

# Aizpilda 182dienas kontroles apgabalu atskaitē.
for h, item in enumerate(temp_ctry_rng.columns.values):
    conditions = [
        (allTable['Country'] == temp_ctry_rng.columns[h]),
        (allTable['Country'] != temp_ctry_rng.columns[h])]
    choices = [allTable['182 spent'],0]
    allTable.at[allTable.index.values,allTable.columns[h+q]] = np.select(conditions, choices, default=0)


# Saraksta ISO kodus Pandas.
for x, item in enumerate(dateRange):
    conditions = [
        (allTable['Start'] <= dateRange[x]) & (allTable['End'] > dateRange[x]) & (allTable['Invoiced'] >= dateRange[x]),
        (allTable['Start'] <= dateRange[x]) & (allTable['End'] > dateRange[x]) & (allTable['Invoiced'] < dateRange[x]),
        (allTable['Start'] <= dateRange[x]) & (allTable['End'] > dateRange[x]),
        (allTable['Start'] > dateRange[x]) & (allTable['End'] < dateRange[x])]
    choices = [allTable['Country'] + ' ', allTable['Country'] + ' ', allTable['Country'], '']
    allTable.at[allTable.index.values,allTable.columns[z+x]] = np.select(conditions, choices, default='')


# Lieko palīgkolonnu nomešana, kā tālāk nevajadzīgu lauku samazinašana.
# CNEFP un KLA var atšķirties tabulas galvenes, tāpēc sākumā mēģina visus iespējamos, un tad atstāj listē tikai esošos_
drop_head = ['Nm', ' ',  'Country', 'WBS Element', 'Local WO', 'KLA Network', 'PO Amount, EUR', 'Start', 'End', 'Invoiced', 'KEN', 'PO Number', 'Site Address', 'Comm', '182 spent']
spec_drop_head = []
for item in drop_head:
    if item not in allTable.columns:
        continue
    spec_drop_head.append(item)
allTable.drop(columns = spec_drop_head , inplace=True)

# Atstājot w.dtype!='str' uz 'int' conditional formatting nestradāja, tika uztaustīts. Uzzināt vairak par lambdām.
group_it = allTable.groupby('Fitter').agg(lambda w : w.sum() if w.dtype != 'str' else ' '.join(w))


# FITTERU VĀRDU SAKĀRTOŠANA LV ALFABĒTISKĀ SECĪBĀ
group_it_suplement_list = [x for x in group_it.index]  # izveido listi no group_it df indeksa (kas ir fitteru vārdi)
group_it_suplement_list.sort(key=locale.strxfrm)  # zortē listi LV alfabētiskā kārtībā
group_it = group_it.loc[group_it_suplement_list]  # sakārto indeksu (kas uz doto brīdi ir fitteru vārdi) LV alfabētiskā kārtībā
# augšminētais pielietojums ar .loc ļaus būt duplikātiem indeksā un arī tad zortēs.
group_it.reset_index(inplace = True) #nomet līdzšinejo indeksu

fitterCount = group_it.iloc[:,1].count()+1 #saskaita Fitterus
fitterCount = range(1,fitterCount,1) #izveido fitteru skaita apgabalu
group_it.insert(0,'Nr.',fitterCount, True) #pieliek kolonnu "Nr." datu apgabala sakumā
group_it = group_it.set_index(['Nr.']) #nomaina līdzšinejo indexu ar kolonnu "Nr."

laiks = dt.now().strftime("%H.%M.%S")

# Definē kolonnu skaitu datumu galvenēm.
dRange_in_group_it = len(group_it.columns)-len(fCountrRange)
home = expanduser("~")
nm = f"{dt.now().date().strftime('%d-%m-%Y')} ({laiks})"
temp_wb_fpath = f'{{}}\\OneDrive - KONE Corporation\\Desktop\\Time Chart {nm}.xlsx'

# Noņem Pandu uzspiesto galvenes formātu Xlsx ieeksportētajā datu laukā (bija grūti sameklēt, kā arī var mainīties ar nākamajiem laidieniem).
pd.io.formats.excel.ExcelFormatter.header_style = None


try:
    writer = pd.ExcelWriter(temp_wb_fpath.format(home),
                        engine='xlsxwriter',
                        datetime_format='dd.mm.yyyy',
                        date_format='dd.mm.yyyy')
except FileNotFoundError as e:
    print('')
    print(e)
    print('')
    input('')


#WXXXXXXX ŠEIT galīgais pd.df variants pirms rakstīšanas excelī XXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
# Konsolidētās tabulas nodošana Xlsxwriter.
group_it.to_excel(writer, sheet_name=nm, startrow=2, index=True, header=True)


xrt_wb  = writer.book
xrt_ws = writer.sheets[nm]

other_than_date_headers = len(group_it.columns.values)-len([x.to_pydatetime().date().strftime('%d.%m.%Y') for x in group_it.columns.values if type(x) != str])


for co in comm_objs:
        date_header = [x.to_pydatetime().date() for x in group_it.columns.values if type(x) == pd.Timestamp].index(co.compl)
        fitter_index = group_it.index[group_it['Fitter'] == co.fitter]
        write_end_cell_comment = (f"${str(xlsxwriter.utility.xl_col_to_name(other_than_date_headers + date_header + 1))}${fitter_index[0] + 3}")
        xrt_ws.write(write_end_cell_comment, co.make_iso(co.iso))
        xrt_ws.write_comment(write_end_cell_comment, co.insert_comment(),
                             {'width': co.get_cbox_dims()[0],
                              'height': co.get_cbox_dims()[1]})


# Lokālais mainigais lai izvairītos no absolūtajām referencēm.
cond_form_rng = xl_range(2, 0, len(group_it.index) + 2, len(group_it.columns))
cond_form_rng_iso_only = xl_range(2, 2 + len(fCountrRange), len(group_it.index) + 2, len(group_it.columns))

# Pārvērš Python datetime uz excel datenumber. Oriģinalajā StackOverflow piemērā bija definēta metode, bet šeit pārvērsta sekvenciālā kodā.

temp = datetime.datetime(1899, 12, 30)# Note, not 31st Dec but 30th!
delta = dt.now() - temp
final = float(delta.days) + (-1.0 if delta.days < 0 else 1.0)*(delta.seconds) / 86400


# Formatē XlsxWriter ieeksportētā Pandas datu lauka galveni.
header_format_dates = xrt_wb.add_format({
    'italic': True,
    'valign': 'vcenter',
    'align' : 'center',
    'num_format':'dd/mm/yy',
    'rotation': 90})

for col_num, value in enumerate(group_it.columns):
    xrt_ws.write(2, col_num + 1 , value, header_format_dates)

# Formatē apgabalu līdz datumu kolonnām, jo šeit jānomet 90grādu rotācija tekstam.
until_dates_rng = xl_range(2, 0, 2, len(fCountrRange))

until_dates_format = xrt_wb.add_format({'italic': False,
                                        'align' : 'left',
                                        'rotation': 0})

xrt_ws.conditional_format(until_dates_rng, {'type': 'formula',
                                            'criteria':'=A$3<>""',
                                            'format':until_dates_format})

for col_num, value in enumerate (itertools.islice(group_it.columns, len(fCountrRange) + 1)):
    xrt_ws.write(2, col_num + 1 , value, until_dates_format)


# Sarkanā "tagadnes vertikāle":).
red_vertic_form = xrt_wb.add_format()
red_vertic_form.set_right(5)
red_vertic_form.set_right_color('#FF0000')
xrt_ws.conditional_format(cond_form_rng,{'type': 'formula',
                                     'criteria':f'=A$3={int(final)}',
                                     'format': red_vertic_form})

# Izveido jaunas vārdnīcas, kurās savāc excel celles un vienlaicigi tās konvertē no RGB uz HEX.
# Svarīgi, ka sākumā zipojot Range(kurš tika interpretēts kā liste, sanāca kļūda, ja atskaiti testējot atstāja tikai vienu valsti, bet tagad ok.
# Jauna vārdnīca + Cilpa ar pašu RGB<->HEX koncertāciju.
my_d = {}
for x in xw.Range('iso'):
  my_d[x] = '#{:02x}{:02x}{:02x}'.format(x.color[0],x.color[1],x.color[2])

# Vēl viena vārdnīca ar ISO kodiem.
my_d2 = {}
for y in xw.Range('iso'):
  my_d2[y] = y.value

# Sazipo xl iso valstu reģionu ar jauniegūto vārdnīcu, kurā ir HEX krāsu kodi.
iso_color_format_dict = dict(zip(my_d2.values(), my_d.values()))


# Iekrāso ISO kodus XLSXWRITER jaunizveidotajā izklājlapā.
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

# Pink Weekends:).
pink_weekends_form = xrt_wb.add_format({'bg_color': '#fce4d6',
                                        'font_color': '#000000',})

pink_wk = datetime.datetime.today().weekday()+1
xrt_ws.conditional_format(cond_form_rng,{'type': 'formula',
                                     'criteria':f'=weekday(A$3,2)>=6',
                                     'format': pink_weekends_form})


# Nosaka fitteru kolonnas garāko uzvārdu lai pēc tam to kolonnas platumam.
fitt_name_str_max = group_it['Fitter'].str.len().max()

# Šeit kombinēts: kolonnas platums noteikts dodot  funkciju, kas atsaucas uz "B:B", un mainīgo kas nosaka garāko fit name str.
xrt_ws.set_column(xl_col_to_name(1)+":"+xl_col_to_name(1), fitt_name_str_max)
xrt_ws.set_column("A:"+xl_col_to_name(len(group_it.columns)),iso_netw_switch_cwidth)
xrt_ws.freeze_panes(3, 2 + len(fCountrRange))

# Nodefinēts stundu kopsavilkuma/knotroles apgabals!
spent_182_print_area = xl_range(3, 2, (len(fitterCount) + 2), len(fCountrRange) + 1)

# Noņem nulles no dienu skaita kontroles apgabala.
remove_zeroes = xrt_wb.add_format({'num_format': '#'})
xrt_ws.conditional_format(spent_182_print_area, {'type': 'cell',
                                                 'criteria':'equal to',
                                                 'value': 0,
                                                 'format':remove_zeroes})
# Iekrāsot dienu skaitu virs 151d.
paint_over_151ds = xrt_wb.add_format({'num_format': '#',
                                     # 'bg_color': '#fce4d6',
                                     'font_color': '#FF0000',
                                      'bold': True})
xrt_ws.conditional_format(spent_182_print_area, {'type': 'cell',
                                                 'criteria':'>',
                                                 'value': 151,
                                                 'format':paint_over_151ds})
# Filtrs uz Fitters, A-Z listes atskaitē.
xrt_ws.autofilter('B3:B3')

format_tchart_header_timestamp = xrt_wb.add_format({'num_format': '#',
                                     'bg_color': '#fce4d6',
                                     'font_color': '#FF0000'})

format_tchart_header_timestamp = xrt_wb.add_format()
format_tchart_header_timestamp.set_font_color('#FF0000')
format_tchart_header_timestamp.set_italic()
tchart_created = f"Time Chart prepared on: {dt.now().date().strftime('%d.%m.%y')} (at: {dt.now().strftime('%H:%M')})"
xrt_ws.write('B1', tchart_created, format_tchart_header_timestamp)


# Saglabā XlsxWriter.
#  writer.save()  #   ar šo neatbrīvoja temp failu, nevarēja izdzēst ar os.unlink(file), izdevās ar writer.close()...
writer.close()

default_cutoff_date = tagad - datetime.timedelta(3*365/12)
xw.Range('choose_date').value = default_cutoff_date


targ_wb = xw.Book(temp_wb_fpath.format(home))
targ_wb.api.ConvertComments() # 23.01.22.: pārkonvertē no notes uz threaded comments
targ_wb.save()