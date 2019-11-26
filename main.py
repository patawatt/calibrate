import pandas as pd
import numpy as np

def get_order(id):
    order = {"CONTROL": 1,
             "CC1": 2,
             "CC2": 3,
             "DB1": 4,
             "DB2": 5}
    return order.get(id, "Invalid id: " + str(id))

def conc_to_area(tup, num=50):

    return tup[0]*num + tup[1]


def get_col_widths(dataframe):
    # FROM: https://stackoverflow.com/questions/
    # 29463274/simulate-autofit-column-in-xslxwriter
    # First we find the maximum length of the index column
    idx_max = max([len(str(s)) for s in dataframe.index.values] +
                  [len(str(dataframe.index.name))])
    # Then, we concatenate this to the max of the lengths of column name and
    # its values for each column, left to right
    return [idx_max] + [max([len(str(s)) for s in dataframe[col].values] +
                            [len(col)]) for col in dataframe.columns]


def autofit_col(dataframe, worksheet):
    for i, width in enumerate(get_col_widths(dataframe)):
        # Modified from: https://stackoverflow.com/questions/29463274/
        # simulate-autofit-column-in-xslxwriter
        worksheet.set_column(i, i, width)

# old data

aug = pd.read_excel('btex_august2019.xls', dtype={'file#': str, 'Sample#': str})
aug['file#'].iloc[31:46] = aug['file#'].iloc[31:46].str[0:3] + \
                           aug['file#'].iloc[31:46].str[4:]
aug['date analysed'] = pd.to_datetime(aug['file#'].str[0:5], format='%m-%d') +\
                       pd.offsets.DateOffset(years=119)
aug.drop('file#', axis=1, inplace=True)
aug.rename(columns={'Sample#':'Sample Name','(m+p)xylene':'(m+p)-xylene'},
           inplace=True)


# new data
df = pd.read_csv('nov25_final.csv')
df.columns = df.iloc[0, :]
df = df.iloc[11:, 1:]
df = df.reset_index(drop=True)
df.rename(columns={'Date Acquired': 'date analysed'}, inplace=True)

file_info = df.iloc[:, :5]
benzene = pd.to_numeric(df.iloc[:, 9])
toluene = pd.to_numeric(df.iloc[:, 15])
e_benzene = pd.to_numeric(df.iloc[:, 21])
mp_xylene = pd.to_numeric(df.iloc[:, 27])
o_xylene = pd.to_numeric(df.iloc[:, 33])

# boundary between the two calibration curves
boundary = 50

# benzene slope and intercept, a = low calibration, b = high calibration
f1a = (94710.0, 0.0)
f1b = (102900, -1389000)
# toluene slope and intercept, a = low calibration, b = high calibration
f2a = (102900.0, 0.0)
f2b = (108000.0, -1684000.0)
# ethyl benzene slope and intercept, a = low calibration, b = high calibration
f3a = (125600.0, 0.0)
f3b = (137400.0, -1940000.0)
# (m+p)-xylene slope and intercept, a = low calibration, b = high calibration
f4a = (249400.0, 0.0)
f4b = (259400.0, -2969000.0)
# o-xylene slope and intercept, a = low calibration, b = high calibration
f5a = (129000, 0.0)
f5b = (136800.0, -1629000.0)

for i in range(len(benzene)):
    if benzene[i] < conc_to_area(f1b):
        benzene[i] = (benzene[i]-f1a[1])/f1a[0]
    else:
        benzene[i] = (benzene[i]-f1b[1])/f1b[0]

for i in range(len(toluene)):
    if toluene[i] < conc_to_area(f2b):
        toluene[i] = (toluene[i] - f2a[1]) / f2a[0]
    else:
        toluene[i] = (toluene[i] - f2b[1]) / f2b[0]

for i in range(len(e_benzene)):
    if e_benzene[i] < conc_to_area(f3b):
        e_benzene[i] = (e_benzene[i] - f3a[1]) / f3a[0]
    else:
        e_benzene[i] = (e_benzene[i] - f3b[1]) / f3b[0]

for i in range(len(mp_xylene)):
    if mp_xylene[i] < conc_to_area(f4b):
        mp_xylene[i] = (mp_xylene[i] - f4a[1]) / f4a[0]
    else:
        mp_xylene[i] = (mp_xylene[i] - f4b[1]) / f4b[0]

for i in range(len(o_xylene)):
    if o_xylene[i] < conc_to_area(f5b):
        o_xylene[i] = (o_xylene[i] - f5a[1]) / f5a[0]
    else:
        o_xylene[i] = (o_xylene[i] - f5b[1]) / f5b[0]

data = [df['Sample Name'], benzene, toluene, e_benzene, mp_xylene, o_xylene]
col_names = ['Sample name', 'benzene', 'toluene', 'ethyl-benzene', '(m+p)-xylene', 'o-xylene']
final = pd.concat(data, axis=1, keys=col_names)

benzene.name = 'benzene'
toluene.name = 'toluene'
e_benzene.name = 'ethyl-benzene'
mp_xylene.name = '(m+p)-xylene'
o_xylene.name = 'o-xylene'

data_compare_cals = [df['Sample Name'], benzene, df['(# 1) Amount'], toluene, df['(# 2) Amount'], e_benzene, df['(# 3) Amount'], mp_xylene, df['(# 4) Amount'], o_xylene, df['(# 5) Amount'], pd.to_datetime(df['date analysed'])]
compare = pd.concat(data_compare_cals, axis=1)
final_cal = compare.drop(['(# 1) Amount', '(# 2) Amount', '(# 3) Amount', '(# 4) Amount','(# 5) Amount'], axis=1)
# join old and new data
final_cal = pd.concat([aug, final_cal], join='outer')
# change all zeros to 'nd' for non-detect (will replace with detection limit.)
final_cal.replace(0, 'nd', inplace=True)
final_cal['Sample Name'] = final_cal['Sample Name'].apply(lambda x: x.upper())

# remove duplicates
no_duplicates = final_cal.drop_duplicates(subset='Sample Name', keep='last')

# divide sample id components in separate columns for easy querying
samples = no_duplicates[no_duplicates['Sample Name'].str.contains('DAY')]
samples['day'] = samples['Sample Name'].str.extract('(\d+)')
samples['id'] = samples['Sample Name'].str.extract\
    ('(DB1|DB2|CC1|CC2|CONTROL|RAIN_BARREL|DI_TAP)')
samples['rep'] = samples['Sample Name'].str.\
    extract('(8.4XDIL|1/2|2/2|POSITION1|POSITION2)')
samples['rep'].fillna('1/2', inplace=True)
samples['date sampled'] = pd.to_datetime(samples['day'].astype(int), unit='D',
                                         origin=pd.Timestamp('2019-08-02')) + \
                          pd.Timedelta(hours=11)

# rearrange column order to make more sense
cols = samples.columns.tolist()
samples = samples[cols[:6] + cols[7:10] + [cols[6]] + cols[10:12]]

# remove spiked sample used to check retention times
samples = samples[~samples['Sample Name'].str.
    contains('SPIKED|RAIN|DI_TAP|DI-TAP')]

# sorting into easy to read order
samples['day'] = samples['day'].astype(int)
# vectorize function so it works on DataFrame
vget_order = np.vectorize(get_order)
samples['order'] = vget_order(samples['id'])
samples.sort_values(by=['day','order','rep'], inplace=True)
samples.reset_index(inplace=True)
samples.drop(['index', 'order'], axis=1, inplace=True)
samples['date analysed'] = samples['date analysed'].dt.date
samples['date sampled'] = samples['date sampled'].dt.date
samples['days before analysis'] = (samples['date analysed'] -
                                   samples['date sampled']).dt.days

samples['Sample Name'].replace(regex=True, inplace=True,
                               to_replace='_BTX', value='')

# make new DataFrame for tox experiments
tox = no_duplicates[no_duplicates['Sample Name'].str.contains('TOX')]

tox['id'] = tox['Sample Name'].str.extract\
    ('(DB1|DB2|CC1|CC2|CONTROL)')
tox['dilution'] = tox['Sample Name'].str.extract\
    ('(12|25|50|100)')
tox['dilution'] = tox['dilution'].fillna(0).astype(int)
tox['experiment'] = tox['Sample Name'].str.extract\
    ('(TOX1|TOX2|TOX3|TOX4|TOX5|TOX6|TOX7)')
tox['order'] = vget_order(tox['id'])
tox.loc[tox['Sample Name'].str.contains('TOX5B'), 'experiment'] = 'TOX5B'

tox.sort_values(by=['experiment', 'order', 'dilution'], inplace=True)
tox.drop(['order'], axis=1, inplace=True)


# make new DataFrames for position test, blanks and standards.
positions = final_cal[final_cal['Sample Name'].str.contains('POSITION1|POSITION2', regex=True)]
field_blks = final_cal[final_cal['Sample Name'].str.contains('RAIN|DI', regex=True)]
analytical_blks = final_cal[final_cal['Sample Name'].str.contains('BLK')]
stds = final_cal[final_cal['Sample Name'].str.contains('PPB')]



# Create a Pandas Excel writer using XlsxWriter as the engine.
excel_file = 'dilbit_btex.xlsx'

df_sheet_pairs = {'column samples': samples,'tox experiments': tox,
                  'position test': positions, 'field blanks': field_blks,
                  'analytical blanks': analytical_blks, 'standards': stds,
                  'raw_data': df}

writer = pd.ExcelWriter(excel_file, engine='xlsxwriter')

# Access the XlsxWriter workbook and worksheet objects from the dataframe.
workbook = writer.book

for sheet_name, data_frame in df_sheet_pairs.items():
    data_frame.to_excel(writer, sheet_name=sheet_name, startrow=1, startcol=1,
                        index=False)
    worksheet = writer.sheets[sheet_name]
    autofit_col(data_frame, worksheet)


# Close the Pandas Excel writer and output the Excel file.
writer.save()

