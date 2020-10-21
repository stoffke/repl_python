import pandas as pd
from pathlib import Path

#define parameters
#path to files
path_old = Path(r'https://repl.it/@mandel99/PlushLavishYottabyte#old.xlsx')
path_new = Path(r'https://repl.it/@mandel99/PlushLavishYottabyte#new.xlsx')
#list of key column(s)
key = ['id']
#sheets to read in
sheet = 'Sheet'

# Read in the two excel files and fill NA
old = pd.read_excel(path_old).fillna(0)
new = pd.read_excel(path_new).fillna(0)
#set index
old = old.set_index(key)
new = new.set_index(key)

#identify dropped rows and added (new) rows
dropped_rows = set(old.index) - set(new.index)
added_rows = set(new.index) - set(old.index)

#combine data
df_all_changes = pd.concat([old, new],
                           axis='columns',
                           keys=['old', 'new'],
                           join='inner')


#prepare functio for comparing old values and new values
def report_diff(x):
    return x[0] if x[0] == x[1] else '{} ---> {}'.format(*x)


#swap column indexes
df_all_changes = df_all_changes.swaplevel(axis='columns')[new.columns[0:]]

#apply the report_diff function
df_changed = df_all_changes.groupby(
    level=0, axis=1).apply(lambda frame: frame.apply(report_diff, axis=1))

#create a list of text columns (int columns do not have '{} ---> {}')
df_changed_text_columns = df_changed.select_dtypes(include='object')

#create 3 datasets:
#diff - contains the differences
#dropped - contains the dropped rows
#added - contains the added rows
diff = df_changed_text_columns[df_changed_text_columns.apply(
    lambda x: x.str.contains("--->") == True, axis=1)]
dropped = old.loc[dropped_rows]
added = new.loc[added_rows]

#create a name for the output excel file
fname = '{} vs {}.xlsx'.format(path_old.stem, path_new.stem)

#write dataframe to excel
writer = pd.ExcelWriter(fname, engine='xlsxwriter')
diff.to_excel(writer, sheet_name='diff', index=True)
dropped.to_excel(writer, sheet_name='dropped', index=True)
added.to_excel(writer, sheet_name='added', index=True)

#get xlswriter objects
workbook = writer.book
worksheet = writer.sheets['diff']
worksheet.hide_gridlines(2)
worksheet.set_default_row(15)

#get number of rows of the df diff
row_count_str = str(len(diff.index) + 1)

#define and apply formats
highligt_fmt = workbook.add_format({
    'font_color': '#FF0000',
    'bg_color': '#B1B3B3'
})
worksheet.conditional_format(
    'A1:ZZ' + row_count_str, {
        'type': 'text',
        'criteria': 'containing',
        'value': '--->',
        'format': highligt_fmt
    })

#save the output
writer.save()
print('\nDone.\n')
