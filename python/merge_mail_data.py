from functools import wraps
import datetime
import os
import io
import glob
import openpyxl
import pandas as pd

mail_folder = '//192.168.0.21/CustomerService/CSD/Incoming Mail Processing'
dir_excel = os.path.join(mail_folder, 'Mail Opening', '{}', '{}')
dir_export = os.path.join(mail_folder, 'Export')
pattern = dir_excel.format('**', '*.xl*')
files = [f for f in glob.glob(pattern, recursive=True)
         if not os.path.basename(f).startswith('~')]
print('>>> {} Excel files found'.format(len(files)))

sheet_name = 'SUMMARY'
range_address = 'B3:S33'
data_header = ('Date', 'No', 'Total', 'Recorder',
               'NoDCA', 'NoDCC', 'NoMCA', 'NoMCC', 'NoList', 'NoOther',
               'ValDCA', 'ValDCC', 'ValMCA', 'ValMCC', 'ValList', 'ValOther',
               'ValCash', 'ValCashPending')

def dataframe(header=None):
  def decorator(f):
    @wraps(f)
    def decorated_function(*args, **kwargs):
        # Run the wrapped function, which should return a range
        rng = f(*args, **kwargs)
        df = pd.DataFrame(rng)
        if header:
          df.columns = header
        # Render the template
        return df
    return decorated_function
  return decorator


@dataframe(data_header)
def read_excel_data(file=None, sheet_name=0, range_address='A1:B2'):
  rows = ()
  if file:
    with open(file,'rb') as f:
      in_mem_file = io.BytesIO(f.read())
      sheet = openpyxl.load_workbook(
          in_mem_file, read_only=True, data_only=True)[sheet_name]
      rng = sheet[range_address]
      rows = ((cell.value for cell in row) for row in rng if isinstance(row[0].value, datetime.datetime))
  return rows


f = files[28]
print('{}\n{}'.format(f, read_excel_data(f, sheet_name, range_address)))

# mgdf = pd.concat([read_excel_data(f, sheet_name, range_address)
#                   for f in files], axis=0, ignore_index=True)
# mgdf.to_csv(
#     os.path.join(dir_export, 'Merged.csv'), encoding='utf-8-sig')
# print('>>> Merged CSV File Saved')
