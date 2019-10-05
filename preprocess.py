import pandas as pd
from PIL import ImageGrab, Image
import win32com.client as win32
import matplotlib
matplotlib.rcParams['font.sans-serif'] = ['Microsoft JhengHei', 'sans-serif']
import matplotlib.pyplot as plt
import os

# %% Set file dependency
ABS_ROOT = os.path.abspath('')
DATA_FOLDER = "data/"
SUMMARY_SOURCE = "103年台南市污水下水道管線更新工程_B表.xls"

# %% Load data summary
summary = pd.read_excel(DATA_FOLDER+SUMMARY_SOURCE, sheet_name="TV檢視-異常統計表", header=[4, 5])
summary["編號"] = summary["編號"].fillna(method="ffill").astype("int32")

# %% Peek summary
summary.head(1)

# %% Plot data histogram
plt.figure(figsize=(15, 10), dpi=200)
plt.subplot(311)
summary[('管材', 'Unnamed: 9_level_1')].value_counts().plot(kind='bar')
plt.title('管材')

plt.subplot(312)
summary[('異常', '方位')].value_counts().plot(kind='bar')
plt.title('異常方位')

plt.subplot(313)
summary[('異常狀況', '說 明')].value_counts().plot(kind='bar')
plt.title('異常狀況')

# %%
summary = summary.groupby(('編號', 'Unnamed: 0_level_1'))

# %% Extract image from .xls file
excel = win32.gencache.EnsureDispatch('Excel.Application')
excel.Visible = False

for i in range(4, 7):
    # Obtain file name
    group = summary.get_group(i)
    pos_start = group.iloc[0][('起始', '人孔')]
    k = group.iloc[0][('管段編號', '上游人孔')]
    pos_end = k if k != pos_start else group.iloc[0][('管段編號', '下游人孔')]
    file = str(i).zfill(2) + "({}-{}).xls".format(pos_start, pos_end)

    # Extract image
    workbook = excel.Workbooks.Open(os.path.join(ABS_ROOT, DATA_FOLDER, file))

    index = 0
    for sheet in workbook.Worksheets:
        for i, shape in enumerate(sheet.Shapes):
            if shape.Name.startswith('Picture'):
                shape.Copy()
                image = ImageGrab.grabclipboard()
                image = image.resize((350, 250),Image.ANTIALIAS)
                plt.figure()
                plt.imshow(image)
                plt.title('{}-{}-{}.jpg'.format(
                    file.strip('.xls'),
                    group.iloc[index][('異常', '起點')],
                    group.iloc[index][('異常狀況', '說 明')]
                ))
                index = index + 1
                #image.save('{}.jpg'.format(i+1), 'jpeg')

    workbook.Close()

excel.Application.Quit()
