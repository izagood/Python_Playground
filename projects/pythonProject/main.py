import pandas as pd
import glob


def test():

    all_data = pd.DataFrame()

    date = '2021-03'
    title = '건기원_공기질측정데이터_' + date

    writer = pd.ExcelWriter('C:/testPyExcel/wid/' + title + '.xlsx', engine='xlsxwriter')
    for f in glob.glob('C:/uploadExcel/' + date + '_*.xlsx'):
        all_data = pd.read_excel(f)
        file_name = f[23:40]
        print(file_name)
        all_data.to_excel(writer, sheet_name=file_name, header=True, index=False)
        worksheet = writer.sheets[file_name]  # pull worksheet object
        for idx, col in enumerate(all_data):  # loop through all columns
            series = all_data[col]
            max_len = max((
                series.astype(str).map(len).max(),  # len of largest item
                len(str(series.name))  # len of column name/header
            )) + 1  # adding a little extra space
            worksheet.set_column(idx, idx, max_len)  # set column width

    print(all_data.shape)
    writer.save()

    print('끝')


# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    test()
