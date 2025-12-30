"""
需传入3和参数，分别是：
    1. 项目路径
    2. 结果路径
    3. pdf路径
"""
import pandas as pd
import sys

path = sys.argv[1]


# =============== 1. 整理样本信息 =============== #
def col_info():
    data = [['样本编号', '样本上机时的编号'],
            ['取样量', '从原始样本中取用的量，固体样本单位为 mg，液体样本单位为 μL']]
    return pd.DataFrame(data, columns=['列名', '说明'])


def sample_data():
    df = pd.read_excel(fr'{path}/sampleInfo.xlsx')
    data = df[df['sample_type'] == 'SPL'][['sample_name', '称样量(mg)']]
    data.columns = ['进样编号', '取样量']
    return data


def save_sheet(writer, df, sheet_name):
    df.to_excel(writer, sheet_name=sheet_name, index=False)
    workbook = writer.book
    worksheet = writer.sheets[sheet_name]

    # -------- 1. 标题行设置 -------- #
    title_format = workbook.add_format({
        'font_name': 'Arial',
        'font_size': 10,
        'font_color': '#FFFFFF',
        'bold': True,
        'bg_color': '#008CCE',
        'align': 'center',
        'valign': 'vcenter'})
    for col_num, value in enumerate(df.columns.values):
        worksheet.write(0, col_num, value, title_format)

    # -------- 2. 内容格式设置 -------- #
    global_format = workbook.add_format({
        'font_name': 'Arial',
        'font_size': 8,
        'font_color': '#000000',
        'valign': 'vcenter'})
    for row_num in range(1, len(df) + 1):
        worksheet.set_row(row_num, 16, global_format)

    # -------- 3. 背景色 -------- #
    worksheet.conditional_format(1, 0, len(df), len(df.columns) - 1, {
        'type': 'formula',
        'criteria': '=MOD(ROW(),2)=0',
        'format': workbook.add_format({
            'bg_color': '#B7DEE8',
            'border': 1,
            'border_color': '#92CDDC'})})
    worksheet.conditional_format(1, 0, len(df), len(df.columns) - 1, {
        'type': 'formula',
        'criteria': '=MOD(ROW(),2)=1',
        'format': workbook.add_format({
            'bg_color': '#DAEEF3',
            'border': 1,
            'border_color': '#92CDDC'})})

    # -------- 4. 行高列宽 -------- #
    worksheet.set_row(row=0, height=24)
    if sheet_name == '说明':
        worksheet.set_column('A:A', 15)
        worksheet.set_column('B:B', 60)
    elif sheet_name == 'Data':
        for col_num in range(len(df.columns)):
            worksheet.set_column(col_num, col_num, 10)


def sample_info():
    save_path = fr'{sys.argv[2]}/样品取用量.xlsx'
    with pd.ExcelWriter(save_path, engine='xlsxwriter') as writer:
        save_sheet(writer, sample_data(), sheet_name='Data')
        save_sheet(writer, col_info(), sheet_name='说明')



if __name__ == '__main__':
    print('\n开始生成样本信息')
    sample_info()
