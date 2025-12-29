import pandas as pd
import numpy as np
import glob
import sys
import os
import re
import json
from functools import wraps
import lark_oapi as lark
from lark_oapi.api.bitable.v1 import *

class FeishuLogger(object):
    def __init__(self):
        pass

    def __call__(self, func):
        @wraps(func)
        def wrapped_function(*args, **kwargs):
            response = func(*args, **kwargs)
            # 处理失败返回
            if not response.success():
                lark.logger.error(
                    f"client.bitable.v1.app_table_record.search failed, code: {response.code}, "
                    f"msg: {response.msg}, "
                    f"log_id: {response.get_log_id()}, "
                    f"resp: \n{json.dumps(json.loads(response.raw.content), indent=4, ensure_ascii=False)}")
                return
            # 处理业务结果
            lark.logger.info(lark.JSON.marshal(response.data, indent=4))
            return response
        return wrapped_function


my_log = FeishuLogger()


class FeishuTable:
    def __init__(self, app_id, app_secret, app_token, table_id):
        self.my_log = FeishuLogger()
        self.app_token = app_token
        self.table_id = table_id
        self.client = lark.Client.builder() \
            .app_id(app_id) \
            .app_secret(app_secret) \
            .log_level(lark.LogLevel.DEBUG) \
            .build()

    @my_log
    def find_record(self, fields):
        # 构造请求对象
        condition_list = [Condition.builder().field_name(k).operator("is").value([v]).build()
                          for k, v in fields.items()]
        request: SearchAppTableRecordRequest = SearchAppTableRecordRequest.builder() \
            .app_token(self.app_token) \
            .table_id(self.table_id) \
            .page_size(20) \
            .request_body(SearchAppTableRecordRequestBody.builder()
                          .filter(FilterInfo.builder()
                                  .conjunction("and")
                                  .conditions(condition_list)
                                  .build())
                          .automatic_fields(False)
                          .build()) \
            .build()
        # 发起请求
        response: SearchAppTableRecordResponse = self.client.bitable.v1.app_table_record.search(request)
        return response

    @my_log
    def add_record(self, fields):
        # 构造请求对象
        request: CreateAppTableRecordRequest = CreateAppTableRecordRequest.builder() \
            .app_token(self.app_token) \
            .table_id(self.table_id) \
            .request_body(AppTableRecord.builder()
                          .fields(fields)
                          .build()) \
            .build()
        # 发起请求
        response: CreateAppTableRecordResponse = self.client.bitable.v1.app_table_record.create(request)
        return response

    @my_log
    def update_record(self, record_id, fields):
        # 构造请求对象
        request: UpdateAppTableRecordRequest = UpdateAppTableRecordRequest.builder() \
            .app_token(self.app_token) \
            .table_id(self.table_id) \
            .record_id(record_id) \
            .request_body(AppTableRecord.builder()
                          .fields(fields)
                          .build()) \
            .build()

        # 发起请求
        response: UpdateAppTableRecordResponse = self.client.bitable.v1.app_table_record.update(request)
        return response


# =============== 结果统计 =============== #
def record(path):
    # 1. 整理表头信息
    head_msg = [p for p in os.listdir(path) if p.startswith("headMessage")][0]
    head = pd.read_excel(fr'{path}/{head_msg}', header=None, index_col=0)
    proj = head.loc["项目类型", 1]
    if "项目类型" in pd.ExcelFile(fr'{path}/{head_msg}').sheet_names:
        info = pd.read_excel(fr'{path}/{head_msg}', sheet_name="项目类型")
        info.fillna("", inplace=True)
        grp_type = info.query(f"`项目类型` == '{proj}'")["合并方式[GRP_TYPE]"].iloc[0]

    elif head.loc["项目类型", 1] == "GM300":
        grp_type = "GCMS|LCMS"
    elif head.loc["项目类型", 1] == "AQ700":
        grp_type = "RP|HILIC"
    else:
        grp_type = ""

    
    
    # 2. 整理进样信息
    df_info = pd.read_excel(fr'{path}/sampleInfo.xlsx')

    info_blank = df_info[df_info['sample_type'] == 'BLK']
    info_qc = df_info[df_info['sample_type'] == 'QC']
    info_sample = df_info[df_info['sample_type'] == 'SPL']
    samples = info_sample['sample_name'].tolist()

    # 3. 整理结果数据
    # result_path = glob.glob(os.path.join(sys.argv[2], f"{head.loc['项目编号', 1]}*.xlsx"))[0]
    # d_quant = fr'{path}/results/Quantification-Spl.xlsx'
    # sheet = "merge" if "merge" in pd.ExcelFile(d_quant).sheet_names else 0
    # df = pd.read_excel(d_quant, sheet_name=sheet)

    quant_list = []
    for grp in grp_type.split('|'):
        path_g = re.sub(pattern=grp_type, repl=grp, string=path)
        df_g = pd.read_excel(fr'{path_g}/results/Quantification-Spl.xlsx')
        quant_list.append(df_g)

    df = pd.concat(quant_list)


    df_result = df[samples].copy()
    df_result.replace(0, np.nan, inplace=True)
    nums = df_result.count()

    df_result['nan'] = df_result.count(axis=1)
    df_full = df_result[df_result['nan'] == len(samples)]

    data = {
        '服务单号': head.loc['项目编号', 1],
        '产品名称': head.loc["项目类型", 1],
        '客户姓名': head.loc['客户姓名', 1],
        '客户单位': head.loc['客户单位', 1],
        '样本': head.loc['项目名称', 2],
        '样本类型': head.loc['样本类型', 1],
        '样本数量': len(info_sample),
        '空白数量': len(info_blank),
        'QC数量': len(info_qc),
        '检出总量': len(df),
        '最高检出量': float(nums.max()),
        '最低检出量': float(nums.min()),
        '检出量平均值': float(round(nums.mean(), 2)),
        '检出量中位值': float(round(nums.median(), 2)),
        '缺失值比例': float(round(df_result.isna().sum().sum() / df_result.size, 4)),
        '无缺失行比例': float(round(len(df_full) / len(df_result), 4)),
    }
    # 4. QC相关数据整理


    rsd_list = []
    cor_list = []

    for grp in grp_type.split('|'):
        path_g = re.sub(pattern=grp_type, repl=grp, string=path)
        if os.path.exists(fr'{path_g}/results/Quantification-Cor.xlsx'):
            df_cor = pd.read_excel(fr'{path_g}/results/Quantification-Cor.xlsx')
            if df_cor.size == 0:
                cor_list.append(0)
            else:
                cor_list.extend(df_cor.to_numpy().flatten())
            df_rsd = pd.read_excel(fr'{path_g}/results/Quantification-IS.xlsx', sheet_name='Statistics')
            if df_rsd.query("~`RSD[%]`.isna()").size:
                rsd_list.extend(df_rsd['RSD[%]'].to_list())
            else:
                rsd_list.append(0)

    data.update({
        'Corr(min)': float(round(np.min(cor_list), 4)),
        'IS_RSD(med)': float(round(np.median(rsd_list)/100, 4)),
    })
    return data


if __name__ == '__main__':
    print('开始统计项目信息')
    # path = r"Y:\gaoba\CCM\demo\20250410\HILIC\TEST-SAMPLE"
    # proj = "氨基酸"
    info = record(sys.argv[1])
    my_table = FeishuTable("cli_a8ae9609017f1013", "oulOf8a5U43wS82zWAEuggfzdo52lSfI",
                           "UAcQbl5o1aAncNsudKfc7GqenSb", 'tblbPssMHM7DZOhP')
    my_record = my_table.find_record({'服务单号': info['服务单号']})
    if my_record.data.items:
        record_id = my_record.data.items[0].record_id
        my_table.update_record(record_id, info)
    else:
        my_table.add_record(info)

