# 创建类
import pandas as np
from os import startfile
import urllib3
import json
from urllib3.util.url import Url
import pandas as pd
urllib3.disable_warnings()

class downloadexport():
    def __init__(self,url,hs_code,star_year,star_month,end_year,end_month):
        self.url = url
        self.hs_code = hs_code
        self.star_year = star_year
        self.star_month = star_month
        self.end_year = end_year
        self.end_month = end_month
    # 定义参数列表
    def firefox_headers(self):
        hs_code = self.hs_code 
        star_year = self.star_year 
        star_month = self.star_month
        end_year = self.end_year 
        end_month = self.end_month
        a = []
        if star_year==end_year:
            for i in range(star_year,end_year+1):
                for j in range(star_month,end_month+1):
                    b =[{
                        "year": i,
                        "month": j,
                        "hs_code":hs_code
                    }]
                    a = a + b
        else:
            for i in range(star_year,end_year+1):
                if(i == star_year):
                    for j in range(star_month,13):
                        b =[{
                            "year": i,
                            "month": j,
                            "hs_code":hs_code
                        }]
                        a = a + b
                else:
                    if(i == end_year):
                        for j in range(1,end_month+1):
                            b=[{
                                "year": i,
                                "month": j,
                                "hs_code":hs_code
                            }]
                            a = a + b
                    else:
                        for j in range(1,13):
                            b=[{
                                "year": i,
                                "month": j,
                                "hs_code":hs_code
                            }]
                            a = a + b
        return a
    def down_data(self):
        # 参数
        hs_code = self.hs_code
        # 输入日期
        da = self.firefox_headers()
        amn = urllib3.PoolManager()
        data = pd.DataFrame()
        for i in range(len(da)):
            res = amn.request("GET",url = self.url,fields=da[i])
            data1 = json.loads(res.data)
            data2 = pd.DataFrame(data1)
            data2["year_month"] = data2["year"].map(str)+"/"+data2["month"].map(str)
            data2 = data2.drop(["year","month"],axis=1).set_index(["year_month"])
            data = data.append(data2)
        return data
    def to_excel(self):
        # 写进excel
        file_name = self.hs_code+"_export_thailand.xlsx"
        writer = pd.ExcelWriter(file_name)
        self.down_data().to_excel(writer)
        writer.save()
        writer.close()

url = "https://dataapi.moc.go.th/export-harmonize-countries"
