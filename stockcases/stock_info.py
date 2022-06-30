import json
import os
import string

import math
import openpyxl
from common.request_util import RequestUtil
from log.log_util import LogUtil
import pandas as pd


class StockInfo:
    ratesdebt_index = ['有息负债', '总资产', '有息负债率','总负债','固定资产', '在建工程','存货','商誉']
    cash_index = ['经营现金流', '营业收入', '现收比', '经营现金流净额', '净利润', '现净比', '投资现金流净额', '资本开支', '融资现金流净额']
    bussiness_index = ['营业收入', '净利润', '扣非净利润', '营收同比', '净利润同比', '扣非净利润同比', '销售费用', '管理费用', '研发费用', '财务费用', '销售费用同比', '管理费用同比', '研发费用同比', '财务费用同比', '毛利率', '扣非净利率','费用率','净资产收益率']
    assets_index =['固定无形资产占比','在建工程占比','应收账款收入比','存货成本比','应收账款周转率','固定无形资周转率','存货周转率','总资产周转率']
    zy_index = ['质押人','质押股数','质押市值','质押比例','占总股份比例','预警线','平仓线','质押机构','质押原因','质押目的','质押开始日','质押结束日']
    @classmethod
    def getzyzb(cls, codes):
        method = "get"
        zyzb_url = "http://emweb.securities.eastmoney.com/PC_HSF10/NewFinanceAnalysis/ZYZBAjaxNew"
        data = {
            "type": 1,
            "code": codes
        }
        log = LogUtil().log_free()
        rep = RequestUtil().send_request(method, zyzb_url, data, "")
        if len(json.loads(rep)["data"]) > 0:
            log.info("获取主要指标数据成功")
            zyzb_info = json.loads(rep)["data"]
            return zyzb_info
        else:
            log.info("获取主要指标数据失败")
            return

    @classmethod
    def getzcfzb(cls,codes,dates):
        method = "get"
        zcfzb_url = "http://emweb.securities.eastmoney.com/PC_HSF10/NewFinanceAnalysis/zcfzbAjaxNew"
        data = {
            "companyType": 4,
            "reportDateType": 0,
            "reportType": 1,
            "dates": dates,
            "code": codes
        }
        log = LogUtil().log_free()
        rep = RequestUtil().send_request(method, zcfzb_url, data, "")
        if len(json.loads(rep)["data"]) > 0:
            log.info("获取资产负债表数据成功")
            zcfzb_info = json.loads(rep)["data"]
            return zcfzb_info
        else:
            log.info("获取资产负债表数据失败")
            return
    @classmethod
    def getxjllb(cls,codes,dates):
        method = "get"
        xjllb_url = "http://emweb.securities.eastmoney.com/PC_HSF10/NewFinanceAnalysis/xjllbAjaxNew"
        data = {
            "companyType": 4,
            "reportDateType": 0,
            "reportType": 1,
            "dates": dates,
            "code": codes
        }
        log = LogUtil().log_free()
        rep = RequestUtil().send_request(method,xjllb_url, data, "")
        if len(json.loads(rep)["data"]) > 0:
            log.info("获取现金流量表数据成功")
            xjllb_info = json.loads(rep)["data"]
            return xjllb_info
        else:
            log.info("获取现金流量表数据失败")
            return

    @classmethod
    def getlrb(cls, codes, dates):
        method = "get"
        lrb_url = "http://emweb.securities.eastmoney.com/PC_HSF10/NewFinanceAnalysis/lrbAjaxNew"
        data = {
            "companyType": 4,
            "reportDateType": 0,
            "reportType": 1,
            "dates": dates,
            "code": codes
        }
        log = LogUtil().log_free()
        rep = RequestUtil().send_request(method, lrb_url, data, "")
        if len(json.loads(rep)["data"]) > 0:
            log.info("获取利润表数据成功")
            lrb_info = json.loads(rep)["data"]
            return lrb_info
        else:
            log.info("获取利润表数据失败")
            return
    @classmethod
    def is_None(cls,num):
        if num == None:
            return 0
        else:
            return num

    @classmethod
    def Merge(cls, dict1, dict2, dict3, dict4):
        res = {**dict1, **dict2,**dict3,**dict4}
        return res

    @classmethod
    def getstock_info(cls,codes,dates):
        zcfzb = {}
        xjllb = {}
        lrb = {}
        zyzb = {}
        stock_info = {}
        zcfzb_info = StockInfo.getzcfzb(codes, dates)
        for i in range(len(zcfzb_info)):
            zcfzb[zcfzb_info[i]["REPORT_DATE"][0:4]] = zcfzb_info[i]
        xjllb_info = StockInfo.getxjllb(codes, dates)
        for j in range(len(xjllb_info)):
            xjllb[xjllb_info[j]["REPORT_DATE"][0:4]] = xjllb_info[j]
        lrb_info = StockInfo.getlrb(codes, dates)
        for k in range(len(lrb_info)):
            lrb[lrb_info[k]["REPORT_DATE"][0:4]] = lrb_info[k]
        year_list_length = len(list(zcfzb.keys()))
        zyzb_info = StockInfo.getzyzb(codes)[0:year_list_length]
        for l in range(len(zyzb_info)):
            zyzb[zyzb_info[l]["REPORT_DATE"][0:4]] = zyzb_info[l]
        for key in zcfzb.keys():
            stock_info[key] = StockInfo.Merge(zcfzb[key], xjllb[key], lrb[key],zyzb[key])
        return stock_info

    @classmethod
    def rates_debt(cls,stock_info):
        #获取有息负债资料
        ratesdebt = {'年份': StockInfo.ratesdebt_index}
        for key in stock_info.keys():
            rates_debt = round((StockInfo.is_None(stock_info[key]["NONCURRENT_LIAB_1YEAR"]) + StockInfo.is_None(stock_info[key]["LONG_LOAN"]) + StockInfo.is_None(stock_info[key]["SHORT_LOAN"]) + StockInfo.is_None(stock_info[key]["BOND_PAYABLE"]))/100000000,2)
            total_assets = round(stock_info[key]["TOTAL_ASSETS"]/100000000,2)
            if total_assets > 0:
                rates_debt_ratio = round(rates_debt/total_assets, 2)
            else:
                rates_debt_ratio = 0
            total_liability = round(stock_info[key]["TOTAL_LIABILITIES"] / 100000000, 2)
            fixed_asset = round(stock_info[key]["FIXED_ASSET"] / 100000000, 2)
            cip = round(StockInfo.is_None(stock_info[key]["CIP"]) / 100000000, 2)
            inventory = round(stock_info[key]["INVENTORY"] / 100000000, 2)
            good_will = round(StockInfo.is_None(stock_info[key]["GOODWILL"]) / 100000000, 2)
            ratesdebt[key] = [rates_debt,total_assets, rates_debt_ratio,total_liability,fixed_asset,cip,inventory,good_will]
        return ratesdebt

    @classmethod
    def cash_flow(cls, stock_info):
        # 获取现金流资料
        cashflow = {'年份': StockInfo.cash_index}
        for key in stock_info.keys():
            total_operate_inflow = round(stock_info[key]["TOTAL_OPERATE_INFLOW"] / 100000000, 2)
            total_operate_income = round(stock_info[key]["TOTAL_OPERATE_INCOME"] / 100000000, 2)
            inflow_income = round(total_operate_inflow / total_operate_income, 2)
            netcash_operate = round(stock_info[key]["NETCASH_OPERATE"] / 100000000, 2)
            netprofit = round(stock_info[key]["PARENT_NETPROFIT"] / 100000000, 2)
            netcash_netprofit = round(netcash_operate/netprofit,2)
            netcash_invest = round(stock_info[key]["NETCASH_INVEST"] / 100000000, 2)
            construct_long_asset = round(stock_info[key]["CONSTRUCT_LONG_ASSET"] / 100000000, 2)
            netcash_finance = round(stock_info[key]["NETCASH_FINANCE"] / 100000000, 2)
            cashflow[key] = [total_operate_inflow, total_operate_income, inflow_income,netcash_operate,netprofit,netcash_netprofit,netcash_invest,construct_long_asset,netcash_finance]
        return cashflow

    @classmethod
    def bussiness_income(cls, stock_info):
        # 获取经营情况资料
        bussinessincome = {'年份': StockInfo.bussiness_index}
        for key in stock_info.keys():
            total_operate_income = round(stock_info[key]["TOTAL_OPERATE_INCOME"] / 100000000, 2)
            netprofit = round(stock_info[key]["PARENT_NETPROFIT"] / 100000000, 2)
            deduct_parent_netprofit = round(stock_info[key]["DEDUCT_PARENT_NETPROFIT"] / 100000000, 2)
            total_operate_income_yoy = round(StockInfo.is_None(stock_info[key]["TOTAL_OPERATE_INCOME_YOY"]) / 100, 2)
            netprofit_YOY = round(StockInfo.is_None(stock_info[key]["PARENT_NETPROFIT_YOY"]) / 100, 2)
            deduct_parent_netprofit_yoy = round(StockInfo.is_None(stock_info[key]["DEDUCT_PARENT_NETPROFIT_YOY"]) / 100, 2)
            sale_expense = round(StockInfo.is_None(stock_info[key]["SALE_EXPENSE"]) / 100000000, 2)
            manage_expense = round(StockInfo.is_None(stock_info[key]["MANAGE_EXPENSE"]) / 100000000, 2)
            research_expense = round(StockInfo.is_None(stock_info[key]["RESEARCH_EXPENSE"]) / 100000000, 2)
            finance_expense = round(StockInfo.is_None(stock_info[key]["FINANCE_EXPENSE"]) / 100000000, 2)
            sale_expense_yoy = round(StockInfo.is_None(stock_info[key]["SALE_EXPENSE_YOY"]) / 100, 2)
            manage_expense_yoy = round(StockInfo.is_None(stock_info[key]["MANAGE_EXPENSE_YOY"]) / 100, 2)
            research_expense_yoy = round(StockInfo.is_None(stock_info[key]["RESEARCH_EXPENSE_YOY"]) / 100, 2)
            finance_expense_yoy = round(StockInfo.is_None(stock_info[key]["FINANCE_EXPENSE_YOY"]) / 100, 2)
            xsmll = round(StockInfo.is_None(stock_info[key]["XSMLL"]) / 100, 2)

            xsjll = round(StockInfo.is_None(stock_info[key]["DEDUCT_PARENT_NETPROFIT"]) / stock_info[key]["TOTAL_OPERATE_INCOME"], 2)
            fee_income = round((StockInfo.is_None(stock_info[key]["SALE_EXPENSE"])+StockInfo.is_None(stock_info[key]["MANAGE_EXPENSE"])+StockInfo.is_None(stock_info[key]["RESEARCH_EXPENSE"])+StockInfo.is_None(stock_info[key]["FINANCE_EXPENSE"])) / stock_info[key]["TOTAL_OPERATE_INCOME"], 2)
            roejq = round(StockInfo.is_None(stock_info[key]["ROEJQ"]) / 100, 2)
            bussinessincome[key] = [total_operate_income,netprofit,deduct_parent_netprofit,total_operate_income_yoy,netprofit_YOY,deduct_parent_netprofit_yoy,
                                    sale_expense,manage_expense,research_expense,finance_expense,sale_expense_yoy,manage_expense_yoy,research_expense_yoy,finance_expense_yoy,xsmll,xsjll,fee_income,roejq]
        return bussinessincome

    @classmethod
    def asset_structure(cls, stock_info):
        # 获取资产结构资料
        assetstructure = {'年份': StockInfo.assets_index}
        for key in stock_info.keys():
            #固定资产及无形资产占比
            fixasset_percent = round((stock_info[key]["FIXED_ASSET"] + StockInfo.is_None(stock_info[key]["INTANGIBLE_ASSET"])) /stock_info[key]["TOTAL_ASSETS"], 2)
            # 在建工程占比
            cip_percent = round(StockInfo.is_None(stock_info[key]["CIP"]) / stock_info[key]["TOTAL_ASSETS"], 2)
            # 应收账款收入比
            rece_income = round(StockInfo.is_None(stock_info[key]["ACCOUNTS_RECE"]) / stock_info[key]["TOTAL_OPERATE_INCOME"], 2)
            # 存货成本比
            ch_cost = round(StockInfo.is_None(stock_info[key]["INVENTORY"]) / stock_info[key]["OPERATE_COST"], 2)
            # 应收账款周转率
            yszkzzl = round(StockInfo.is_None(stock_info[key]["YSZKZZL"]) , 2)
            # 固定及无形资产周转率
            gudingwuxingzzl = round(StockInfo.is_None(stock_info[key]["TOTAL_OPERATE_INCOME"])/(stock_info[key]["FIXED_ASSET"] + StockInfo.is_None(stock_info[key]["INTANGIBLE_ASSET"])), 2)
            # 应收账款周转率
            chzzl = round(StockInfo.is_None(stock_info[key]["CHZZL"]), 2)
            # 应收账款周转率
            toazzl = round(StockInfo.is_None(stock_info[key]["TOAZZL"]), 2)
            assetstructure[key] = [fixasset_percent,cip_percent,rece_income,ch_cost,yszkzzl,gudingwuxingzzl,chzzl,toazzl]
        return assetstructure

    @classmethod
    def get_zyinfo(cls,codes):
        zy_url = "https://datacenter-web.eastmoney.com/api/data/v1/get?"
        method = "get"
        code = str(codes)[2:]
        data = {
            "reportName": "RPTA_APP_ACCUMDETAILS",
            "sortColumns": "PF_START_DATE",
            "columns": "ALL",
            "quoteColumns": "",
            "filter": "(TRADE_DATE>'2017-04-01')("+"SECURITY_CODE="+code+")"
        }
        log = LogUtil().log_free()
        rep = RequestUtil().send_request(method, zy_url, data, "")
        if json.loads(rep)["result"] == None:
            log.info("质押数据为空")
            return
        elif len(json.loads(rep)["result"]["data"]) > 0:
            log.info("获取质押数据成功")
            zyinfo = json.loads(rep)["result"]["data"]
            return zyinfo
        else:
            log.info("获取质押数据失败")
            return

    @classmethod
    def zy_info(cls,codes):
        zy_info = StockInfo.get_zyinfo(codes)
        zyinfo = []
        if zy_info == None:
            return
        else:
            for i in range(len(zy_info)):
                zyinfo_list = []
                if zy_info[i]["UNFREEZE_STATE"] == '未解押':
                    zyinfo_list.append(zy_info[i]["HOLDER_NAME"])
                    zyinfo_list.append(round(zy_info[i]["PF_NUM"]/100000000,2))
                    zyinfo_list.append(round(zy_info[i]["MARKET_CAP"] / 100000000,2))
                    zyinfo_list.append(round(zy_info[i]["PF_HOLD_RATIO"]/100,2))
                    zyinfo_list.append(round(zy_info[i]["PF_TSR"]/100,2))
                    zyinfo_list.append(zy_info[i]["WARNING_LINE"])
                    zyinfo_list.append(zy_info[i]["OPENLINE"])
                    zyinfo_list.append(zy_info[i]["PF_ORG"])
                    zyinfo_list.append(zy_info[i]["PF_REASON"])
                    zyinfo_list.append(zy_info[i]["PF_PURPOSE"])
                    zyinfo_list.append(str(zy_info[i]["PF_START_DATE"])[0:10])
                    zyinfo_list.append(str(zy_info[i]["UNFREEZE_DATE"])[0:10])
                else:
                    continue
                zyinfo.append(zyinfo_list)
        return zyinfo

    def get_stockdata(self,codes,stock_name,dates):
        stock_info = StockInfo.getstock_info(codes, dates)
        excel_file = stock_name + '.xlsx'
        if os.path.exists("E:\\stock_data\\" + excel_file):
            print("文件已存在！")
            return
        else:
            writer = pd.ExcelWriter("E:\\stock_data\\" + excel_file)
            workbook = writer.book
            fmt = workbook.add_format({"font_name": "微软雅黑", "bg_color": '#2A61A1', 'num_format': '0.00%', 'bold': True, 'font_color': 'white'})
            fmt1 = workbook.add_format({"font_name": "微软雅黑", "bg_color": '#FEFDC5', 'num_format': '#,##0.00', 'bold': True, 'font_color': 'black', 'border':1})
            fmt2 = workbook.add_format({"font_name": "微软雅黑", "bg_color": '#FEFDC5', 'num_format': '0.00%', 'bold': True, 'font_color': 'black','border':1})
            fmt3 = workbook.add_format(
                {"font_name": "微软雅黑", "bg_color": '#FEFDC5', 'num_format': '0.00%', 'bold': True, 'font_color': 'green','border': 1})
            fmt4 = workbook.add_format(
                {"font_name": "微软雅黑", "bg_color": '#CAFCFD', 'num_format': '#,##0.00', 'bold': True,'font_color': 'black',  'border': 1})
            fmt5 = workbook.add_format(
                {"font_name": "微软雅黑", "bg_color": '#CAFCFD', 'num_format': '0.00%', 'bold': True, 'font_color': 'black','border': 1})
            fmt.set_align('center')
            fmt1.set_align('center')
            fmt2.set_align('center')
            fmt3.set_align('center')
            fmt4.set_align('center')
            letter_list = list(string.ascii_uppercase)
            #资产负债图表信息
            ratesdebt = StockInfo.rates_debt(stock_info)
            year_list = sorted(list(ratesdebt.keys())[1:])
            table_length = len(year_list)+1
            ratesdebt_excle = pd.DataFrame(ratesdebt, index=StockInfo.ratesdebt_index, columns=year_list)
            sheet_ratesdebt = '资产负债'
            ratesdebt_excle.to_excel(writer, sheet_name=sheet_ratesdebt, encoding='utf8', startcol=0, startrow=1)
            worksheet1 = writer.sheets[sheet_ratesdebt]
            worksheet1.set_column('A:H', 15)
            worksheet1.write(letter_list[math.floor((table_length-1)/2)]+str(1),stock_name+"资产负债统计")
            worksheet1.write(letter_list[math.floor((table_length-1)/2)] + str(11), "有息负债率=有息负债/总资产")
            worksheet1.write(letter_list[table_length-1]+str(1), "单位：亿元")
            worksheet1.conditional_format('A1:'+ letter_list[table_length-1]+str(1),
                                          {'type': 'cell', 'criteria': '>=', 'value': 0, 'format': fmt})
            worksheet1.conditional_format('A2:'+ letter_list[table_length-1]+str(4),
                                          {'type': 'cell','criteria': '>=', 'value': 0,'format': fmt1})
            worksheet1.conditional_format('A2:' + letter_list[table_length - 1] + str(4),
                                          {'type': 'cell', 'criteria': '<', 'value': 0, 'format': fmt4})
            worksheet1.conditional_format('A5:'+ letter_list[table_length-1]+str(5),
                                          {'type': 'cell', 'criteria': '>=', 'value': 0, 'format': fmt2})
            worksheet1.conditional_format('A5:' + letter_list[table_length - 1] + str(5),
                                          {'type': 'cell', 'criteria': '<', 'value': 0, 'format': fmt5})
            worksheet1.conditional_format(
                'A6:' + letter_list[table_length - 1] + str(10),
                {'type': 'cell', 'criteria': '>=', 'value': 0, 'format': fmt1})
            worksheet1.conditional_format(
                'A6:' + letter_list[table_length - 1] + str(10),
                {'type': 'cell', 'criteria': '<', 'value': 0, 'format': fmt4})
            worksheet1.conditional_format('A11:' + letter_list[table_length - 1] + str(11), {'type': 'cell', 'criteria': '>=', 'value': 0, 'format': fmt3})
            chart_width = 110*table_length
            chart_height = 384
            chart_ratesdebt1 = workbook.add_chart({'type': 'column'})
            chart_ratesdebt1.set_size({'width': chart_width,'height': chart_height})
            chart_ratesdebt1.add_series({
                'categories': [sheet_ratesdebt, 1, 1, 1, table_length],
                'values': [sheet_ratesdebt, 2, 1, 2, table_length],
                'name': [sheet_ratesdebt, 2, 0],
                'line': {'color': 'white'}
            })
            chart_ratesdebt1.add_series({
                'name': [sheet_ratesdebt, 3, 0],
                'values': [sheet_ratesdebt, 3, 1, 3, table_length]
            })
            chart_ratesdebt1.add_series({
                'name': [sheet_ratesdebt, 5, 0],
                'values': [sheet_ratesdebt, 5, 1, 5, table_length]
            })
            chart_ratesdebt1.add_series({
                'name': [sheet_ratesdebt, 6, 0],
                'values': [sheet_ratesdebt, 6, 1, 6, table_length]
            })
            chart_ratesdebt1.add_series({
                'name': [sheet_ratesdebt, 7, 0],
                'values': [sheet_ratesdebt, 7, 1, 7, table_length]
            })
            chart_ratesdebt1.add_series({
                'name': [sheet_ratesdebt, 8, 0],
                'data_labels': {'value': True, 'num_format': '#,##0.00', 'position': 'outside_end'},
                'values': [sheet_ratesdebt, 8, 1, 8, table_length]
            })
            chart_ratesdebt1.add_series({
                'name': [sheet_ratesdebt, 9, 0],
                'data_labels': {'value': True, 'num_format': '#,##0.00', 'position': 'inside_base'},
                'values': [sheet_ratesdebt, 9, 1, 9, table_length]
            })
            chart_ratesdebt2 = workbook.add_chart({'type': 'line'})
            chart_ratesdebt2.set_size({'width': chart_width, 'height': chart_height})
            chart_ratesdebt2.add_series({
                'categories': [sheet_ratesdebt, 1, 1, 1, table_length],
                'values': [sheet_ratesdebt, 4, 1, 4, table_length],
                'name': [sheet_ratesdebt, 4, 0],
                'line': {'color': 'brown'},
                'data_labels': {'value': True, 'num_format': '0.00%','position': 'above'},
                'marker': {'type': 'automatic', 'color': 'brown'},
                'y2_axis': True
            })
            chart_ratesdebt1.set_title({'name': stock_name+'历年资产负债'})
            chart_ratesdebt1.set_style(10)
            chart_ratesdebt1.set_plotarea({
                'fill': {'color': '#D4FCFC'}
            })
            chart_ratesdebt1.combine(chart_ratesdebt2)
            worksheet1.insert_chart('A12', chart_ratesdebt1)
            #现金流图表
            cashflow = StockInfo.cash_flow(stock_info)
            year_list = sorted(list(cashflow.keys())[1:])
            cashflow_excle = pd.DataFrame(cashflow, index=StockInfo.cash_index, columns=year_list)
            sheet_cashflow = '现金流'
            cashflow_excle.to_excel(writer, sheet_name=sheet_cashflow, encoding='utf8', startcol=0, startrow=1)
            worksheet2 = writer.sheets[sheet_cashflow]
            worksheet2.set_column('A:H', 15)
            worksheet2.write(letter_list[math.floor((table_length - 1) / 2)] + str(1), stock_name + "现金流统计")
            worksheet2.write(letter_list[1] + str(12), "现收比=经营现金流/营业收入 现净比=经营现金流净额/净利润")
            worksheet2.write(letter_list[table_length - 1] + str(1), "单位：亿元")
            worksheet2.conditional_format('A1:' + letter_list[table_length - 1] + str(1),
                                          {'type': 'cell', 'criteria': '>=', 'value': 0, 'format': fmt})
            worksheet2.conditional_format('A2:' + letter_list[table_length - 1] + str(4),
                                          {'type': 'cell', 'criteria': '>=', 'value': 0, 'format': fmt1})
            worksheet2.conditional_format('A2:' + letter_list[table_length - 1] + str(4),
                                          {'type': 'cell', 'criteria': '<', 'value': 0, 'format': fmt4})
            worksheet2.conditional_format('A6:' + letter_list[table_length - 1] + str(7),
                                          {'type': 'cell', 'criteria': '>=', 'value': 0, 'format': fmt1})
            worksheet2.conditional_format('A6:' + letter_list[table_length - 1] + str(7),
                                          {'type': 'cell', 'criteria': '<', 'value': 0, 'format': fmt4})
            worksheet2.conditional_format('A9:' + letter_list[table_length - 1] + str(11),
                                          {'type': 'cell', 'criteria': '>=', 'value': 0, 'format': fmt1})
            worksheet2.conditional_format('A9:' + letter_list[table_length - 1] + str(11),
                                          {'type': 'cell', 'criteria': '<', 'value': 0, 'format': fmt4})
            worksheet2.conditional_format('A5:' + letter_list[table_length - 1] + str(5),
                                          {'type': 'cell', 'criteria': '>=', 'value': 0, 'format': fmt2})
            worksheet2.conditional_format('A5:' + letter_list[table_length - 1] + str(5),
                                          {'type': 'cell', 'criteria': '<', 'value': 0, 'format': fmt5})
            worksheet2.conditional_format('A8:' + letter_list[table_length - 1] + str(8),
                                          {'type': 'cell', 'criteria': '>=', 'value': 0, 'format': fmt2})
            worksheet2.conditional_format('A8:' + letter_list[table_length - 1] + str(8),
                                          {'type': 'cell', 'criteria': '<', 'value': 0, 'format': fmt5})
            worksheet2.conditional_format('A12:' + letter_list[table_length - 1] + str(12),
                                          {'type': 'cell', 'criteria': '>=', 'value': 0, 'format': fmt3})
            chart_cashflow1 = workbook.add_chart({'type': 'column'})
            chart_cashflow1.set_size({'width': chart_width, 'height': chart_height})
            chart_cashflow1.add_series({
                'categories': [sheet_cashflow, 1, 1, 1, table_length],
                'values': [sheet_cashflow, 2, 1, 2, table_length],
                'name': [sheet_cashflow, 2, 0],
                'line': {'color': 'white'}
            })
            chart_cashflow1.add_series({
                'name': [sheet_cashflow, 3, 0],
                'values': [sheet_cashflow, 3, 1, 3, table_length]
            })
            chart_cashflow1.add_series({
                'name': [sheet_cashflow, 5, 0],
                'values': [sheet_cashflow, 5, 1, 5, table_length]
            })
            chart_cashflow1.add_series({
                'name': [sheet_cashflow, 6, 0],
                'values': [sheet_cashflow, 6, 1, 6, table_length]
            })
            chart_cashflow1.add_series({
                'name': [sheet_cashflow, 9, 0],
                'values': [sheet_cashflow, 9, 1, 9, table_length]
            })

            chart_cashflow2 = workbook.add_chart({'type': 'line'})
            chart_cashflow2.set_size({'width': chart_width, 'height': chart_height})
            chart_cashflow2.add_series({
                'categories': [sheet_cashflow, 1, 1, 1, table_length],
                'values': [sheet_cashflow, 4, 1, 4, table_length],
                'name': [sheet_cashflow, 4, 0],
                'line': {'color': 'brown'},
                'data_labels': {'value': True, 'num_format': '0.00%','position': 'above'},
                'marker': {'type': 'automatic', 'color': 'brown'},
                'y2_axis': True
            })
            chart_cashflow2.add_series({
                'name': [sheet_cashflow, 7, 0],
                'values': [sheet_cashflow, 7, 1, 7, table_length],
                'data_labels': {'value': True, 'num_format': '0.00%','position': 'center'},
                'marker': {'type': 'automatic', 'color': 'green'},
                'y2_axis': True
            })
            chart_cashflow1.set_title({'name': stock_name + '历年现金流'})
            chart_cashflow1.set_style(10)
            chart_cashflow1.set_plotarea({
                'fill': {'color': '#D4FCFC'}
            })
            chart_cashflow1.combine(chart_cashflow2)
            chart_cashflow3 = workbook.add_chart({'type': 'column'})
            chart_cashflow3.set_size({'width': chart_width, 'height': chart_height})
            chart_cashflow3.add_series({
                'categories': [sheet_cashflow, 1, 1, 1, table_length],
                'values': [sheet_cashflow, 6, 1, 6, table_length],
                'name': [sheet_cashflow, 6, 0],
                'data_labels': {'value': True, 'num_format': '#,##0.00'},
                'line': {'color': 'white'}
            })
            chart_cashflow3.add_series({
                'name': [sheet_cashflow, 9, 0],
                'data_labels': {'value': True, 'num_format': '#,##0.00'},
                'values': [sheet_cashflow, 9, 1, 9, table_length]
            })
            chart_cashflow3.add_series({
                'name': [sheet_cashflow, 8, 0],
                'data_labels': {'value': True, 'num_format': '#,##0.00'},
                'values': [sheet_cashflow, 8, 1, 8, table_length]
            })
            chart_cashflow3.add_series({
                'name': [sheet_cashflow, 10, 0],
                'data_labels': {'value': True, 'num_format': '#,##0.00'},
                'values': [sheet_cashflow, 10, 1, 10, table_length]
            })
            chart_cashflow3.set_title({'name': stock_name + '历年投融资和资本开支'})
            chart_cashflow3.set_plotarea({
                'fill': {'color': '#D4FCFC'}
            })
            worksheet2.insert_chart('A13', chart_cashflow1)
            worksheet2.insert_chart('A32', chart_cashflow3)
            #经营情况图表
            bussinessincome = StockInfo.bussiness_income(stock_info)
            year_list = sorted(list(bussinessincome.keys())[1:])
            bussinessincome_excle = pd.DataFrame(bussinessincome, index=StockInfo.bussiness_index, columns=year_list)
            sheet_bussiness = '经营情况'
            bussinessincome_excle.to_excel(writer, sheet_name=sheet_bussiness, encoding='utf8', startcol=0, startrow=1)
            worksheet3 = writer.sheets[sheet_bussiness]
            worksheet3.set_column('A:H', 15)
            worksheet3.write(letter_list[math.floor((table_length - 1) / 2)] + str(1), stock_name + "经营情况")
            #worksheet3.write(letter_list[1] + str(12), "现收比=经营现金流/营业收入 现净比=经营现金流净额/净利润")
            worksheet3.write(letter_list[table_length - 1] + str(1), "单位：亿元")
            worksheet3.conditional_format('A1:' + letter_list[table_length - 1] + str(1),
                                          {'type': 'cell', 'criteria': '>=', 'value': 0, 'format': fmt})
            worksheet3.conditional_format('A2:' + letter_list[table_length - 1] + str(5),
                                          {'type': 'cell', 'criteria': '>=', 'value': 0, 'format': fmt1})
            worksheet3.conditional_format('A2:' + letter_list[table_length - 1] + str(5),
                                          {'type': 'cell', 'criteria': '<', 'value': 0, 'format': fmt4})
            worksheet3.conditional_format('A6:' + letter_list[table_length - 1] + str(8),
                                          {'type': 'cell', 'criteria': '>=', 'value': 0, 'format': fmt2})
            worksheet3.conditional_format('A6:' + letter_list[table_length - 1] + str(8),
                                          {'type': 'cell', 'criteria': '<', 'value': 0, 'format': fmt5})
            worksheet3.conditional_format('A9:' + letter_list[table_length - 1] + str(12),
                                          {'type': 'cell', 'criteria': '>=', 'value': 0, 'format': fmt1})
            worksheet3.conditional_format('A9:' + letter_list[table_length - 1] + str(12),
                                          {'type': 'cell', 'criteria': '<', 'value': 0, 'format': fmt4})
            worksheet3.conditional_format('A13:' + letter_list[table_length - 1] + str(20),
                                          {'type': 'cell', 'criteria': '>=', 'value': 0, 'format': fmt2})
            worksheet3.conditional_format('A13:' + letter_list[table_length - 1] + str(20),
                                          {'type': 'cell', 'criteria': '<', 'value': 0, 'format': fmt5})
            # worksheet3.conditional_format('A12:' + letter_list[table_length - 1] + str(12),
            #                               {'type': 'cell', 'criteria': '>=', 'value': 0, 'format': fmt3})
            chart_business1 = workbook.add_chart({'type': 'column'})
            chart_business1.set_size({'width': chart_width, 'height': chart_height})
            chart_business1.add_series({
                'categories': [sheet_bussiness, 1, 1, 1, table_length],
                'values': [sheet_bussiness, 2, 1, 2, table_length],
                'name': [sheet_bussiness, 2, 0],
                'line': {'color': 'white'}
            })
            chart_business1.add_series({
                'name': [sheet_bussiness, 3, 0],
                'values': [sheet_bussiness, 3, 1, 3, table_length]
            })
            chart_business1.add_series({
                'name': [sheet_bussiness, 4, 0],
                'values': [sheet_bussiness, 4, 1, 4, table_length]
            })
            chart_business2 = workbook.add_chart({'type': 'line'})
            chart_business2.set_size({'width': chart_width, 'height': chart_height})
            chart_business2.add_series({
                'categories': [sheet_bussiness, 1, 1, 1, table_length],
                'name': [sheet_bussiness, 16, 0],
                'values': [sheet_bussiness, 16, 1, 16, table_length],
                'data_labels': {'value': True, 'num_format': '0.00%', 'position': 'above',
                                'font': {'name': 'Consolas', 'color': 'black'}},
                'marker': {'type': 'automatic', 'color': 'blue'},
                'y2_axis': True
            })
            chart_business2.add_series({
                'name': [sheet_bussiness, 17, 0],
                'values': [sheet_bussiness, 17, 1, 17, table_length],
                'data_labels': {'value': True, 'num_format': '0.00%', 'position': 'right',
                                'font': {'name': 'Consolas', 'color': 'black'}},
                'marker': {'type': 'automatic', 'color': 'blue'},
                'y2_axis': True
            })
            chart_business2.add_series({
                'name': [sheet_bussiness, 18, 0],
                'values': [sheet_bussiness, 18, 1, 18, table_length],
                'data_labels': {'value': True, 'num_format': '0.00%', 'position': 'below',
                                'font': {'name': 'Consolas', 'color': 'black'}},
                'marker': {'type': 'automatic', 'color': 'blue'},
                'y2_axis': True
            })
            chart_business2.add_series({
                'name': [sheet_bussiness, 19, 0],
                'values': [sheet_bussiness, 19, 1, 19, table_length],
                'data_labels': {'value': True, 'num_format': '0.00%', 'position': 'below',
                                'font': {'name': 'Consolas', 'color': 'black'}},
                'marker': {'type': 'automatic', 'color': 'blue'},
                'y2_axis': True
            })
            chart_business1.set_title({'name': stock_name + '历年经营情况'})
            chart_business1.set_style(10)
            chart_business1.set_plotarea({
                'fill': {'color': '#D4FCFC'}
            })
            chart_business1.combine(chart_business2)
            chart_business3 = workbook.add_chart({'type': 'column'})
            chart_business3.set_size({'width': chart_width, 'height': chart_height})
            chart_business3.add_series({
                'categories': [sheet_bussiness, 1, 1, 1, table_length],
                'values': [sheet_bussiness, 8, 1, 8, table_length],
                'name': [sheet_bussiness, 8, 0],
                'line': {'color': 'white'}
            })
            chart_business3.add_series({
                'name': [sheet_bussiness, 9, 0],
                'values': [sheet_bussiness, 9, 1, 9, table_length]
            })
            chart_business3.add_series({
                'name': [sheet_bussiness, 10, 0],
                'values': [sheet_bussiness, 10, 1, 10, table_length]
            })
            chart_business3.add_series({
                'name': [sheet_bussiness, 11, 0],
                'values': [sheet_bussiness, 11, 1, 11, table_length]
            })
            chart_business3.set_title({'name': stock_name + '历年经营和费用情况'})
            chart_business3.set_plotarea({
                'fill': {'color': '#D4FCFC'}
            })
            chart_business4 = workbook.add_chart({'type': 'line'})
            chart_business4.set_size({'width': chart_width, 'height': chart_height})
            chart_business4.add_series({
                'categories': [sheet_bussiness, 1, 1, 1, table_length],
                'values': [sheet_bussiness, 5, 1, 5, table_length],
                'name': [sheet_bussiness, 5, 0],
                #'data_labels': {'value': True, 'num_format': '0.00%', 'position': 'above','font': {'name': 'Consolas', 'color': 'brown'}},
                'marker': {'type': 'automatic', 'color': 'brown'},
                'y2_axis': True
            })
            chart_business4.add_series({
                'values': [sheet_bussiness, 12, 1, 12, table_length],
                'name': [sheet_bussiness, 12, 0],
                #'data_labels': {'value': True, 'num_format': '0.00%', 'position': 'below','font': {'name': 'Consolas', 'color': 'red'}},
                'marker': {'type': 'automatic', 'color': 'brown'},
                'y2_axis': True
            })
            chart_business4.add_series({
                'values': [sheet_bussiness, 13, 1, 13, table_length],
                'name': [sheet_bussiness, 13, 0],
                #'data_labels': {'value': True, 'num_format': '0.00%', 'position': 'right','font': {'name': 'Consolas', 'color': 'black'}},
                'marker': {'type': 'automatic', 'color': 'brown'},
                'y2_axis': True
            })
            chart_business4.add_series({
                'values': [sheet_bussiness, 14, 1, 14, table_length],
                'name': [sheet_bussiness, 14, 0],
                #'data_labels': {'value': True, 'num_format': '0.00%', 'position': 'left','font': {'name': 'Consolas', 'color': 'blue'}},
                'marker': {'type': 'automatic', 'color': 'brown'},
                'y2_axis': True
            })
            chart_business4.add_series({
                'values': [sheet_bussiness, 15, 1, 15, table_length],
                'name': [sheet_bussiness, 15, 0],
                #'data_labels': {'value': True, 'num_format': '0.00%', 'position': 'below','font': {'name': 'Consolas', 'color': 'green'}},
                'marker': {'type': 'automatic', 'color': 'brown'},
                'y2_axis': True
            })
            chart_business3.combine(chart_business4)
            worksheet3.insert_chart('A21', chart_business1)
            worksheet3.insert_chart('A40', chart_business3)
            # 资产结构图表
            assetstructure = StockInfo.asset_structure(stock_info)
            year_list = sorted(list(assetstructure.keys())[1:])
            assetstructure_excle = pd.DataFrame(assetstructure, index=StockInfo.assets_index, columns=year_list)
            sheet_asset = '资产结构及周转'
            assetstructure_excle.to_excel(writer, sheet_name=sheet_asset, encoding='utf8', startcol=0, startrow=1)
            worksheet4 = writer.sheets[sheet_asset]
            worksheet4.set_column('A:H', 15)
            worksheet4.write(letter_list[math.floor((table_length - 1) / 2)] + str(1), stock_name + "资产结构")
            worksheet4.write(letter_list[1] + str(11), "固定和无形资产占比=(固定资产+无形资产)/总资产 在建工程占比=在建工程/总资产")
            worksheet4.write(letter_list[1] + str(12), "应收账款收入比=应收账款/营业收入 存货成本比=存货/营业成本")
            worksheet4.write(letter_list[table_length - 1] + str(1), "单位：亿元")
            worksheet4.conditional_format('A1:' + letter_list[table_length - 1] + str(1),
                                          {'type': 'cell', 'criteria': '>=', 'value': 0, 'format': fmt})
            worksheet4.conditional_format('A7:' + letter_list[table_length - 1] + str(10),
                                          {'type': 'cell', 'criteria': '>=', 'value': 0, 'format': fmt1})
            worksheet4.conditional_format('A7:' + letter_list[table_length - 1] + str(10),
                                          {'type': 'cell', 'criteria': '<', 'value': 0, 'format': fmt4})
            worksheet4.conditional_format('A2:' + letter_list[table_length - 1] + str(6),
                                          {'type': 'cell', 'criteria': '>=', 'value': 0, 'format': fmt2})
            worksheet4.conditional_format('A2:' + letter_list[table_length - 1] + str(6),
                                          {'type': 'cell', 'criteria': '<', 'value': 0, 'format': fmt5})
            worksheet4.conditional_format('A11:' + letter_list[table_length - 1] + str(12),
                                          {'type': 'cell', 'criteria': '>=', 'value': 0, 'format': fmt3})
            chart_asset1 = workbook.add_chart({'type': 'column'})
            chart_asset1.set_size({'width': chart_width, 'height': chart_height})
            chart_asset1.add_series({
                'categories': [sheet_asset, 1, 1, 1, table_length],
                'values': [sheet_asset, 6, 1, 6, table_length],
                'name': [sheet_asset, 6, 0],
                'line': {'color': 'white'}
            })
            chart_asset1.add_series({
                'name': [sheet_asset, 7, 0],
                'values': [sheet_asset, 7, 1, 7, table_length]
            })
            chart_asset1.add_series({
                'name': [sheet_asset, 8, 0],
                'values': [sheet_asset, 8, 1, 8, table_length]
            })
            chart_asset1.add_series({
                'name': [sheet_asset, 9, 0],
                'values': [sheet_asset, 9, 1, 9, table_length]
            })
            chart_asset2 = workbook.add_chart({'type': 'line'})
            chart_asset2.set_size({'width': chart_width, 'height': chart_height})
            chart_asset2.add_series({
                'categories': [sheet_asset, 1, 1, 1, table_length],
                'name': [sheet_asset, 2, 0],
                'values': [sheet_asset, 2, 1, 2, table_length],
                'data_labels': {'value': True, 'num_format': '0.00%', 'position': 'right',
                                'font': {'name': 'Consolas', 'color': 'black'}},
                'marker': {'type': 'automatic', 'color': 'blue'},
                'y2_axis': True
            })
            chart_asset2.add_series({
                'name': [sheet_asset, 3, 0],
                'values': [sheet_asset, 3, 1, 3, table_length],
                'data_labels': {'value': True, 'num_format': '0.00%', 'position': 'right',
                                'font': {'name': 'Consolas', 'color': 'black'}},
                'marker': {'type': 'automatic', 'color': 'blue'},
                'y2_axis': True
            })
            chart_asset2.add_series({
                'name': [sheet_asset, 4, 0],
                'values': [sheet_asset, 4, 1, 4, table_length],
                'data_labels': {'value': True, 'num_format': '0.00%', 'position': 'left',
                                'font': {'name': 'Consolas', 'color': 'black'}},
                'marker': {'type': 'automatic', 'color': 'blue'},
                'y2_axis': True
            })
            chart_asset2.add_series({
                'name': [sheet_asset, 5, 0],
                'values': [sheet_asset, 5, 1, 5, table_length],
                'data_labels': {'value': True, 'num_format': '0.00%', 'position': 'left',
                                'font': {'name': 'Consolas', 'color': 'black'}},
                'marker': {'type': 'automatic', 'color': 'blue'},
                'y2_axis': True
            })
            chart_asset1.set_title({'name': stock_name + '历年资产结构及周转'})
            chart_asset1.set_style(10)
            chart_asset1.set_plotarea({
                'fill': {'color': '#D4FCFC'}
            })
            chart_asset1.combine(chart_asset2)
            worksheet4.insert_chart('A13', chart_asset1)

            # 股东质押信息
            zyinfo = StockInfo.zy_info(codes)
            if zyinfo == None:
                pass
            else:
                zytable_hight = len(zyinfo)+2
                zyinfo_excle = pd.DataFrame(zyinfo,columns=StockInfo.zy_index)
                sheet_zyinfo = '股份质押'
                zyinfo_excle.to_excel(writer, sheet_name=sheet_zyinfo, encoding='utf8', startcol=0, startrow=1)
                worksheet5 = writer.sheets[sheet_zyinfo]
                worksheet5.set_column('A:A', 5)
                worksheet5.set_column('B:B', 20)
                worksheet5.set_column('C:H', 10)
                worksheet5.set_column('I:I', 25)
                worksheet5.set_column('J:J', 60)
                worksheet5.set_column('K:M', 15)
                worksheet5.write('E1', stock_name + "股份质押统计")
                worksheet5.write('H1', "单位：亿")
                worksheet5.conditional_format('A1:M1',
                                              {'type': 'cell', 'criteria': '>=', 'value': 0, 'format': fmt})
                worksheet5.conditional_format('A2:D'+str(zytable_hight),
                                              {'type': 'cell', 'criteria': '>=', 'value': 0, 'format': fmt1})
                worksheet5.conditional_format('A2:D'+str(zytable_hight),
                                              {'type': 'cell', 'criteria': '<', 'value': 0, 'format': fmt4})
                worksheet5.conditional_format('E2:F'+str(zytable_hight),
                                              {'type': 'cell', 'criteria': '>=', 'value': 0, 'format': fmt2})
                worksheet5.conditional_format('F2:F'+str(zytable_hight),
                                              {'type': 'cell', 'criteria': '<', 'value': 0, 'format': fmt5})
                worksheet5.conditional_format('G2:M' + str(zytable_hight),
                                              {'type': 'cell', 'criteria': '>=', 'value': 0, 'format': fmt1})
                worksheet5.conditional_format('G2:M' + str(zytable_hight),
                                              {'type': 'cell', 'criteria': '<', 'value': 0, 'format': fmt4})
            writer.save()


if __name__ == '__main__':

    StockInfo().get_stockdata("SZ002129","中环股份","2017-12-31,2018-12-31,2019-12-31,2020-12-31,2021-12-31")
