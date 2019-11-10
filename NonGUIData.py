#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Wed Nov  6 21:29:28 2019
Splits QuickBooks general ledgers into a dicttionary of pandas dataframes
@author: michael
"""
import importlib
import pandas as pd
import numpy as np


class QuickBooks:

    def __init__(self, file_name, sections=None):

        if sections is None:
            sections = []

        df = pd.read_excel(file_name, header=None)

        comp_name = df.iloc[0, 0].split(' ')[0].replace(',', '')
        imported_module = importlib.import_module(comp_name)

        name = df.iloc[0, 0] + ' ' + df.iloc[1, 0] + ' ' + df.iloc[2, 0]
        name.replace('  ', ' ')
        cols = df.iloc[4, 1:].tolist()
        data_set = df.drop([0, 1, 2, 3, 4]).reset_index(drop=True)
        data_set = data_set[data_set[1] != "Beginning Balance"]

        frame = []
        self.frame_dict = {}
        frame_dict_keys = []

        for idx, line in enumerate(data_set[1]):
            if (pd.isna(line) is False) and (pd.isna(data_set.iloc[idx-1, 1])):
                frame_title = data_set.iloc[idx - 1, 0].strip()
                frame_dict_keys.append(frame_title)
                frame.append(data_set.iloc[idx, :])
            elif pd.isna(line) is False:
                frame.append(data_set.iloc[idx, :])
            elif (idx != 0) and (pd.isna(data_set.iloc[idx-1, 1]) is False):
                temp_frame = pd.DataFrame(frame).drop(0, axis=1).reset_index(drop=True)
                temp_frame.columns = cols
                self.frame_dict[frame_title] = temp_frame
                frame = []

        if len(sections) != 0:
            for key in frame_dict_keys:
                if key[0:2] not in sections:
                    del self.frame_dict[key]

        pvt_dict = {}
        for choice in self.frame_dict:
            data = self.frame_dict[choice]
            vendor_list_1 = imported_module.first_parser(data)
            vendor_list_2 = imported_module.second_parser(data)
            vendors = imported_module.third_parser(vendor_list_1, vendor_list_2)
            vendors = imported_module.fourth_parser(vendors)
            if choice == '6122 Computer & Software':
                vendors = imported_module.fifth_parser(vendors)

            data['Vendors'] = vendors
            data['Amount'] = pd.to_numeric(data['Amount'])
            data['Month'] = pd.to_datetime(data['Date']).dt.month
            data['Month'] = data['Month'].apply(lambda x: imported_module.months[x])

            pvt = pd.pivot_table(data, values=['Amount'], index=['Vendors'], columns=['Month'], aggfunc=sum)
            pvt = pvt['Amount']

            for index in range(1, 10):
                if imported_module.months[index] not in pvt.columns:
                    pvt[imported_module.months[index]] = np.nan

            new_col = []
            for index in range(1, 10):
                for element in pvt.columns:
                    if imported_module.months[index] == element:
                        new_col.append(element)
                        break
            pvt.columns = new_col

            pvt[''] = np.nan
            pvt['Vendor Mean'] = pvt.mean(axis=1).round(2)
            pvt['Vendor Total'] = pvt[[imported_module.months[key] for key in range(1, 10)]].sum(axis=1)
            pvt = pvt.sort_values(by=['Vendor Mean'], ascending=False)
            mean_series = pd.Series(pvt[[imported_module.months[key] for key in range(1, 10)]].mean(axis=0).round(2))
            sum_series = (pd.Series(pvt[[imported_module.months[key] for key in range(1, 10)]]
                                    .sum(min_count=1, axis=0).round(2)))
            mean_series.name = 'Monthly Mean'
            sum_series.name = 'Monthly Total'
            pvt = pvt.append(pd.Series(name=''))
            pvt = pvt.append(mean_series)
            pvt = pvt.append(sum_series)
            pvt.loc[['Monthly Mean'], ['Vendor Total']] = "Grand Total"
            pvt.loc[['Monthly Total'], ['Vendor Total']] = sum_series.sum(min_count=1).round(2)
            pvt_dict[choice] = pvt

            self.pvt_dict = pvt_dict

        def highlight_cols1():
            color = '#FFEB9B'
            return 'background-color: %s' % color

        def highlight_cols2():
            color = '#FCD5B4'
            return 'background-color: %s' % color

        def highlight_cols3():
            color = '#CCC0DA'
            return 'background-color: %s' % color

        def highlight_cols4():
            color = '#8065A1'
            return 'background-color: %s' % color

        def highlight_cols5():
            color = '#B8CBE3'
            return 'background-color: %s' % color

        def highlight_cols6():
            color = '#EBF0DE'
            return 'background-color: %s' % color

        def cent():
            return 'text-align: center'

        def bolder():
            return 'font-weight: bold'

        def color_white():
            return 'color: white'

        def full_bord():
            return 'border-style: solid'

        def bottom_bord():
            return 'border-bottom-style: solid'

        def right_bord():
            return 'border-right-style: solid'

        def bord_weight():
            return 'border-width: 1px'

        def bord_color():
            return 'border-color: black'

        sub_dict = {}
        for item in self.pvt_dict:
            sub = self.pvt_dict[item].reset_index()
            sub = sub.T.reset_index().T.reset_index(drop=True)
            r, c = sub.shape
            sub = sub.style.applymap(highlight_cols1, subset=pd.IndexSlice[r-3, :c-3])\
                .applymap(highlight_cols1, subset=pd.IndexSlice[:r-4, c-3])\
                .applymap(highlight_cols2, subset=pd.IndexSlice[0, c-2])\
                .applymap(bolder, subset=pd.IndexSlice[0, c-2])\
                .applymap(cent, subset=pd.IndexSlice[0, c-2])\
                .applymap(highlight_cols3, subset=pd.IndexSlice[0, c-1])\
                .applymap(bolder, subset=pd.IndexSlice[0, c-1])\
                .applymap(cent, subset=pd.IndexSlice[0, c-1])\
                .applymap(highlight_cols2, subset=pd.IndexSlice[r-2, 0])\
                .applymap(bolder, subset=pd.IndexSlice[r-2, 0])\
                .applymap(cent, subset=pd.IndexSlice[r-2, 0])\
                .applymap(highlight_cols3, subset=pd.IndexSlice[r-1, 0])\
                .applymap(bolder, subset=pd.IndexSlice[r-1, 0])\
                .applymap(cent, subset=pd.IndexSlice[r-1, 0])\
                .applymap(highlight_cols4, subset=pd.IndexSlice[r-2, c-1])\
                .applymap(bolder, subset=pd.IndexSlice[r-2, c-1])\
                .applymap(cent, subset=pd.IndexSlice[r-2, c-1])\
                .applymap(color_white, subset=pd.IndexSlice[r-2, c-1])\
                .applymap(highlight_cols5, subset=pd.IndexSlice[0, :c-4])\
                .applymap(bolder, subset=pd.IndexSlice[0, :c-4])\
                .applymap(cent, subset=pd.IndexSlice[0, :c-4])\
                .applymap(highlight_cols6, subset=pd.IndexSlice[1:r-4, 0])\
                .applymap(bolder, subset=pd.IndexSlice[1:r-4, 0])\
                .applymap(cent, subset=pd.IndexSlice[1:r-4, 0])\
                .applymap(full_bord, subset=pd.IndexSlice[:r-4, :c-4])\
                .applymap(bord_weight, subset=pd.IndexSlice[:r-4, :c-4])\
                .applymap(full_bord, subset=pd.IndexSlice[r-2:, :c-4])\
                .applymap(bord_weight, subset=pd.IndexSlice[r-2:, :c-4])\
                .applymap(full_bord, subset=pd.IndexSlice[:r-4, c-2:])\
                .applymap(bord_weight, subset=pd.IndexSlice[:r-4, c-2:])\
                .applymap(full_bord, subset=pd.IndexSlice[r-2:r-1, c-1])\
                .applymap(bord_weight, subset=pd.IndexSlice[r-2:r-1, c-1])\
                .applymap(bord_color, subset=pd.IndexSlice[r-2:r-1, c-1])\
                .applymap(bottom_bord, subset=pd.IndexSlice[r-3, c-3])\
                .applymap(right_bord, subset=pd.IndexSlice[r-3, c-3])\
                .applymap(bord_weight, subset=pd.IndexSlice[r-3, c-3])
            sub_dict[item] = sub

            self.sub_dict = sub_dict

    def excel_printer(self):
        writer = pd.ExcelWriter('/path/output.xlsx')
        for sheet, frame in self.sub_dict.items():
            new_sheet = sheet.replace('/', '-')[:30]
            frame.to_excel(writer, sheet_name=new_sheet, index=False, header=False)
        writer.save()

    def data_keys(self):
        for item in self.frame_dict.keys():
            print(item)

    def data_view(self, key):
        return self.frame_dict[key]

    def pvt_view(self, key):
        return self.pvt_dict[key]
