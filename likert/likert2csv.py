#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import sys
import math
import pandas as pd
import numpy as np
import matplotlib
import matplotlib.pyplot as plt
from   optparse import OptionParser
from   xlsxwriter.utility import xl_rowcol_to_cell


class Report:

    def __init__(self, filename):
        self.all_top_box = None
        self.all_average = None

        self.split_reports_by = u'custom_3'
        self.likert_names = ["SD", "D", "N", "A", "SA"]

        df = pd.read_excel(filename)
        self.df = self.drop_unwanted_columns(df, keep=[self.split_reports_by])
        questions, question_key_list = self.get_questions_in_order()


    def drop_unwanted_columns(self, df, keep=[]):
        # The SurveyMonkey report might have unwanted rows
        for column_name in df.columns.tolist():
            if column_name in keep:
                continue
            if df[column_name][0] != "Response":
                df = df.drop(columns=[column_name])
        return df

    def get_questions_in_order(self, key_prefix = 'F%2i'):
        # SurveyMonkey adds the key word "Response"
        ret = {}
        key_list = []
        index = 0
        for column_name in self.df.columns.tolist():
            if self.df[column_name][0] == "Response":
                index = index + 1
                ret[key_prefix % index] = column_name
                key_list.append(key_prefix % index)
        self.questions = ret
        self.question_key_list = key_list
        return ret, key_list

    def report_total(self):
        report_name = "ALL"
        print("Processing Report: %s" % report_name)
        df = self.df.copy(deep=True)
        df = df.drop(columns=[self.split_reports_by])
        df = pd.DataFrame(df.stack())
        df = pd.DataFrame(df.unstack(0))
        df = df.drop(columns=[(0,0)])
        self.generate_report(df, report_name)

    def report(self):
        for report_name in self.df[self.split_reports_by].unique():
            # filter garbage out
            if type(report_name) != str:
                continue

            print("Processing Report: %s" % report_name)
            df = self.df.copy(deep=True)
            df = df[df[self.split_reports_by] == report_name]
            df = df.drop(columns=[self.split_reports_by])
            df = pd.DataFrame(df.stack())
            df = pd.DataFrame(df.unstack(0))
            self.generate_report(df, report_name)


    def write_xlsx(self, df, name):
        """
        Make a shiny XLSX
        """

        # Add full question texts
        for key in self.questions:
            df.loc[key,'Q'] = self.questions[key]

        # Swap question with question-key
        col_index = ["Q", 'CNT', 'AVG', "CAV", "BOX", "CBO" ] # 'SD','D', 'N','A', 'SA'
        df = df.reindex(col_index, axis=1)

        # Numan readable column names
        df = df.rename(columns={
        "SD":  "Strongly disagree",
        "D":   "Disagree",
        "N":   "Neither agree nor disagree",
        "A":   "Agree",
        "SA":  "Strongly agree",
        "CNT": "Anzahl Antworten",
        "AVG": "Your Average",
        "CAV": "Company Average",
        "BOX": "Your Score",
        "CBO": "Company Score",
        "Q":   "Frage"
        })

        # open the XLSX writer
        writer = pd.ExcelWriter(name + '.xlsx', engine='xlsxwriter')
        sheet  = "Report"

        # add data frame to sheet
        df.to_excel(writer, sheet)

        # define and set number formats
        workbook = writer.book
        worksheet = writer.sheets[sheet]

        # set default cell format
        workbook.formats[0].set_font_size(10)
        workbook.formats[0].set_font_name('Arial')

        # https://xlsxwriter.readthedocs.io/format.html
        format1 = workbook.add_format({'num_format': '#,##0.00'})
        format2 = workbook.add_format({'num_format': '0%'})
        format3 = workbook.add_format({'bg_color'  : '008046',
                                       'font_color': 'ffffff'})
        format4 = workbook.add_format({'bg_color'  : 'F79646'})
        fromat5   = workbook.add_format({'rotation'  : 90,
                                       'bg_color'  : 'F2F2F2' })

        # set column formats based on index
        worksheet.set_row   (0, 20, cell_format = fromat5)
        worksheet.set_column(col_index.index('AVG') + 1,
                             col_index.index('CAV') + 1,
                             width = 10, cell_format = format1)
        worksheet.set_column(col_index.index('BOX') + 1,
                             col_index.index('CBO') + 1,
                             width = 10, cell_format = format2)

        # add conditional formats
        # https://xlsxwriter.readthedocs.io/working_with_conditional_formats.html
        compare_to = xl_rowcol_to_cell(1, col_index.index('CBO') + 1,
                                       col_abs = True)
        col  = col_index.index('BOX') + 1
        worksheet.conditional_format(1, col, 30, col,{'type':     'cell',
                                                'criteria': '>=',
                                                'value':    f'{compare_to}+0.15',
                                                'format':   format3
                                                })
        worksheet.conditional_format(1, col, 30, col, {'type':     'cell',
                                                'criteria': '<=',
                                                'value':    f'{compare_to}-0.15',
                                                'format':   format4
                                                })

        # final save
        writer.save()


    def generate_report(self, df,  name):

        g_fnc = lambda x,y: x.loc[y] if y in x.index else 0
        columns_names = ["SD", "D", "N", "A", "SA"]

        # Count all values by column name
        for key in self.question_key_list:
            for i in range(1, len(columns_names) + 1):
                counts = g_fnc (df.loc[self.questions[key],0].value_counts(), i)
                df.loc[self.questions[key], columns_names[i-1]] = int(counts)

        # Delete the raw survey data
        df.drop(0, axis=1, inplace=True)

        for key in self.questions:
            df.loc[self.questions[key],'ID'] = key
        df = df.set_index('ID')

        for key in self.question_key_list:

            topbox  = [0, 0, 0, 1, 1]
            botbox  = [1, 1, 1, 0, 0]
            scores  = [1, 2, 3, 4, 5]
            weights = [2, 1, 0, 1, 2]
            respons = df.loc[key,columns_names]
            topvalue = (respons * topbox).sum()
            botvalue = (respons * botbox).sum()

            df.loc[key, "CNT"] = respons.sum()
            df.loc[key, "AVG"] = (respons * scores).sum() / respons.sum()
            if ((respons * weights).sum() > 0):
                df.loc[key, "AVGW"] = (respons * weights * scores).sum() / (respons * weights).sum()
            else:
                df.loc[key, "AVGW"] = 0
            df.loc[key, "BOX"] = (1 / (topvalue + botvalue) * topvalue)

        if name == 'ALL':
            self.all_top_box = df.loc[:,"BOX"]
            self.all_average = df.loc[:,"AVG"]

        df['CAV'] = self.all_average
        df['CBO'] = self.all_top_box

        df[columns_names] = df[["SD", "D", "N", "A", "SA"]].fillna(0.0).astype(int)
        df.columns = df.columns.droplevel(level=1)

        self.write_xlsx(df, name)



if __name__ == "__main__":

    parser = OptionParser()

    parser.add_option("-v", "--verbose",
        action="store_true", dest="verbose", default="False",
        help="Verbose prints enabled")


    parser.add_option("-f", "--file", dest="filename",
                  help="write report to FILE", metavar="FILE")

    (options, args) = parser.parse_args()

    r = Report(options.filename)
    r.report_total()
    r.report()


