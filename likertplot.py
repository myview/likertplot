#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import sys
import math
import pandas as pd
import numpy as np
import matplotlib
import matplotlib.pyplot as plt
from optparse import OptionParser


def get_questions_in_order(df, key_prefix = 'F%2i'):

    # SurveyMonkey adds the key word "Response"
    ret = {}
    key_list = []
    index = 0
    for column_name in df.columns.tolist():
        if df[column_name][0] == "Response":  
            index = index + 1 
            ret[key_prefix % index] = column_name
            key_list.append(key_prefix % index)
    return ret, key_list

def drop_unwanted_columns(df, keep=[]):

    # The SurveyMonkey report might have unwanted rows
    for column_name in df.columns.tolist():
        if column_name in keep:
            continue
        if df[column_name][0] != "Response":
            df = df.drop(columns=[column_name])
    return df

def main(options):

    split_reports_by = u'custom_3'
    columns_names = ["SD", "D", "N", "A", "SA"]
    likert_colors = ['white', 'firebrick','lightcoral','gainsboro','cornflowerblue', 'darkblue']
    g_fnc = lambda x,y: x.loc[y] if y in x.index else 0

    df = pd.read_excel(options.filename)
    df = drop_unwanted_columns(df, keep=[split_reports_by])
    questions, question_key_list = get_questions_in_order(df)

    if options.questions:
        for key in questions:
            print("%s: %s" % (key, questions[key]))
        sys.exit(0)

    for report_name in df[split_reports_by].unique():
        # filter garbage out
        if type(report_name) != str:
            continue

        print("Processing Report: %s" % report_name)
        report_data = df.copy(deep=True)
        report_data = report_data[report_data[split_reports_by] == report_name]
        report_data = report_data.drop(columns=[split_reports_by])
        report_data=pd.DataFrame(report_data.stack())
        report_data=pd.DataFrame(report_data.unstack(0))

        # Count all values by column name
        for key in question_key_list:
            for i in range(1, len(columns_names) + 1):
                counts = g_fnc (report_data.loc[questions[key],:].value_counts(), i)
                report_data.loc[questions[key], columns_names[i-1]] = int(counts)
        report_data[columns_names] = report_data[["SD", "D", "N", "A", "SA"]].fillna(0.0).astype(int)

        # Swap question with question-key
        for key in questions:
            report_data.loc[questions[key],'Questions'] = key
        report_data = report_data.set_index('Questions')

        # Remove remaining SurveyMonkey Data which is on index 0
        report_data.drop(0, axis=1,inplace=True) 
        report_data.columns = report_data.columns.droplevel(level=1)

        # Change the order in the plot
        # report_data = report_data.loc[reversed(question_key_list), :]

        if options.verbose == True:
            print(report_data)


        middles = report_data[["SD", "D"]].sum(axis=1)+report_data["N"]*.5
        toppers = report_data[["SA", "A"]].sum(axis=1)+report_data["N"]*.5

        total = int(middles.max() + toppers.max())

        longest = middles.max()
        complete_longest = report_data.sum(axis=1).max()
        report_data.insert(0, '', (middles - longest).abs())
        report_data = report_data.sort_index()   

        report_data.plot.barh(stacked=True, 
                 color=likert_colors, 
                 edgecolor='none', 
                 legend=False)

        z = plt.axvline(longest, 
                    linestyle='--', 
                    color='red', 
                    alpha=.5)
        z.set_zorder(-1)


        range_steps = 1
        if total > 10:
            range_steps = 10

        xvalues = range(0,total+1, range_steps)
        xlabels = [str(x-longest) for x in xvalues]

        plt.xticks(xvalues, xlabels)
        plt.savefig(report_name + '.png')


if __name__ == "__main__":

    parser = OptionParser()

    parser.add_option("-v", "--verbose",
        action="store_true", dest="verbose", default="False",
        help="Verbose prints enabled")

    parser.add_option("-q", "--questions", dest="questions",
                  action="store_true",
                  help="print questions")

    parser.add_option("-f", "--file", dest="filename",
                  help="write report to FILE", metavar="FILE")

    (options, args) = parser.parse_args()

    main(options)


