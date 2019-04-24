#!/usr/bin/python3

from pprint import pprint
from optparse import OptionParser
import pandas as pd
import json
import sys
import re
import os
from datetime import datetime

class Process:


    def __init__(self, options):
        """
        """
        self.options = options

        # Set directory where the file is located
        path = os.path.abspath(options.filename)
        self.path = os.path.dirname(path)
        self.filename = os.path.basename(path)

        print(f'> set working path to: {self.path}')
        print(f'> set file: {self.filename}')

        with open(os.path.join(self.path, 'history-2019-01-01.json')) as json_file:
            self.history_2 = json.load(json_file)

        with open(os.path.join(self.path, 'master.json')) as json_file:
            self.master = json.load(json_file)

        self.history = {}

        self.date = datetime.strptime(options.date, '%Y.%M.%d')

    def get_id(self, name):

        """
        https://pythex.org
        https://tinyurl.com/yd6p4kbz
        """
        m = re.search('^([A-Z]{2,}|[0-9]{2,})', name)
        if m:
            return m.group(0)
        return name

    def get_vorgesetzter_by_id(self, name):
        return name

    def get_gruppe_by_id(self, idx):
        return self.get_layer_by_id(idx, 'Gruppe')

    def get_abteilung_by_id(self, idx):
        return self.get_layer_by_id(idx, 'Abteilung')

    def get_team_by_id(self, idx):
        return self.get_layer_by_id(idx, 'Team')

    def get_subabteilung_by_id(self, idx):
        return self.get_layer_by_id(idx, 'Sub-Abteilung')

    def get_layer_by_id(self, idx, layer):
        ret = f'{idx}'
        for name in self.master['id'][layer][idx]:
            ret = ret + " | " + name
        return ret

    def rec_add_to(self, name, res, vglst):
        res.append(name)
        for subname in vglst[name]:
            self.rec_add_to(subname, res, vglst)
        return res

    def write_df_sheet(self, df, col_idx, writer, sheet, filter):

        # Write dataframe to the xlsx sheet
        # --------------------------------------------------------------

        # add data frame to sheet
        df.to_excel(writer, sheet)

        last_row = df.shape[0] - 1
        last_col = df.shape[1] - 1

        # define and set number formats
        workbook  = writer.book
        worksheet = writer.sheets[sheet]

        #https://xlsxwriter.readthedocs.io/worksheet.html#autofilter
        worksheet.autofilter(0, 0, last_row, last_col)

        text_format = workbook.add_format()
        text_format.set_text_wrap()
        part_format = workbook.add_format({'num_format': '0%'})
        floa_format = workbook.add_format({'num_format': '#,##0.00'})

        """
        TODO FIXME
        worksheet.set_column(col_idx.index(f'{filter}-%') + 1,
                    col_idx.index(f'{filter}-%') + 1,
                    width = 10, cell_format = part_format)

        worksheet.set_column(col_idx.index(f'{filter}-Mean') + 1,
                            col_idx.index(f'{filter}-Mean') + 1,
                            width = 10,
                            cell_format = floa_format
                            )
        """
        worksheet.set_column(col_idx.index('Motivation 1') + 1,
                            col_idx.index('Verbesserung 2') + 1,
                            width = 40,
                            cell_format = text_format
                            )

        colors = [  "#F8696B",
                    "#FBAA77",
                    "#FFEB84",
                    "#E9E583",
                    "#D3DF82",
                    "#BDD881",
                    "#A6D27F",
                    "#90CB7E",
                    "#7AC57D",
                    "#63BE7B"
                 ]

        for idx in range(0, len(colors)):
            idx_format = workbook.add_format({'bg_color': colors[idx]})
            worksheet.conditional_format(0,
                                     col_idx.index('Stimmungswert') + 1,
                                     last_row + 1,
                                     col_idx.index('Stimmungswert') + 1,
                                     {'type'    : 'cell',
                                      'criteria': "=",
                                      'value'   : idx + 1,
                                      'format'  : idx_format
                                     })


    def read_data_from_csv(self):

        df = pd.read_csv(os.path.join(self.path,self.filename))

        # Remove SurveyMonkey columns not required
        df = df.drop(columns=['respondent_id',
                            'collector_id',
                            'date_created',
                            'date_modified',
                            'ip_address',
                            ])

        # TODO remove unknown columns, to be checked
        df = df.drop(columns=['Name',
                            'Boss Name'])

        # Rename columns
        df = df.rename(columns={
            'Gebe bitte an, wie zufrieden du bist als Angestellte/r von digitec/Galaxus. Die Skala geht von 1 (schlechtester Wert) bis 10 (bester Wert).Indique ton degré de satisfaction en tant qu’employé(e) digitec/Galaxus. L’échelle va de 1 (la moins bonne note) à 10 (la meilleure note).' : 'Stimmungswert',
            'Was motiviert dich an deinem Job besonders, was trägt besonders zu deiner Zufriedenheit bei?Qu’est-ce qui te motive spécialement dans ton travail, qu’est-ce qui contribue particulièrement à ta satisfaction ?' : 'Motivation 1',
            'Unnamed: 24': 'Motivation 2',
            'Was müsste man verbessern, damit du (noch) zufriedener wärst?Que devrait-on améliorer pour que tu sois (encore) plus satisfait(e)?': 'Verbesserung 1',
            'Unnamed: 26': 'Verbesserung 2',
            'email_address': 'EMAIL',
            'custom_1' : 'VGFIRST',
            'custom_2' : 'VGLAST',
            'custom_3' : 'Abteilung',
            'custom_4' : 'Team',
            'custom_5' : 'Gruppe'
            })

        return df


    def individual_report(self, df, vg, filters):

        res = []
        res = self.rec_add_to(vg, res, self.master['tree'])

        dfl = df.copy(deep=True)
        dfl = dfl.loc[ dfl['Vorgesetzter'].isin(res)]

        # open the XLSX writer
        filename = os.path.join(self.path, self.master['filenames'][vg])
        writer = pd.ExcelWriter(filename, engine='xlsxwriter')

        try:
            if self.master['span'][vg]['leader'] == 0 and self.master['span'][vg]['staff'] < 3:
                print(f'<<< remove {vg}')
                return
        except:
            print("span check not supported")

        for filter in filters:

            sheet  = filter

            dfc= dfl.copy(deep=True)

            col_idx = [
                filter,
                'Stimmungswert',
                f'{filter}-Mean',
                f'{filter}-Count',
                f'{filter}-Max',
                f'{filter}-%',
                'Motivation 1',
                'Motivation 2',
                'Verbesserung 1',
                'Verbesserung 2',
                f'{filter}-2019-01-01'
                ]

            dfc = dfl.reindex(col_idx, axis=1)

            # Remove all responses where its count is below minimum
            # --------------------------------------------------------------

            # Remove all rows where the possible max response is too low
            dfc = dfc.loc[dfc[f'{filter}-Max'] > self.options.min_nr_of_resp]

            # Remove all rows where the response cound is too low
            # independant of the management hierarchie
            # TODO dfc = dfc.loc[dfc[f'{filter}-Count'] > options.min_nr_of_resp]

            # Check the number of responses based on the current filter
            # ... this is special case when the filter removes some values
            # ... from the total count. Occures when the same group id
            # ... is used over different management levels

            if not dfc.empty:
                # if not already empty, check all single reports

                # get the counts grouped by the column with the name of
                # the filter and add it to the new column Filter-Count
                c = dfc.groupby(filter).count()['Stimmungswert']
                dfc['Filter-Count'] = df[filter]
                dfc = dfc.replace({'Filter-Count' : c.to_dict()})
                # Drop all rows where the Filter-Count is below the min
                # dfc = dfc.loc[dfc['Filter-Count'] > options.min_nr_of_resp]

            # Do not add empty sheets
            if dfc.empty:
                return

            # Switch back to full name
            # --------------------------------------------------------------

            function_name = "get_%s_by_id" % filter.lower().replace('-', '')
            function = getattr(self, function_name)
            dfc[filter] = dfc[filter].apply(function)

            # Simplify column names and sort
            # --------------------------------------------------------------
            dfc = dfc.rename({
                f'{filter}-%'    : 'Beteiligung',
                f'{filter}-Mean' : 'Mittelwert',
                f'{filter}-Count': 'Anzahl',
                f'{filter}-Max'  : 'Max',
                f'{filter}-2019.01.01' : '2019.01.01',
                             }, axis='columns')
            if 'Filter-Count' in dfc.columns:
                dfc = dfc.drop(columns=['Filter-Count'])
                pass
            dfc = dfc.sort_values([filter, 'Stimmungswert'], ascending=[True, False])

            # Simplify the Report
            # --------------------------------------------------------------
            # changed for report 2019 Q2

            if filter == "Gruppe":
                self.write_df_sheet(dfc, col_idx, writer, "Rückmeldungen", filter)
                return

            dfc = dfc.drop(columns=['Motivation 1',
                                    'Motivation 2',
                                    'Verbesserung 1',
                                    'Verbesserung 2',
                                    'Stimmungswert'])

            dfc = dfc.drop_duplicates()

            # Write dataframe to the xlsx sheet
            # --------------------------------------------------------------
            self.write_df_sheet(dfc, col_idx, writer, sheet, filter)

        # final save
        writer.save()


    def read_survey(self):

        df = self.read_data_from_csv()

        filters = ['Vorgesetzter', 'Gruppe', 'Team', 'Sub-Abteilung', 'Abteilung']

        # Merge all values (unamed columns 14 to 22) into one
        for idx in range (14, 23):
            df['Stimmungswert'] = df['Stimmungswert'].fillna(df[f'Unnamed: {idx}'])

        # Drop the unnamed columns 14 to 22)
        df = df.drop(columns= ['Unnamed: %s' % x for x in range(14,23)])

        # Delte Row 2 some helper text
        df = df.drop(df.index[0])

        df['Stimmungswert'] = df['Stimmungswert'].astype(int)
        df['Vorgesetzter']  = df.apply(lambda row: row['VGLAST'] + ", " + row['VGFIRST'], axis=1)
        df['Mitarbeiter']   = df.apply(lambda row: row['last_name'] + ", " + row['first_name'], axis=1)

        # NOTE Workaround, we do nhat have the sub-abt information
        abt = pd.read_json('collector-abt.json')
        df = pd.merge(df, abt, on='Mitarbeiter')

        # Calculate all filters

        for name in filters:

            # Reduce all columnes used as layer to the ID
            df[name] = df[name].map(self.get_id)

            df[f'{name}-Mean'] = df[name]
            df[f'{name}-Count'] = df[name]
            df[f'{name}-Max'] = df[name]
            df[f'{name}-2019-01-01'] = df[name]

            mean  = df.groupby(name).mean(numeric_only=True).to_dict()['Stimmungswert']
            count = df.groupby(name).count()['Stimmungswert'].to_dict()

            # create dictionary for history
            self.history[name] = {}
            self.history[name] = mean

            #df = df.replace({f'2019-01-01' : mean})

            df[f'{name}-2019-01-01'] = df[f'{name}-2019-01-01'].map(self.history_2[name])

            df = df.replace({f'{name}-Mean' : mean})
            df = df.replace({f'{name}-Count': count})
            df = df.replace({f'{name}-Max'  : self.master['counts'][name]})

            try:
                df[f'{name}-%'] = df.loc[:,f'{name}-Count'].astype(int) / df.loc[:,f'{name}-Max'].astype(int)
            except:
                print("xxx")

        # Store the history date into an external file for later use
        filename = os.path.join(self.path, f"history-{self.date.strftime('%Y-%M-%d')}.json")
        with open(filename, "w") as historyfile:
            json.dump(self.history, historyfile, indent=4, sort_keys=True)


        self.individual_report(df, self.master["ceo"], filters)

        if options.ceo_only == True:
            sys.exit(1)

        for vg in self.master['tree']:

            if vg == "NaN" or vg == self.master["ceo"]:
                continue

            self.individual_report(df, vg, filters)




    def tokenzie(self, text):
        print("---------------")
        print(text)
        from nltk.tokenize import word_tokenize

        tokenized_word=word_tokenize("why is this")
        print(tokenized_word)

    def nlp(self):

        df = self.read_data_from_csv()

        df = df.drop(columns=['VGFIRST',
                            'EMAIL',
                            'VGLAST',
                            'Abteilung',
                            'Team',
                            'first_name',
                            'last_name',
                            'Unnamed: 14', 'Unnamed: 15', 'Unnamed: 16',
                            'Unnamed: 17', 'Unnamed: 18', 'Unnamed: 19', 'Unnamed: 20',
                            'Unnamed: 21', 'Unnamed: 22',
                            'Motivation 1', 'Motivation 2'
                            ])

        # Reduce all columnes used as layer to the ID
        #df['Gruppe'] = df['Gruppe'].map(self.get_id)

        df['Verbesserung 1'].map(self.tokenzie)

        print(df.columns)


        # open the XLSX writer
        writer = pd.ExcelWriter('nlp.xlsx', engine='xlsxwriter')
        df.to_excel(writer, 'x')
        writer.save()

if __name__ == "__main__":

    parser = OptionParser()

    parser.add_option("-v", "--verbose",
        action="store_true", dest="verbose", default="False",
        help="Verbose prints enabled")

    parser.add_option("-c", "--ceo-only",
        action="store_true", dest="ceo_only", default="False",
        help="Only the CEO report")

    parser.add_option("-x", "--exclude",
        type="int", dest="min_nr_of_resp", default="3")

    parser.add_option("-d", "--date", dest="date",
                  help="date of the survey")

    parser.add_option("-f", "--file", dest="filename",
                  help="write report to FILE", metavar="FILE")

    parser.add_option("-m", "--mode", dest="mode",
                  help="vorgesetzter, abteilung, team, gruppe")

    (options, args) = parser.parse_args()

    x = Process(options)
    x.read_survey()
    #x.nlp()
