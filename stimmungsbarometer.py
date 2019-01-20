#!/usr/bin/python3

from pprint import pprint
from optparse import OptionParser
import pandas as pd
import json
import sys




def read_survey(options):


    df = pd.read_csv(options.filename)

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


    # Merge all values into one
    # TODO could we fix this in SurveyMonkey direct?
    df['Stimmungswert'] = df['Stimmungswert'].fillna(df['Unnamed: 14'])
    df['Stimmungswert'] = df['Stimmungswert'].fillna(df['Unnamed: 15'])
    df['Stimmungswert'] = df['Stimmungswert'].fillna(df['Unnamed: 16'])
    df['Stimmungswert'] = df['Stimmungswert'].fillna(df['Unnamed: 17'])
    df['Stimmungswert'] = df['Stimmungswert'].fillna(df['Unnamed: 18'])
    df['Stimmungswert'] = df['Stimmungswert'].fillna(df['Unnamed: 19'])
    df['Stimmungswert'] = df['Stimmungswert'].fillna(df['Unnamed: 20'])
    df['Stimmungswert'] = df['Stimmungswert'].fillna(df['Unnamed: 21'])
    df['Stimmungswert'] = df['Stimmungswert'].fillna(df['Unnamed: 22'])
    df = df.drop(columns=['Unnamed: 14',
                          'Unnamed: 15',
                          'Unnamed: 16',
                          'Unnamed: 17',
                          'Unnamed: 18',
                          'Unnamed: 19',
                          'Unnamed: 20',
                          'Unnamed: 21',
                          'Unnamed: 22',
                          ])

    # Delte Row 2 some helper text
    df = df.drop(df.index[0])

    df['Stimmungswert'] = df['Stimmungswert'].astype(int)


    #df['Vorgesetzter'] = df.apply(lambda row: row['VGFIRST'] + " " + row['VGLAST'], axis=1)
    #df['Mitarbeiter']  = df.apply(lambda row: row['first_name'] + " " + row['last_name'], axis=1)

    df['Vorgesetzter'] = df.apply(lambda row: row['VGLAST'] + ", " + row['VGFIRST'], axis=1)
    df['Mitarbeiter']  = df.apply(lambda row: row['last_name'] + ", " + row['first_name'], axis=1)

    v1 = df.groupby('Vorgesetzter').mean(numeric_only=True)
    v2 = df.groupby('Gruppe').mean(numeric_only=True)
    v3 = df.groupby('Team').mean(numeric_only=True)
    v4 = df.groupby('Abteilung').mean(numeric_only=True)

    #print(v4)
    #print(v4.to_dict())

    c1 = df.groupby('Vorgesetzter').count()['Stimmungswert']
    c2 = df.groupby('Gruppe').count()['Stimmungswert']
    c3 = df.groupby('Team').count()['Stimmungswert']
    c4 = df.groupby('Abteilung').count()['Stimmungswert']

    #print(df.groupby('Vorgesetzter').count().to_dict()['Stimmungswert'])

    #print(v.to_dict()['Stimmungswert'])
    #print(v.to_dict()['Stimmungswert']['Adelisa Kapic'])

    df['Gruppe-Mean'] = df['Gruppe']
    df['Gruppe-Count'] = df['Gruppe']

    df['Vorgesetzter-Mean'] = df['Vorgesetzter']
    df['Vorgesetzter-Count'] = df['Vorgesetzter']

    df['Team-Mean'] = df['Team']
    df['Team-Count'] = df['Team']

    df['Abteilung-Mean'] = df['Abteilung']
    df['Abteilung-Count'] = df['Abteilung']

    df = df.replace({'Vorgesetzter-Mean'  : v1.to_dict()['Stimmungswert']})
    df = df.replace({'Vorgesetzter-Count' : c1.to_dict()})

    df = df.replace({'Gruppe-Mean'  : v2.to_dict()['Stimmungswert']})
    df = df.replace({'Gruppe-Count' : c2.to_dict()})

    df = df.replace({'Team-Mean'  : v3.to_dict()['Stimmungswert']})
    df = df.replace({'Team-Count'  : c3.to_dict()})

    df = df.replace({'Abteilung-Mean'  : v4.to_dict()['Stimmungswert']})
    df = df.replace({'Abteilung-Count'  : c4.to_dict()})

    #CEO = 'Florian Teuteberg'
    #tree = {CEO}
    #res = df.loc[df['Vorgesetzter'] == CEO]
    #for ma in res.Mitarbeiter:
        #res2 = df.loc[df['Vorgesetzter'] == ma]
        ##print (res2)


    def rec_add_to(name, res, vglst):
        res.append(name)
        for subname in vglst[name]:
            rec_add_to(subname, res, vglst)
        return res

    with open('vg.json') as json_file:
        vgl = json.load(json_file)
        for vg in vgl:
            print(f"# start new list with: {vg}")
            res = []
            res = rec_add_to(vg, res, vgl)

            #print(res)
            #sys.exit(0)
            #print(type(res))
            dfl = df.copy(deep=True)
            dfl = dfl.loc[ dfl['Vorgesetzter'].isin(res)]

            #print(dfl)

            """
            Datum
            Gruppe
            Team
            Abteilung
            Stimmungswert
            Motivation 1
            Motivation 2
            Verbesserung 1
            Verbesserung 2

            """
            #sys.exit(0)

            # open the XLSX writer
            writer = pd.ExcelWriter(vg + 'output.xlsx', engine='xlsxwriter')


            x = ['Vorgesetzter', 'Gruppe', 'Team', 'Abteilung']

            for mode in x:
                print(f'MODE: {mode} - with min number of respones: {options.min_nr_of_resp}')

                sheet  = mode

                dfc= dfl.copy(deep=True)

                dfc = dfl.reindex([
                    mode,
                    'Stimmungswert',

                    f'{mode}-Mean',
                    f'{mode}-Count',
                    'Motivation 1',
                    'Motivation 2',
                    'Verbesserung 1',
                    'Verbesserung 2'

                    ], axis=1)

                dfc = dfc.loc[dfc[f'{mode}-Count'] > options.min_nr_of_resp]
                #dfnok = df.loc[df[f'{mode}-Count'] <= options.min_nr_of_resp]
                ##print(dfnok)

                # add data frame to sheet
                dfc.to_excel(writer, sheet)

                last_row = dfc.shape[0] - 1
                last_col = dfc.shape[1] - 1

                # define and set number formats
                workbook  = writer.book
                worksheet = writer.sheets[sheet]

                #https://xlsxwriter.readthedocs.io/worksheet.html#autofilter
                worksheet.autofilter(0, 0, last_row, last_col)

                # final save
            writer.save()

    """
    else:
        df = df.reindex([
            'EMAIL',
            'Vorgesetzter',
            'Vorgesetzter-Mean',
            'Vorgesetzter-Count',
            'Gruppe',
            'Gruppe-Mean',
            'Gruppe-Count',
            'Team',
            'Team-Mean',
            'Team-Count',
            'Abteilung',
            'Abteilung-Mean',
            'Abteilung-Count',
            ], axis=1)
    """

    df.to_csv('output.csv', index=True) #, float_format = '%0.00f')




if __name__ == "__main__":

    parser = OptionParser()

    parser.add_option("-v", "--verbose",
        action="store_true", dest="verbose", default="False",
        help="Verbose prints enabled")

    parser.add_option("-x", "--exclude",
        type="int", dest="min_nr_of_resp", default="3")

    parser.add_option("-f", "--file", dest="filename",
                  help="write report to FILE", metavar="FILE")

    parser.add_option("-m", "--mode", dest="mode",
                  help="vorgesetzter, abteilung, team, gruppe")

    (options, args) = parser.parse_args()

    read_survey(options)
