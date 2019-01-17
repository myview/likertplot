#!/usr/bin/python3

from pprint import pprint
from optparse import OptionParser
import pandas as pd
import json





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
        'Gebe bitte an, wie zufrieden du bist als Angestellte/r von digitec/Galaxus. Die Skala geht von 1 (schlechtester Wert) bis 10 (bester Wert).Indique ton degré de satisfaction en tant qu’employé(e) digitec/Galaxus. L’échelle va de 1 (la moins bonne note) à 10 (la meilleure note).' : 'VALUE',
        'Was motiviert dich an deinem Job besonders, was trägt besonders zu deiner Zufriedenheit bei?Qu’est-ce qui te motive spécialement dans ton travail, qu’est-ce qui contribue particulièrement à ta satisfaction ?' : 'GOOD 1',
        'Unnamed: 24': 'GOOD 2',
        'Was müsste man verbessern, damit du (noch) zufriedener wärst?Que devrait-on améliorer pour que tu sois (encore) plus satisfait(e)?': 'BAD 1',
        'Unnamed: 26': 'BAD 2',
        'email_address': 'EMAIL',
        'first_name': 'FIRST',
        'last_name': 'LAST',
        'custom_1' : 'VGFIRST',
        'custom_2' : 'VGLAST',
        'custom_3' : 'ABTEILUNG',
        'custom_4' : 'TEAM',
        'custom_5' : 'GRUPPE'
        })


    # Merge all values into one
    # TODO could we fix this in SurveyMonkey direct?
    df['VALUE'] = df['VALUE'].fillna(df['Unnamed: 14'])
    df['VALUE'] = df['VALUE'].fillna(df['Unnamed: 15'])
    df['VALUE'] = df['VALUE'].fillna(df['Unnamed: 16'])
    df['VALUE'] = df['VALUE'].fillna(df['Unnamed: 17'])
    df['VALUE'] = df['VALUE'].fillna(df['Unnamed: 18'])
    df['VALUE'] = df['VALUE'].fillna(df['Unnamed: 19'])
    df['VALUE'] = df['VALUE'].fillna(df['Unnamed: 20'])
    df['VALUE'] = df['VALUE'].fillna(df['Unnamed: 21'])
    df['VALUE'] = df['VALUE'].fillna(df['Unnamed: 22'])
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

    df['VALUE'] = df['VALUE'].astype(int)


    df['Vorgesetzter'] = df.apply(lambda row: row['VGFIRST'] + " " + row['VGLAST'], axis=1)

    v1 = df.groupby('Vorgesetzter').mean(numeric_only=True)
    v2 = df.groupby('GRUPPE').mean(numeric_only=True)
    v3 = df.groupby('TEAM').mean(numeric_only=True)
    v4 = df.groupby('ABTEILUNG').mean(numeric_only=True)

    print(v4)
    print(v4.to_dict())

    c1 = df.groupby('Vorgesetzter').count()['VALUE']
    c2 = df.groupby('GRUPPE').count()['VALUE']
    c3 = df.groupby('TEAM').count()['VALUE']
    c4 = df.groupby('ABTEILUNG').count()['VALUE']

    #print(df.groupby('Vorgesetzter').count().to_dict()['VALUE'])

    #print(v.to_dict()['VALUE'])
    #print(v.to_dict()['VALUE']['Adelisa Kapic'])

    df['Gruppe-Mean'] = df['GRUPPE']
    df['Gruppe-Count'] = df['GRUPPE']

    df['Vorgesetzter-Mean'] = df['Vorgesetzter']
    df['Vorgesetzter-Count'] = df['Vorgesetzter']

    df['Team-Mean'] = df['TEAM']
    df['Team-Count'] = df['TEAM']

    df['Abteilung-Mean'] = df['ABTEILUNG']
    df['Abteilung-Count'] = df['ABTEILUNG']

    #df['Mean-Value'].apply(v.to_dict()['VALUE'])
    df = df.replace({'Vorgesetzter-Mean'  : v1.to_dict()['VALUE']})
    df = df.replace({'Vorgesetzter-Count' : c1.to_dict()})

    df = df.replace({'Gruppe-Mean'  : v2.to_dict()['VALUE']})
    df = df.replace({'Gruppe-Count' : c2.to_dict()})

    #print(v3.to_dict()['VALUE'])

    df = df.replace({'Team-Mean'  : v3.to_dict()['VALUE']})
    df = df.replace({'Team-Count'  : c3.to_dict()})

    df = df.replace({'Abteilung-Mean'  : v4.to_dict()['VALUE']})
    df = df.replace({'Abteilung-Count'  : c4.to_dict()})

    print(df['Abteilung-Mean'])


    df = df.reindex([
        'EMAIL',
        'Vorgesetzter',
        'Vorgesetzter-Mean',
        'Vorgesetzter-Count',
        'GRUPPE',
        'Gruppe-Mean',
        'Gruppe-Count',
        'TEAM',
        'Team-Mean',
        'Team-Count',
        'ABTEILUNG',
        'Abteilung-Mean',
        'Abteilung-Count',
        ], axis=1)


    df.to_csv('output.csv', index=True) #, float_format = '%0.00f')




if __name__ == "__main__":

    parser = OptionParser()

    parser.add_option("-v", "--verbose",
        action="store_true", dest="verbose", default="False",
        help="Verbose prints enabled")

    parser.add_option("-f", "--file", dest="filename",
                  help="write report to FILE", metavar="FILE")

    (options, args) = parser.parse_args()

    read_survey(options)
