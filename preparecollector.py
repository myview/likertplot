#!/usr/bin/python3

from pprint import pprint
from optparse import OptionParser
import pandas as pd
import json
import sys


groups = {}

gpt = {}

def drop_unwanted_columns(df, keep=[]):

    # The SurveyMonkey report might have unwanted rows
    for column_name in df.columns.tolist():
        if column_name in keep:
            continue
        df = df.drop(columns=[column_name])
    return df

def gen_groups(name):
    id_nr   = name.split('-', 1)[0]
    id_name = name.split('-', 1)[1]
    if id_nr not in groups:
        groups[id_nr] = [id_name]
        return name
    elif id_name not in groups[id_nr]:
        groups[id_nr].append(id_name)

def name_split(name):
    n_short = name.split('-', 1)[0]
    n_name  = name.split('-', 1)[1]
    return f'{n_short} | {n_name}'

def rename_groups(name):
    id_nr   = name.split('-', 1)[0]
    id_name = name.split('-', 1)[1]
    ret = f'{id_nr}'
    for n in groups[id_nr]:
        ret = ret + " | " + n
    return ret

def create_collector(options):

    mitarbeiter = pd.read_csv(options.filename)

    mitarbeiter = drop_unwanted_columns(mitarbeiter, [
        'Geschäftlich  Informationen zur E-Mail E-Mail-Adresse',
        'Abteilung',
        'Team',
        'Gruppe',
        'Nachname',
        'Vorname',
        'Vorgesetzter',
        'Mitarbeitergruppe'])


    print("Remove:")
    print(mitarbeiter.loc[mitarbeiter['Mitarbeitergruppe'].isna()])
    print(mitarbeiter.loc[mitarbeiter.Mitarbeitergruppe == 'Lernende'])

    mitarbeiter = mitarbeiter.loc[mitarbeiter.Mitarbeitergruppe != 'Lernende']
    mitarbeiter = mitarbeiter.loc[mitarbeiter['Mitarbeitergruppe'].notna()]

    mitarbeiter = mitarbeiter.drop(columns=['Mitarbeitergruppe'])

    vg = mitarbeiter['Vorgesetzter'].str.split(", ", n = 1, expand = True)
    mitarbeiter["VG-LAST"] = vg[0]
    mitarbeiter["VG-FIRST"]= vg[1]

    for leader in mitarbeiter['Vorgesetzter'].unique():
        leader_group = mitarbeiter.loc[(mitarbeiter['Vorgesetzter'] == leader) ]
        groups = leader_group['Gruppe'].unique()
        if len(groups) > 1:
            print(leader_group['Vorgesetzter'].unique(), leader_group['Gruppe'].unique())


    mitarbeiter = mitarbeiter.rename(columns={
        'Nachname' : 'LAST',
        'Vorname'  : 'FIRST',
        'Geschäftlich  Informationen zur E-Mail E-Mail-Adresse' : 'EMAIL',
        'Abteilung' : 'ABTEILUNG',
        'Team' : 'TEAM',
        'Gruppe' : 'GRUPPE',
        })

    mitarbeiter = mitarbeiter.reindex_axis([
        'EMAIL',
        'FIRST',
        'LAST',
        'VG-FIRST',
        'VG-LAST',
        'ABTEILUNG',
        'TEAM',
        'GRUPPE',
        ], axis=1)

    mitarbeiter['GRUPPE'].map(gen_groups)
    mitarbeiter['GRUPPE'] = mitarbeiter['GRUPPE'].map(rename_groups)
    mitarbeiter['ABTEILUNG'] = mitarbeiter['ABTEILUNG'].map(name_split)
    mitarbeiter['TEAM'] = mitarbeiter['TEAM'].map(name_split)

    mitarbeiter.to_csv('collector-data.csv', index=False)

if __name__ == "__main__":

    parser = OptionParser()

    parser.add_option("-v", "--verbose",
        action="store_true", dest="verbose", default="False",
        help="Verbose prints enabled")

    parser.add_option("-f", "--file", dest="filename",
                  help="write report to FILE", metavar="FILE")

    (options, args) = parser.parse_args()

    create_collector(options)
