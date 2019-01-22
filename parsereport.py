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


def create_abt_tree(options):

    df = pd.read_csv(options.filename)

    keep = [
        'Nachname',
        'Vorname',
        'Sub-Abteilung',
        ]

    for column_name in df.columns.tolist():
        if column_name not in keep:
            df = df.drop(columns=[column_name])

    df = df.loc[df['Nachname'].notna()]
    df['Mitarbeiter']  = df.apply(lambda row: row['Nachname'] + ", " + row['Vorname'], axis=1)
    df = df.loc[df['Mitarbeiter'].notna()]

    df = df.drop(columns=['Nachname', 'Vorname'])
    print( df.columns.tolist())

    df.to_json('collector-abt.json')

def get_low_management_span(options):
    """list bottom leaders with low management span
    """

    df = pd.read_csv(options.filename)
    df = df.loc[df['Nachname'].notna()]
    df['Mitarbeiter']  = df.apply(lambda row: row['Nachname'] + ", " + row['Vorname'], axis=1)
    df = df.loc[df['Mitarbeiter'].notna()]
    vgl = df['Vorgesetzter'].unique().tolist()

    ret = {}

    # Go through all leaders
    for vg in vgl:

        # Remove empty values (might be CEO)
        if type(vg) != str:
            continue

        vn = 0
        fs = 0
        for idx, ma in df.loc[(df['Vorgesetzter'] == vg)].iterrows():
            if ma['Mitarbeiter'] in vgl:
                vn += 1
                fs += 1
            else:
                fs += 1

        if (vn == 0 and fs < 3):
            ret[vg] = fs

    pprint(ret)

def create_vg_tree(options):

    df = pd.read_csv(options.filename)

    keep = [
        'Nachname',
        'Vorname',
        'Vorgesetzter',
        'Abteilung',
        'Sub-Abteilung',
        'Team',
        'Gruppe'
        ]

    for column_name in df.columns.tolist():
        if column_name not in keep:
            df = df.drop(columns=[column_name])

    df = df.loc[df['Nachname'].notna()]
    df['Mitarbeiter']  = df.apply(lambda row: row['Nachname'] + ", " + row['Vorname'], axis=1)
    df = df.loc[df['Mitarbeiter'].notna()]

    tree = {}
    fname = {}

    vgl = df['Vorgesetzter'].unique().tolist()

    for vg in vgl:
        tree[vg] = []
        print(f'Vorgesetzter: {vg}')
        if type(vg) != str:
            # CEO
            continue

        ret  = df.loc[ (df['Mitarbeiter'] == vg) ]
        name = ""
        name += ret['Abteilung'].values[0].split('-')[0] + "-"
        name += ret['Sub-Abteilung'].values[0].split('-')[0] + "-"
        name += ret['Team'].values[0].split('-')[0] + "-"
        name += ret['Gruppe'].values[0].split('-')[0] + "-"
        name += vg.replace(', ', '-') + ".xlsx"
        print(name)
        fname[vg] = name

        vn = 0
        fs = 0

        mas = df.loc[ (df['Vorgesetzter'] == vg) ]
        for idx,ma in mas.iterrows():
            name = ma['Mitarbeiter']
            if name in vgl:
                tree[vg].append(name)
                print(f' - append: {name}')
                vn += 1
                fs += 1
            else:
                print(f' - not: {name}')
                fs += 1

        print (f'{vg} F.Spanne {fs} deept {vn}')

    with open('collector-vg-tree.json', 'w') as outfile:
        json.dump(tree, outfile)

    with open('collector-vg-fnames.json', 'w') as outfile:
        json.dump(fname, outfile)


if __name__ == "__main__":

    parser = OptionParser()

    parser.add_option("-v", "--verbose",
        action="store_true", dest="verbose", default="False",
        help="Verbose prints enabled")

    parser.add_option("-f", "--file", dest="filename",
                  help="write report to FILE", metavar="FILE")

    (options, args) = parser.parse_args()

    get_low_management_span(options)
    #create_vg_tree(options)
    #create_abt_tree(options)
