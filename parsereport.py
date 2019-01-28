#!/usr/bin/python3

from pprint import pprint
from optparse import OptionParser
import pandas as pd
import json
import sys
import re

class Process:

    def __init__(self, options):

        # Configure the layers
        self.layers = ['Abteilung', 'Sub-Abteilung', 'Team', 'Gruppe']

        # Create an empty data master structure
        self.master = {}
        self.master['counts'] = {
            'Vorgesetzter'  : {},
            'Abteilung'     : {},
            'Sub-Abteilung' : {},
            'Team'          : {},
            'Gruppe'        : {} }
        self.master['tree'] = {}
        self.master['filenames'] = {}
        self.master['id'] = {}
        for layer in self.layers:
            self.master['id'][layer] = {}
        self.master['leaders'] = []

        self.groups = {}
        self.gpt = {}
        self.create_collector(options)
        self.create_vg_tree(options)
        self.create_abt_tree(options)

    def _get_id(self, name):
        """
        Retruns the unique identifier of the name
        if no itentifier is found, the name is returned
            https://pythex.org
            https://tinyurl.com/yd6p4kbz
        """
        m = re.search('^([A-Z]{2,}|[0-9]{2,})', name)
        if m:
            return m.group(0)
        return name

    def _gruppe_id_to_master(self, name):
        self._id_to_master(name, 'Gruppe')

    def _team_id_to_master(self, name):
        self._id_to_master(name, 'Team')

    def _subabteilung_id_to_master(self, name):
        self._id_to_master(name, 'Sub-Abteilung')

    def _abteilung_id_to_master(self, name):
        self._id_to_master(name, 'Abteilung')

    def _id_to_master(self, name, layer):
        m = re.search('^([A-Z]{2,}|[0-9]{2,})-(.*)', name)
        i = m.group(1)
        t = m.group(2)
        if i not in self.master['id'][layer]:
            self.master['id'][layer][i] = []
        if t not in self.master['id'][layer][i]:
            self.master['id'][layer][i].append(t)

    def drop_unwanted_columns(self, df, keep=[]):

        # The SurveyMonkey report might have unwanted rows
        for column_name in df.columns.tolist():
            if column_name in keep:
                continue
            df = df.drop(columns=[column_name])
        return df

    def gen_groups(self, name):
        #print(name)
        id_nr   = name.split('-', 1)[0]
        id_name = name.split('-', 1)[1]
        if id_nr not in self.groups:
            self.groups[id_nr] = [id_name]
            return name
        elif id_name not in self.groups[id_nr]:
            self.groups[id_nr].append(id_name)

    def name_split(self, name):
        n_short = name.split('-', 1)[0]
        n_name  = name.split('-', 1)[1]
        return f'{n_short} | {n_name}'

    def rename_groups(self, name):
        id_nr   = name.split('-', 1)[0]
        id_name = name.split('-', 1)[1]
        ret = f'{id_nr}'
        for n in self.groups[id_nr]:
            ret = ret + " | " + n
        return ret

    def create_abt_tree(self, options):

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
        #print( df.columns.tolist())

        df.to_json('collector-abt.json')

    def get_low_management_span(self, options):
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

    def create_vg_tree(self, options):

        df = pd.read_csv(options.filename)
        df = df.loc[df.Mitarbeitergruppe != 'Lernende']
        df = df.loc[df['Mitarbeitergruppe'].notna()]

        cols = [
            'Nachname',
            'Vorname',
            'Vorgesetzter',
            'Abteilung',
            'Sub-Abteilung',
            'Team',
            'Gruppe',
            ]

        for column_name in df.columns.tolist():
            if column_name not in cols:
                df = df.drop(columns=[column_name])

        df['Gruppe'].map(self._gruppe_id_to_master)
        df['Team'].map(self._team_id_to_master)
        df['Sub-Abteilung'].map(self._subabteilung_id_to_master)
        df['Abteilung'].map(self._abteilung_id_to_master)

        # Reduce all columnes used as layer to the ID
        for layer in self.layers:
            df[layer] = df[layer].map(self._get_id)

        # Get list of all leaders
        self.master['leaders'] = df['Vorgesetzter'].unique().tolist()

        # Count all Filters
        for col in self.master['counts']:
            for name, cnt in df.groupby(col).count()['Nachname'].iteritems():
                self.master['counts'][col][name.split(' | ')[0]] = cnt

        df['Mitarbeiter']  = df.apply(lambda row: row['Nachname'] + ", " + row['Vorname'], axis=1)
        df = df.loc[df['Mitarbeiter'].notna()]

        for vg in self.master['leaders']:
            self.master['tree'][vg] = []
            if type(vg) != str:
                # CEO
                continue

            ret  = df.loc[ (df['Mitarbeiter'] == vg) ]
            name = ""
            for layer in self.layers:
                name += ret[layer].values[0] + "-"
            name += vg.replace(', ', '-') + ".xlsx"
            self.master['filenames'][vg] = name.replace(' ', '-')

            vn = 0
            fs = 0
            mas = df.loc[ (df['Vorgesetzter'] == vg) ]
            for idx,ma in mas.iterrows():
                name = ma['Mitarbeiter']
                if name in self.master['leaders']:
                    self.master['tree'][vg].append(name)
                    vn += 1
                    fs += 1
                else:
                    fs += 1


    def write_master_to_json(self, name = 'master'):
        """
        Write the master data to a json file for post-processing
        """
        print(f'> write: {name}.json')
        with open(f'{name}.json', 'w') as outfile:
            json.dump(self.master, outfile)


    def create_collector(self, options):

        df = pd.read_csv(options.filename)

        df = self.drop_unwanted_columns(df, [
            'Geschäftlich  Informationen zur E-Mail E-Mail-Adresse',
            'Abteilung',
            'Team',
            'Gruppe',
            'Nachname',
            'Vorname',
            'Vorgesetzter',
            'Mitarbeitergruppe'])

        df = df.loc[df.Mitarbeitergruppe != 'Lernende']
        df = df.loc[df['Mitarbeitergruppe'].notna()]

        df = df.drop(columns=['Mitarbeitergruppe'])

        vg = df['Vorgesetzter'].str.split(", ", n = 1, expand = True)
        df["VG-LAST"] = vg[0]
        df["VG-FIRST"]= vg[1]

        for leader in df['Vorgesetzter'].unique():
            leader_group = df.loc[(df['Vorgesetzter'] == leader) ]
            groups = leader_group['Gruppe'].unique()
            #if len(groups) > 1:
                #print(leader_group['Vorgesetzter'].unique(), leader_group['Gruppe'].unique())


        df = df.rename(columns={
            'Nachname' : 'LAST',
            'Vorname'  : 'FIRST',
            'Geschäftlich  Informationen zur E-Mail E-Mail-Adresse' : 'EMAIL',
            'Abteilung' : 'ABTEILUNG',
            'Team' : 'TEAM',
            'Gruppe' : 'GRUPPE',
            })

        df = df.reindex_axis([
            'EMAIL',
            'FIRST',
            'LAST',
            'VG-FIRST',
            'VG-LAST',
            'ABTEILUNG',
            'TEAM',
            'GRUPPE',
            ], axis=1)

        df['GRUPPE'].map(self.gen_groups)
        df['GRUPPE'] = df['GRUPPE'].map(self.rename_groups)
        df['ABTEILUNG'] = df['ABTEILUNG'].map(self.name_split)
        df['TEAM'] = df['TEAM'].map(self.name_split)

        df.to_csv('collector-data.csv', index=False)

if __name__ == "__main__":

    parser = OptionParser()

    parser.add_option("-v", "--verbose",
        action="store_true", dest="verbose", default="False",
        help="Verbose prints enabled")

    parser.add_option("-f", "--file", dest="filename",
                  help="write report to FILE", metavar="FILE")

    (options, args) = parser.parse_args()

    run = Process(options)
    run.write_master_to_json()
