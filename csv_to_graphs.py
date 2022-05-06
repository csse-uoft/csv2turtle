#!/usr/bin/env python3
######################################################
# File: csv_to_graph.py
# Author: Bart Gajderowicz
# Date: May 6, 2022
# Description:
#   create visual graphs from csv/excel test data files
######################################################

import os
from misc_lib import *
from graph_lib import *

import numpy as np
import pandas as pd

import networkx as nx


xls = pd.ExcelFile('csv/unit_tests3.xlsx')

funding = pd.read_excel(xls,'Funding', header=1)
funding = funding.dropna(how='all')

services = pd.read_excel(xls,'Services', header=1)
services = services.dropna(how='all')


programs = pd.read_excel(xls,'Programs', header=1)
programs = programs.dropna(how='all')

models = pd.read_excel(xls,'LogicModels', header=1)
models = models.dropna(how='all')

orgs = pd.read_excel(xls,'Organizations', header=1)
orgs = orgs.dropna(how='all')


df = funding[['Funding','receivedFrom','fundersProgram','receivedAmount','forProgram']].merge(programs[['Program', 'hasService']], left_on='forProgram', right_on=['Program']).\
  merge(services[['Service','hasRequirement','hasCode']], left_on=['hasService'], right_on=['Service']).\
  merge(models[['LogicModel','forOrganization', 'hasProgram']], left_on=['fundersProgram'], right_on=['hasProgram']).\
  merge(models[['LogicModel','forOrganization', 'hasProgram']], left_on=['Program'], right_on=['hasProgram'])

df = df.rename(columns={
  'LogicModel_x':'fromLogicModel',
  'LogicModel_y':'toLogicModel',
  'receivedFrom':'fromOrg',
  'forOrganization_y':'toOrg',
  'hasProgram_x':'fromProgram',
  'hasProgram_y':'toProgram',
  'hasCode':'serviceCode',
})
funding_df = df[['Funding','fromLogicModel','toLogicModel','fromOrg','toOrg','fromProgram','toProgram','receivedAmount', 'Service','serviceCode']].\
    drop_duplicates()


service_df = []
for (org,ser),grp in df[df.serviceCode!='CL-Funding'][['toOrg','serviceCode','hasRequirement']].drop_duplicates().groupby(['toOrg','serviceCode']):
  codes = flatten([cs.replace('cids:hasCode CL-','').split(',') for cs in grp.hasRequirement.values])
  codes = [c.strip() for c in codes]
  for code in codes:
    service_df.append([org,ser,code])
service_df = pd.DataFrame(service_df,columns=['toOrg','serviceCode', 'clientCode'])

color_tab = gen_color_tab(pd.concat([funding_df, service_df]), cols=['Funding','toOrg'])
G0 = nx.MultiDiGraph()
_=[G0.add_edge(y,x,weight=w, title=w, color=clamp(color_tab[f])) for [f,x,y,w] in funding_df[['Funding','fromOrg','toOrg','receivedAmount']].values]
_=[G0.add_edge(y,x) for [x,y] in service_df[['toOrg','serviceCode']].values]

G0 = update_graph(G0)
plot_g_pyviz(G0, filename='by_service_codes.html')

