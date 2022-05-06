#!/usr/bin/env python3
######################################################
# File: sparql_to_graph.py
# Author: Bart Gajderowicz
# Date: May 6, 2022
# Description:
#   create visual graphs from sparql queries
######################################################
import os
from misc_lib import *

import numpy as np
import pandas as pd

import xlsxwriter

import networkx as nx

from graph_lib import *


df = pd.read_csv('csv/query-result-cq1-sh_Homeless_Female_Youth_Area0.csv')
df = strip_namespace(df)
G0 = nx.MultiDiGraph()
_=[G0.add_edge(org,serv) for [org,serv] in df[['forOrg','forProgram']].drop_duplicates().values]
_=[G0.add_edge(serv, cd) for [serv,cd] in df[['forProgram','forService']].drop_duplicates().values]
_=[G0.add_edge(serv, cd) for [serv,cd] in df[['forService','forCode']].drop_duplicates().values]
_=[G0.add_edge(cd,"%s @ %s"%(reqcd,loc)) for [cd,reqcd,loc] in df[['forCode', 'reqCode', 'location']].drop_duplicates().values]

G0 = update_graph(G0)
plot_g_pyviz(G0, filename='organizations_by_service_for_Homeless_Female_Youth_in_Area0.html', subtitle="What cp:Organization (s) by cp:Service type exist in my cp:Community")



df = pd.read_csv('csv/query-result-cq1.csv')
df = strip_namespace(df)
G1 = nx.MultiDiGraph()
_=[G1.add_edge(y,x) for [x,y] in df[['forOrg','reqCode']].drop_duplicates().values]
# _=[G0.add_edge(x,y) for [x,y] in df[~df['StakeholderType'].str.contains('CL-Funder')][['StakeholderType','forOrg']].drop_duplicates().values]
_=[G1.add_edge(y,"%s @ %s"%(x,a)) for [x,y,a] in df[['location','forOrg','forService']].drop_duplicates().values]

G1 = update_graph(G1)
plot_g_pyviz(G1, filename='organization_by_community_service_codes.html', subtitle="What cp:Organizations service a community?")



df = pd.read_csv('csv/query-result-cq2.csv')
df = strip_namespace(df)
G1 = nx.MultiDiGraph()
_=[G1.add_edge(y,x) for [x,y] in df[['forOrg','fundersOrg']].drop_duplicates().values]
# _=[G0.add_edge(x,y) for [x,y] in df[~df['StakeholderType'].str.contains('CL-Funder')][['StakeholderType','forOrg']].drop_duplicates().values]
_=[G1.add_edge(y,"%s @ %s"%(x,a)) for [x,y,a] in df[df['forCode'] != 'CL-Funding'][['location','forOrg', 'forCode']].drop_duplicates().values]

G1 = update_graph(G1)
plot_g_pyviz(G1, filename='organization_funding_by_community_characteristic_codes.html', subtitle="What are the funding flows for the organizations broken down by cp:Stakeholder type?")

