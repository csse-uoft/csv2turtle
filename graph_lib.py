#!/usr/bin/env python3
######################################################
# File: graph_lib.py
# Author: Bart Gajderowicz
# Date: May 6, 2022
# Description:
#   Helper functions for graphing netowrkx graphs
######################################################

import re, colorsys, random, os
import numpy as np
import pandas as pd
from misc_lib import *

import networkx as nx
from pylab import rcParams
from pyvis import network as pvnet



def clamp(rgb): 
    r,g,b = [max(0, min(round(255*x), 255)) for x in rgb]
    return "#{0:02x}{1:02x}{2:02x}".format(r,g,b)

def gen_color_tab(df, cols=['Funding']):
    random.seed(42)
    color_tab = {}
    vals = pd.concat([df[c] for c in cols]).dropna().drop_duplicates().sort_values()
    for val in vals:
        r,g,b = random.randint(200,255),random.randint(100,255),random.randint(100,255)
        r = 1.0
        rgb = colorsys.hsv_to_rgb(r/255,g/255,b/255)
        color_tab[val] = rgb

    return color_tab

def plot_g_pyviz(G, filename='out.html', dir='graphs', height='90%', width='100%', notebook=False, subtitle=None):
    _ = os.makedirs(dir) if not os.path.exists(dir) else None        

    g = G.copy() # some attributes added to nodes
    title = filename.split('.')[0].replace('_',' ').title()
    if subtitle is not None:
        title += "<br/><br/><i>"+subtitle+"</i>"
    net = pvnet.Network(notebook=notebook, directed=True, height=height, width=width, heading=title)
    opts = '''
        var options = {
          "physics": {
            "enabled": true,
            "forceAtlas2Based": {
              "gravitationalConstant": -100,
              "centralGravity": 0.11,
              "springLength": 100,
              "springConstant": 0.09,
              "avoidOverlap": 1
            },
            "minVelocity": 0.05,
            "solver": "forceAtlas2Based",
            "timestep": 0.22
          }
        }
    '''
    opts = '''
        var options = {
          "physics": {
            "enabled": true
          },
            "minVelocity": 0.05,
            "timestep": 0.22,
            "forceAtlas2Based": {
              "gravitationalConstant": -100,
              "centralGravity": 0.11,
              "springLength": 100,
              "springConstant": 0.09,
              "avoidOverlap": 1
            },
          "barnesHut":{
             "gravitationalConstant": 100,
              "avoidOverlap": 1
          },
          "solver":"barnesHut"
        }
    '''

    net.set_options(opts)
    # uncomment this to play with layout
    # pvnet.stabilize(2000)
    # net.show_buttons(filter_=['physics'])
    net.from_nx(g)

    return net.show(dir+'/'+filename)

def update_graph(g):
    ranks = pd.DataFrame([nx.pagerank(g, alpha=0.9)]).T.reset_index(drop=False)
    # ranks = pd.DataFrame([nx.betweenness_centrality(g)]).T.reset_index(drop=False) # <- cant't use for MultiDigraphs
    ranks[0] = ranks[0]/ranks[0].max()+0.5
    ranks = ranks.set_index('index')[0]

    for n in g.nodes():
        g.nodes[n]['size'] = (np.log1p(g.degree(n))+1)**1.7 * ranks[n]
        # print([n,str(n),ranks[n]])
        g.nodes[n]['size'] = 10*ranks[n]
        # g.nodes[n]['size'] = (np.log1p(g.degree(n))+1)**2.2
    for n in g.nodes:
        if re.search(r'^Org[0-9]+',n):
            g.nodes[n]['group'] = 1
        elif re.search(r'^P[0-9]+',n):
            g.nodes[n]['group'] = 2
        elif re.search(r'^S[0-9]+',n):
            g.nodes[n]['group'] = 3
        elif re.search(r'^CL\-',n):
            g.nodes[n]['group'] = 4
        else:
            g.nodes[n]['group'] = 99
    return g
def strip_namespace(df):
    for c in df.columns:
        df[c] = df[c].apply(lambda node: node.split('#')[-1] if type(node) is str else node)
        df[c] = df[c].apply(lambda node: node.split('_Location')[0] if type(node) is str else node)
    return df
