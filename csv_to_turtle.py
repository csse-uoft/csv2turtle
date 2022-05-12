#!/usr/bin/env python3
######################################################
# File: csv_to_turtle.py
# Author: Bart Gajderowicz
# Date: May 6, 2022
# Description:
#   create a turtle (.ttl) file from csv/excel data files
######################################################


import os
from datetime import datetime
from misc_lib import *

import numpy as np
import pandas as pd

import xlsxwriter

class CompositeCharacteristicException(Exception):
    """Raised when CompositeCharacteristic has less than 2 codes"""
    pass

class CharacterisrticParamException(Exception):
    """Raised when the incorrect format of characteristics is passed"""
    pass

# filein  = 'unit_tests_ieee.xlsx'
# dirin =  os.path.expanduser("~") + '/Dropbox/Compass Shared Folder/Use Cases/Competency Questions/IEEE Smart Cities 2022'
# fileout = 'unit_tests_ieee.ttl'
# dirout = os.path.expanduser("~") +'/Dropbox/Compass Shared Folder/Use Cases/Competency Questions/IEEE Smart Cities 2022'

# file_date = datetime.now().strftime("%B %d, %Y %H:%M:%S")
file_date = datetime.now().strftime("%B %Y Release")
filein  = 'unit_tests3.xlsx'
dirin = 'csv'
fileout = 'unit_test3.ttl'
dirout = 'turtle'

class_map = {
    "Organization":"Organization",
    "Funding":"cp:Funding",
    "Service":"cp:Service",
    "Client":"cp:Client",
    'Client Code':"cp:Code",
    'Service Code':"cp:Code",
    "Stakeholder":"cids:Stakeholder",
    "Program":'cp:Program',
    'ImpactModel':'cids:ImpactModel',
    'LogicModel':'cids:LogicModel',
    'Community':'cp:Community',
    'Code':'cids:Code',
    'Characteristic':'cids:Characteristic',
    'PrimitiveCharacteristic':'cids:PrimitiveCharacteristic',
    'CompositeCharacteristic':'cids:CompositeCharacteristic',
    'CommunityCharacteristic':'CommunityCharacteristic',
    'LandArea':'LandArea',
    'Feature':'loc_50871:Feature',
    'OrganizationID':'org:OrganizationID',
    'ServiceEvent':'ServiceEvent',
}
prop_map = {
    'hasLegalName':'org:hasLegalName',
    'receivedAmount':'cp:receivedAmount',
	'requestedAmount':'cp:requestedAmount',
    'fundersProgram':'cp:fundersProgram',
    'hasProgram':'cids:hasProgram',
    'forProgram':'cp:forProgram',
    'hasImpactModel': 'cids:hasImpactModel',
    'hasService':'cids:hasService',
    'hasCode':'cids:hasCode',
    'forStakeholder':'cids:forStakeholder',
    'locatedIn':'i72:located_in',
    'hasCommunityCharacteristic':'cp:hasCommunityCharacteristic',
    'hasCharacteristic':'cids:hasCharacteristic',
    'hasRequirement':'cp:hasRequirement',
    'hasFocus':'cp:hasFocus',
    'hasLocation':'landuse_50872:hasLocation',
    'satisfiesStakeholder':'cp:satisfiesStakeholder',
    'hasLandArea':'landuse_50872:hasLandArea',
    'hasIdentifier':'org:hasIdentifier',
    'hasStakeholder':'cids:hasStakeholder',
    'hasName':'hasName',
    'hasPart':'oep:hasPart',
    'hasDescription':'cids:hasDescription',
    'hasContributingStakeholder':'cids:hasContributingStakeholder',
    'hasBeneficialStakeholder':'cids:hasBeneficialStakeholder',
    'partOf':'oep:partOf',
    'hasMode':'hasMode',
    'forOrganization':'cids:forOrganization',
    'hasIndicator':'cids:hasIndicator',
    'hasOutcome':'cids:hasOutcome',
    'hasID':'org:hasID',
    'hasNumber':'hasNumber',
    'hasStatus':'hasStatus',
    'forClient':'forClient',
    'satisfiesStakeholder':'satisfiesStakeholder',
    'AtOrganization':'AtOrganization',
    'forReferral':'forReferral',
    'occursAt':'occursAt',
    'previousEvent':'previousEvent',
    'nextEvent':'nextEvent',
    'hasBeginning':'time:hasBeginning',
    'hasEnd':'time:hasEnd',
    'hasTemporalDuration':'time:hasTemporalDuration',
    'hasGender':'hasGender',
    'hasSex':'hasSex',
    'hasIncome':'hasIncome',
    'hasSkill':'hasSkill',
    'hasEthnicity':'hasEthnicity',
    'memberOfAboriginalGroup':'memberOfAboriginalGroup',
    'hasReligion':'hasReligion',
    'hasDependent':'hasDependent',
    'schema:knowsLanguage':'schema:knowsLanguage',
    'hasNeed':'hasNeed',
    'hasGoal':'hasGoal',
    'hasProblem':'hasProblem',
    'hasStatus':'hasStatus',
    'hasClientState':'hasClientState',


}


PREFIX = 'cp'
# @prefix geo:  <http://release.niem.gov/niem/adapters/geospatial/3.0#>.
# w3_org = default_world.get_namespace("http://www.w3.org/ns/org#")
# @prefix gml:  <http://www.opengis.net/gml/3.2#>.
# @prefix loc:  <http://ontology.eil.utoronto.ca/5087/1/SpatialLoc/>.


text = '''
################################################################
# Turtle File Generated by: csv_to_turtle.py
# Date : %s
# github: https://github.com/csse-uoft/csv2turtle
################################################################
'''%(file_date)
text += '''
@prefix act:  <http://ontology.eil.utoronto.ca/tove/activity#>.
@prefix dcat: <http://www.w3.org/ns/dcat#>.
@prefix ic:   <http://ontology.eil.utoronto.ca/tove/icontact#">.
@prefix owl:  <http://www.w3.org/2002/07/owl#>.

@prefix sur: <http://ontology.eil.utoronto.ca/tove/survey#>.
@prefix dqv: <http://www.w3.org/ns/dqv#>.
@prefix qb:  <http://purl.org/linked-data/cube#>.
@prefix dcterms: <http://purl.org/dc/terms#>.
@prefix loc_50871:  <http://ontology.eil.utoronto.ca/5087/1/SpatialLoc/>.
@prefix act_50871: <http://ontology.eil.utoronto.ca/5087/1/Activity#>.
@prefix city_50872: <http://ontology.eil.utoronto.ca/5087/2/City#>.
@prefix cityservice_50872: <http://ontology.eil.utoronto.ca/5087/2/CityService#>.
@prefix landuse_50872: <http://ontology.eil.utoronto.ca/5087/2/LandUse/>.

@prefix time: <http://www.w3.org/2006/time#>.
@prefix oep:  <http://www.w3.org/2001/sw/BestPractices/OEP/SimplePartWhole/part.owl#>.
@prefix xsd:  <http://www.w3.org/2001/XMLSchema#>.
@prefix rdf:  <http://www.w3.org/1999/02/22-rdf-syntax-ns#>.
@prefix rdfs: <http://www.w3.org/2000/01/rdf-schema#>.
@prefix foaf: <http://xmlns.com/foaf/0.1/> .
@prefix rel:  <http://purl.org/vocab/relationship/> .
@prefix geo:  <http://www.opengis.net/ont/geosparql#>.
@prefix i72:  <http://ontology.eil.utoronto.ca/ISO21972/iso21972#>.
@prefix org:  <http://ontology.eil.utoronto.ca/tove/organization#>.
@prefix cids: <http://ontology.eil.utoronto.ca/cids/cids#>.
@prefix cp:   <http://helpseeker.co/compass#> .
@prefix :     <http://helpseeker.co/compass#> .

<local> rdf:type owl:Ontology ;
    owl:imports <https://github.com/csse-uoft/compass-ontology/releases/download/latest/compass.owl> .

'''

def date_to_xsd(s):
    d = None
    if s is None or s != s:
        return s
    elif type(s) is  pd._libs.tslibs.timestamps.Timestamp:
        d = s
    elif type(s) is str:
        d = datetime.fromisoformat(s)
    return d.strftime("%Y-%m-%dT%H:%M:%S.000")
def entity_str(e,prefix=PREFIX):
    e=e.strip()
    return e if ':' in e else "%s:%s"%(prefix,e)
def format_lists(insts, entity=True):
    insts = flatten([s.split(',') for s in insts if s == s])
    insts = list(set([s.strip() for s in insts]))
    if entity:
        insts = [entity_str(s) for s in insts]
    return insts

def format_characteristics_text(inst, chars, prop0='hasCharacteristic'):
    text = ''
    if type(chars) == str:
        tmp = [c.strip() for c in chars.split(',')]
    elif type(chars) == list:
        tmp = chars
    elif pd.isnull(chars):
        return text
    else:
        raise CharacterisrticParamException("chars parameter must be a str or a list \n\tpassed (%s)\n\tfor (%s)"%(chars, inst))

    for code_text in tmp:
        if code_text != code_text:
            return ''
        elif code_text.startswith('cids:hasCode '):
            code_text = code_text.replace('cids:hasCode ','')
        prop = entity_str(prop_map['hasCode'])
        text1 = None
        if code_text.startswith('Comp-INST-'):
            # providing list of composite char labels
            code_text = entity_str(code_text)
            prop1 = entity_str(prop_map[prop0])
            text += "%s %s %s.\n"%(inst,prop1,code_text)
            klass = entity_str(class_map['CompositeCharacteristic'])
            prop1 = entity_str(prop_map['hasPart'])
            codes = [entity_str('INST-%s'%(c)) for c in re.sub(r'[a-z:]*Comp\-INST\-','',code_text).split('-')]
            if len(codes) < 2:
                raise CompositeCharacteristicException("CompositeCharacteristic has less than 2 parts (%s)"%(code_text))
            text1 = '; '.join(["%s %s"%(prop,c) for c in codes])
            # text += "%s rdf:type %s.\n"%(code_text,klass)
            # text += "%s %s [%s].\n"%(code_text,prop1,text1)
            COLLECT_NAMED_CHARS[code_text] = "%s rdf:type %s.\n"%(code_text,klass)
            COLLECT_NAMED_CHARS[code_text] += "%s %s [%s].\n"%(code_text,prop1,text1)
        elif code_text.startswith('INST-'):
            # providing list of codes
            codes = [entity_str(c.strip()) for c in code_text.split(',')]
            codes.sort()
            if len(codes) == 1:
                klass = entity_str(class_map['PrimitiveCharacteristic'])
                prop1 = entity_str(prop_map[prop0])
                code = codes[0]
                text += "%s %s [%s %s].\n"%(inst,prop1,prop,code)

            elif len(codes) > 1:
                comp_inst = entity_str('Comp-INST-'+'-'.join([re.sub(r'[a-z:]*INST-','',c) for c in codes]))
                klass = entity_str(class_map['CompositeCharacteristic'])
                prop1 = entity_str(prop_map['hasPart'])
                prop0 = entity_str(prop_map[prop0])
                text1 = '; '.join(["%s %s"%(prop,c) for c in codes])
                # text += "%s rdf:type %s.\n"%(comp_inst,klass)
                # text += "%s %s [%s].\n"%(comp_inst,prop1,text1)
                text += "%s %s %s.\n"%(inst, prop0, comp_inst)
                COLLECT_NAMED_CHARS[comp_inst] = "%s rdf:type %s.\n"%(comp_inst,klass)
                COLLECT_NAMED_CHARS[comp_inst] += "%s %s [%s].\n"%(comp_inst,prop1,text1)


    return text


# Read main Ex ffle
xls = pd.ExcelFile(dirin+'/'+filein)

# collects skateholder definitions
COLLECT_STAKEHOLDERS = []
COLLECT_NAMED_CHARS = {}





# insert Taxonomy: CodeList classes and instances
try:
    text += "#####################\n# Taxonomies\n####################\n"
    df = pd.read_excel(xls,'Taxonomies', header=1)
    df = df.drop_duplicates().dropna(how='all')
    for (klass, subclass),grp in df.groupby(['Class','subClassOf'], dropna=False):
        if not pd.isnull(subclass):
            text += "\n"
            subj, obj = entity_str(klass),entity_str(subclass)
            text += "%s rdfs:subClassOf %s.\n"%(subj, obj)

        if grp['Instance'].any():
            for _,row in grp.iterrows():
                subj, obj = entity_str(row['Instance']), entity_str(row['Class'])
                text += "%s rdf:type %s.\n"%(subj,obj)
except ValueError as e:
    print(e)


##########################################################
# Get Communities
# to get communities fro collected stakeholders, 
# set IGNORE_COMM_SHEET = True
##########################################################
IGNORE_COMM_SHEET = False
if not IGNORE_COMM_SHEET:
    try:
        # Communities
        df = pd.read_excel(xls,'Communities', header=1)
        df = df.drop_duplicates().dropna(how='all')
        communities = {}
        for _,row in df[~df['Community'].isnull()].iterrows():
            comm = row['Community']
            communities[comm] = {}
            chars = []
            if not pd.isnull(row['CommunityCharacteristic']):
                for s in [s.strip() for s in row['CommunityCharacteristic'].split(',')]:
                    chars.append(s)
            communities[comm]['CommunityCharacteristic'] = list(set(chars))
            communities[comm]['hasNumber'] = row['hasNumber']
            communities[comm]['hasLandArea'] = row['hasLandArea']
            communities[comm]['parcelHasLocation'] = row['parcelHasLocation']

        # CityDivition
        # cp:Community subClassOf iso5087-2:CityDivision
        # ios5087-2:CityDivision iso5087 hasLandArea iso5087-2:LandArea
        # iso587-2:LandArea subClassOf iso5078-1:Manifestation
        #                 landuse_50872:hasLocation exactly 1 iso5087-1:Feature
        # isoFeature subClassOf loc:Feature
        # >>>>>> If Feauter was geo:Feature we woudl use the gml:hasIdentifier
        #         isoFeature subClassOf geo:Feature
        #         geo:Feature gml:identifier "Area1"
        text += "#####################\n# Communities\n####################\n"
        for comm, props in communities.items():
            inst = entity_str(comm)
            klass = entity_str(class_map['Community'])
            text += "%s rdf:type %s.\n"%(inst,klass)

            land = entity_str(props['hasLandArea'])
            prop = entity_str(prop_map['hasLandArea'])
            text += "%s %s %s.\n"%(inst,prop,land)
        
            laklass = entity_str(class_map['LandArea'])
            text += "%s rdf:type %s;\n"%(land,laklass)
            parcel = entity_str(props['parcelHasLocation'])
            prop = 'landuse_50872:parcelHasLocation'
            text += "   %s %s.\n"%(prop, parcel)

            fklass = entity_str(class_map['Feature'])
            text += "%s rdf:type %s.\n"%(parcel, fklass)

            # Community Char
            cklass = entity_str(class_map['CommunityCharacteristic'])
            compchar_inst = entity_str("%s_CommunityCharacteristic"%(inst))
            text += "%s rdf:type %s.\n"%(compchar_inst, cklass)
            prop = entity_str(prop_map['hasCommunityCharacteristic'])
            text += "%s %s %s.\n"%(inst, prop, compchar_inst)
            # Characteristic
            # char_inst = entity_str("%s_Characteristic"%(inst))
            # prop = entity_str(prop_map['hasCharacteristic'])
            # text += "%s %s %s.\n"%(compchar_inst, prop,char_inst)
            if props['CommunityCharacteristic'] == props['CommunityCharacteristic'] and len(props['CommunityCharacteristic'])>0:
                text += format_characteristics_text(compchar_inst, [','.join(props['CommunityCharacteristic'])])

            if props['hasNumber'] == props['hasNumber']:
                num = float(props['hasNumber'])
                prop = entity_str(prop_map['hasNumber'])
                text += "%s %s %s.\n"%(compchar_inst,prop,num)

            text += "\n\n"
    except ValueError as e:
        print(e)

#####################################################
# Org properties
#####################################################
try:
    text += "#####################\n# Organizations\n####################\n"
    df = pd.read_excel(xls,'Organizations', header=1)
    df = df.drop_duplicates().dropna(how='all')
    klass = entity_str(class_map["Organization"])
    for _,row in df.iterrows():
        idinst = row['hasID']
        text += "# ------------------\n# Org (%s)\n# ------------------\n"%(idinst)
        oinst = entity_str(row['Organization'])
        text += "%s rdf:type %s;\n"%(oinst, klass)

        nminst = row["hasLegalName"]
        prop =   entity_str(prop_map['hasLegalName'])
        text += "   %s \"%s\";\n"%(prop, nminst)

        iminst = row["hasImpactModel"]
        iminst = format_lists([iminst])
        prop =   entity_str(prop_map['hasImpactModel'])
        text += "   %s %s;\n"%(prop, ','.join(iminst))

        iminst = row["hasIndicator"]
        iminst = format_lists([iminst])
        prop =   entity_str(prop_map['hasIndicator'])
        text += "   %s %s;\n"%(prop, ','.join(iminst))

        iminst = row["hasOutcome"]
        iminst = format_lists([iminst])
        prop =   entity_str(prop_map['hasOutcome'])
        text += "   %s %s;\n"%(prop, ','.join(iminst))


        inst = entity_str(idinst)
        prop = entity_str(prop_map["hasID"])
        text += "   %s %s;\n"%(prop, inst)

        text += ".\n"
        # Org hasID
        prop = entity_str(prop_map['hasIdentifier'])
        idklass = entity_str(class_map['OrganizationID'])
        text += "%s rdf:type %s;\n"%(inst, idklass)
        text += "   %s \"%s\".\n"%(prop,idinst)


        # Characteristics
        text += format_characteristics_text(inst, row['hasCharacteristic'])
        text += "\n"
except ValueError as e:
    print(e)


#####################################################
# Funding
#####################################################
try:
    text += "#####################\n# Funding\n####################\n"
    df = pd.read_excel(xls,'Funding', header=1)
    df = df.drop_duplicates().dropna(how='all')
    klass = entity_str(class_map["Funding"])

    for (finst, receivedFrom, fundersProgram, receivedAmount, requestedAmount),grp in df.groupby(['Funding','receivedFrom','fundersProgram','receivedAmount','requestedAmount']):
        finst = entity_str(finst)
        text += "\n# Funding %s ------------------\n"%(finst)
        # generate Funding instances, includes organizations and programs
        fklass = entity_str(class_map["Funding"])
        text += "%s rdf:type %s ;\n"%(finst, fklass)

        prop = entity_str(prop_map['receivedAmount'])
        inst = grp["receivedAmount"]
        text += "   %s %s ;\n"%(prop, receivedAmount)

        prop = entity_str(prop_map['requestedAmount'])
        inst = grp["requestedAmount"]
        text += "   %s %s ;\n"%(prop, requestedAmount)

        prop = entity_str(prop_map['fundersProgram'])
        inst = entity_str(fundersProgram)
        text += "   %s %s ;\n"%(prop, inst)


        insts = format_lists(grp["forProgram"])
        prop = entity_str(prop_map['forProgram'])
        if len(insts)==1:
            text += "   %s %s;\n"%(prop, insts[0])
        if len(insts)>1:
            text += "   %s %s;\n"%(prop,', '.join(insts))

        insts = format_lists(grp["forStakeholder"])
        for inst in insts:
            COLLECT_STAKEHOLDERS.append(inst)

        prop = entity_str(prop_map['forStakeholder'])
        if len(insts)==1:
            text += "   %s %s;\n"%(prop, insts[0])
        if len(insts)>1:
            text += "   %s %s;\n"%(prop,', '.join(insts))

        text += '.\n'

        text += '\n'
except ValueError as e:
    print(e)

#####################################################
# Logic Models
#####################################################
try:
    text += "#####################\n# Logic Models\n####################\n"
    df = pd.read_excel(xls,'LogicModels', header=1)
    df = df.drop_duplicates().dropna(how='all')
    klass = entity_str(class_map["LogicModel"])
    for (lminst, ninst),grp in df.groupby(['LogicModel','hasName']):
        lminst = entity_str(lminst)
        text += "%s rdf:type %s ;\n"%(lminst, klass)

        insts = format_lists(grp["forOrganization"])
        prop = entity_str(prop_map["forOrganization"])
        if len(insts)==1:
            text += "   %s %s;\n"%(prop, insts[0])
        elif len(insts)>1:
            text += "   %s %s;\n"%(prop,', '.join(insts))

        prop = entity_str(prop_map['hasName'])
        text += "   %s \"%s\";\n"%(prop, ninst)

        inst = '; '.join(grp["hasDescription"].unique())
        prop = entity_str(prop_map['hasDescription'])
        text += "   %s \"%s\";\n"%(prop, inst)

        insts = format_lists(grp["hasProgram"])
        prop = entity_str(prop_map["hasProgram"])
        if len(insts)==1:
            text += "   %s %s;\n"%(prop, insts[0])
        elif len(insts)>1:
            text += "   %s %s;\n"%(prop,', '.join(insts))

        insts = format_lists(grp["hasStakeholder"])
        for inst in insts:
            COLLECT_STAKEHOLDERS.append(inst)
        prop = entity_str(prop_map['hasStakeholder'])
        if len(insts)==1:
            text += "   %s %s;\n"%(prop, insts[0])
        elif len(insts)>1:
            text += "   %s %s;\n"%(prop,', '.join(insts))
        text += ".\n"

        # Characteristic
        text += format_characteristics_text(lminst, format_lists(grp['hasCharacteristic'],entity=False))

        text += "\n"
except ValueError as e:
    print(e)


#####################################################
# Programs
#####################################################
try:
    text += "#####################\n# Programs\n####################\n"
    df = pd.read_excel(xls,'Programs', header=1)
    df = df.drop_duplicates().dropna(how='all')
    klass = entity_str(class_map["Program"])
    for (pinst,ninst),grp in df.groupby(['Program','hasName']):
        #################################################
        # generate Program Class and links for Service to Program to and Code
        pinst = entity_str(pinst)
        text += "%s rdf:type %s ;\n"%(pinst, klass)
        prop = entity_str(prop_map['hasName'])
        text += "   %s \"%s\";\n"%(prop, ninst)

        inst = '; '.join(grp["hasDescription"].values)
        prop = entity_str(prop_map['hasDescription'])
        text += "   %s \"%s\";\n"%(prop, inst)

        insts = format_lists(grp["hasService"])
        prop = entity_str(prop_map['hasService'])
        if len(insts)==1:
            text += "   %s %s;\n"%(prop, insts[0])
        if len(insts)>1:
            text += "   %s %s;\n"%(prop,', '.join(insts))

        insts = format_lists(grp["hasContributingStakeholder"])
        for inst in insts:
            COLLECT_STAKEHOLDERS.append(inst)
        prop = entity_str(prop_map['hasContributingStakeholder'])
        if len(insts)==1:
            text += "   %s %s;\n"%(prop, insts[0])
        if len(insts)>1:
            text += "   %s %s;\n"%(prop,', '.join(insts))

        insts = format_lists(grp["hasBeneficialStakeholder"])
        for inst in insts:
            COLLECT_STAKEHOLDERS.append(inst)
        prop = entity_str(prop_map['hasBeneficialStakeholder'])
        if len(insts)==1:
            text += "   %s %s;\n"%(prop, insts[0])
        if len(insts)>1:
            text += "   %s %s;\n"%(prop,', '.join(insts))

        text += ".\n"
        text +="\n"
except ValueError as e:
    print(e)

#####################################################
# Services
#####################################################
try:
    text += "#####################\n# Services\n####################\n"
    df = pd.read_excel(xls,'Services', header=1)
    df = df.drop_duplicates().dropna(how='all')
    klass = entity_str(class_map["Service"])
    for (sinst,ninst),grp in df.groupby(['Service','hasName']):
        #################################################
        # generate Program Class and links for Service to Program to and Code
        sinst = entity_str(sinst)
        text += "%s rdf:type %s ;\n"%(sinst, klass)

        prop = entity_str(prop_map['hasName'])
        text += "   %s \"%s\";\n"%(prop, ninst)

        inst = '; '.join(grp["hasDescription"].unique())
        prop = entity_str(prop_map['hasDescription'])
        text += "   %s \"%s\";\n"%(prop, inst)

        insts = format_lists(grp["oep:partOf"])
        prop = entity_str(prop_map['partOf'])
        if len(insts)==1:
            text += "   %s %s;\n"%(prop, insts[0])
        elif len(insts)>1:
            text += "   %s %s;\n"%(prop,', '.join(insts))

        insts = format_lists(grp["hasCode"])
        prop = entity_str(prop_map['hasCode'])
        if len(insts)==1:
            text += "   %s %s;\n"%(prop, insts[0])
        elif len(insts)>1:
            text += "   %s %s;\n"%(prop,', '.join(insts))

        insts = format_lists(grp["hasContributingStakeholder"])
        for inst in insts:
            COLLECT_STAKEHOLDERS.append(inst)
        prop = entity_str(prop_map['hasContributingStakeholder'])
        if len(insts)==1:
            text += "   %s %s;\n"%(prop, insts[0])
        elif len(insts)>1:
            text += "   %s %s;\n"%(prop,', '.join(insts))

        
        insts = format_lists(grp["hasBeneficialStakeholder"])
        for inst in insts:
            COLLECT_STAKEHOLDERS.append(inst)

        prop = entity_str(prop_map['hasBeneficialStakeholder'])
        if len(insts)==1:
            text += "   %s %s;\n"%(prop, insts[0])
        if len(insts)>1:
            text += "   %s %s;\n"%(prop,', '.join(insts))

        insts = format_lists(grp["hasMode"], entity=False)
        insts = [entity_str(t) for t in insts]
        prop = entity_str(prop_map['hasMode'])
        if len(insts)==1:
            text += "   %s %s;\n"%(prop, insts[0])
        elif len(insts)>1:
            text += "   %s %s;\n"%(prop,', '.join(insts))


        text += ".\n"

        # generate Service hasRequirement to Client Codes
        for r in list(set(grp['hasRequirement'].values)):
            text += format_characteristics_text(sinst, [r], prop0='hasRequirement')
        # generate Service hasFocus to Client Codes
        for r in list(set(grp['hasFocus'].values)):
            text += format_characteristics_text(sinst, [r], prop0='hasFocus')


        text += "\n"
except ValueError as e:
    print(e)


#####################################################
# ServiceEvents
#####################################################
try:
    text += "#####################\n# ServiceEvents\n####################\n"
    df = pd.read_excel(xls,'ServiceEvents', header=1)
    df = df.drop_duplicates().dropna(how='all')
    klass = entity_str(class_map['ServiceEvent'])
    for _,row in df.iterrows():
        
        seinst = entity_str(row['ServiceEvent'])
        text += "%s rdf:type %s;\n"%(seinst, klass)

        # strings, no namespace
        for col in ['hasName','hasDescription']:
            if not pd.isna(row[col]):
                inst = row[col]
                prop = entity_str(prop_map[col])
                text += "   %s \"%s\";\n"%(prop, inst)

        # annotations with namespace
        for col in ['hasStatus', 'forClient','AtOrganization','forReferral','hasLocation','previousEvent','nextEvent']:
            if not pd.isna(row[col]):
                inst = entity_str(row[col])
                prop = entity_str(prop_map[col])
                text += "   %s %s;\n"%(prop, inst)

        # Dates with convertion: YYYY-CCYY-MM-DD HH:MM:SS to CCYY-MM-DDThh:mm:ss.sss
        for col in ['occursAt','hasBeginning', 'hasEnd']:
            inst = date_to_xsd(row[col])
            if not pd.isna(inst):
                prop = entity_str(prop_map[col])
                text += "   %s \"%s\";\n"%(prop, inst)
            

        text += ".\n"
        text += "\n"
except ValueError as e:
    print(e)


#####################################################
# Clients
#####################################################
try:
    text += "#####################\n# Clients\n####################\n"
    df = pd.read_excel(xls,'Clients', header=1)
    df = df.drop_duplicates().dropna(how='all')
    klass = entity_str(class_map['Client'])
    for _,row in df.iterrows():
        # break    
        cinst = entity_str(row['Client'])
        text += "%s rdf:type %s;\n"%(cinst, klass)

        # strings, with namespace
        for col in ['hasDescription']:
            if not pd.isna(row[col]):
                inst = row[col]
                prop = entity_str(prop_map[col])
                text += "   %s \"%s\";\n"%(prop, inst)

        # annotations with namespace
        for col in ['hasIdentifier','satisfiesStakeholder','hasGender','hasSex','hasIncome','hasSkill','hasEthnicity','memberOfAboriginalGroup','hasReligion','hasDependent','schema:knowsLanguage','hasOutcome','hasNeed','hasGoal','hasProblem','hasStatus','hasClientState']:
            if not pd.isna(row[col]):
                inst = entity_str(row[col])
                prop = entity_str(prop_map[col])
                text += "   %s %s;\n"%(prop, inst)
            
        text += ".\n"
        text += "\n"

        COLLECT_STAKEHOLDERS.append(row['satisfiesStakeholder'])
except ValueError as e:
    print(e)


########################################################
# Collect and write Stakeholder defintions to text file
########################################################
# To ignore the strakeholder sheet, you can collect the ones found in other sheets.
# If you want to add the collected stakeholders into the sheet, 
# then import the printed output into the sheet, and replace existing content
#      set IGNORE_SH_SHEET = True
IGNORE_SH_SHEET = False

stakeholders = {}
# get stakeholders collected form other sheets
for sids in COLLECT_STAKEHOLDERS:
    for sid in sids.split(','):
        sid = entity_str(sid)
        # get each part, skip first, assuming 'sh-' or some prefix
        parts = re.findall(r'([^\-]+)',sid)[1:]
        areas = [p for p in parts if p.startswith('in_')]
        chars = list(set([p for p in parts if not p.startswith('in_')]))
        chars.sort()
        stakeholders[sid] = {'hasCode':[], "location":np.nan}
        if len(chars) > 0:
            stakeholders[sid]['hasCode'] = ['INST-%s'%(char) for char in chars]
        if len(areas) > 0:
            stakeholders[sid]['location'] = areas[0].replace('in_','') + '_Location'
        

# Get stakeholders from 'Stakeholders' sheet
if not IGNORE_SH_SHEET:
    try:
        df = pd.read_excel(xls,'Stakeholders', header=1)
        df = df.drop_duplicates().dropna(how='all')
        sids = []
        for _,row in df.iterrows():
            sids.append(row['Stakeholder'].split(','))
        sids = list(set(flatten(sids)))
        sids.sort()


        # initialize any stakeholdes that have not been found already
        for sid in sids:
            sid = entity_str(sid)
            if sid not in stakeholders.keys():
                stakeholders[sid] = {'hasCode':[], "location":np.nan}
        for _,row in df.iterrows():
            for sid in row['Stakeholder'].split(','):
                sid = entity_str(sid)
                if stakeholders[sid]['hasCode'] == []:
                    if row['hasCode'] == row['hasCode']:
                        for code in [c.strip() for c in row['hasCode'].split(',')]:
                            stakeholders[sid]['hasCode'].append(code)
                if pd.isnull(stakeholders[sid]['location']) and not pd.isnull(row['location']):
                    stakeholders[sid]['location'] = row['location']
    except ValueError as e:
        print(e)

# sort codes for all stakeholders
for k,v in stakeholders.items():
    stakeholders[k]['hasCode'] = list(set(v['hasCode']))
    stakeholders[k]['hasCode'].sort()

# insert stakeholders into file
text += "#####################\n# Stakeholders\n####################\n"
for sid,props in stakeholders.items():
    text += "\n# Stakeholder (%s)\n"%(sid)
    sklass = entity_str(class_map["Stakeholder"])
    sinst = entity_str(sid)

    text += "%s rdf:type %s.\n"%(sinst, sklass)
    prop = prop_map['locatedIn']
    if props['location']== props['location']:
        inst = entity_str(props['location'])
        text += "%s %s %s.\n"%(sinst, prop, inst)

    text += format_characteristics_text(sinst, props["hasCode"])

    text += "\n"

if IGNORE_SH_SHEET:
    for k,v in stakeholders.items():
        print("\"%s\",\"%s\",\"%s\""%(k.split(':')[1],','.join(v['hasCode']),v['location']))


if IGNORE_COMM_SHEET:
    for k,v in stakeholders.items():
        Community = k.split(':')[1].replace('sh-','')
        tmp = Community.split('-in_')
        if len(tmp)==1:
            Community = tmp[0].replace('in_','')
        else:
            Community = '-'.join(tmp[::-1])
        CommunityCharacteristic = v['hasCode']
        parcelHasLocation = v['location']
        if type(parcelHasLocation) is str:
            hasLandArea = parcelHasLocation.replace('_Location', '_Land_Area')
        else:
            hasLandArea = np.nan
        print("\"%s\",\"%s\",\"\",\"%s\",\"%s\""%(Community, ','.join(CommunityCharacteristic),hasLandArea, parcelHasLocation))

# write out collected named Characteristics, e.g. CompositeCharacteristic
text += "#####################\n# Named Charcteristics\n####################\n"
for k,v in COLLECT_NAMED_CHARS.items():
    text += v+"\n"


#######################################################################
# Write ttl content to file
#######################################################################
f = open(dirout+'/'+fileout, "w")
f.write(text)
f.close()
