from osvc_python import *
from win10toast import ToastNotifier
from time import sleep, time
import sys
from getpass import getpass
import urllib.parse
import urllib.request
import os
import random
import pymsteams
from tkinter import *
from PIL import Image, ImageTk
import json
import traceback
import datetime
import re
import requests
import email
from io import StringIO
import base64
from _io import BytesIO
from test.support import resource
from urllib import response


class SendToTeams:
    def __init__(self, team, text, incident, team_text, custom_message, subject, formatted_date):
        self.team = team
        self.text = text
        self.incident = incident
        self.team_text = team_text
        self.custom_message = custom_message
        self.subject = subject
        self.connector = None
        self.webhook = None
        self.formatted_date = formatted_date


    def connect(self):
        self.connector = pymsteams.connectorcard(self.webhook)
    
    def add_text(self):
        self.connector.summary("Severity 1 ärende")
        myMessageSection = pymsteams.cardsection()
        myTextSection = pymsteams.cardsection()
        myMessageSection.title(self.custom_message)
        myMessageSection.activityTitle(self.subject)
        myMessageSection.activityImage("https://i.vimeocdn.com/portrait/27569337_640x640")
        myMessageSection.addFact('ID', self.incident)
        myMessageSection.addFact("Assigned to", "None/JeevesERP")
        myMessageSection.addFact("Team", self.team_text)
        myMessageSection.addFact("Ärende Inkommit", self.formatted_date)
        myTextSection.text(self.text)
        self.connector.addSection(myMessageSection)
        self.connector.addSection(myTextSection)

    def send_text(self):
        self.connector.send()

    def get_web_hook(self):
        if self.team == 8:
            self.webhook = ''
        if self.team == 6:
            self.webhook = ''
        if self.team == 7:
            self.webhook = ''
        else:
            self.webhook = ''

def rn_start(username, password):
    return OSvCPythonClient(
        
        ## Interface to connect with 
        interface='jeeves',
        
        ## Basic Authentication
        username=username,
        password=password,
        
        ## Session Authentication
        # session=<session_token>,
        
        ## OAuth Token Authentication
        # oauth=<oauth_token>,

        ## Optional Client Settings
        demo_site=False,						## Changes domain from 'custhelp' to 'rightnowdemo'
        version="v1.3",						## Changes REST API version, default is 'v1.3'
        no_ssl_verify=False,					## Turns off SSL verification
        suppress_rules="True",				## Supresses Business Rules
        access_token="My access token",		## Adds an access token to ensure quality of service
    )



def get_v_and_k_from_org(org_id, rn_client):
    #vi selectar på Product Null och 1 för att vissa poster är NULL
    try:

        res_org = OSvCPythonQueryResults().query(
                    query=f"SELECT Version, Other, Product, orgid from CO.AssetData C where orgid={org_id} and (Product is NULL or Product=1)",
                    client=rn_client,
                    annotation="Custom annotation",
                    exclude_null=True
        )

        if len(res_org) > 0:
            return {
            "kernel": res_org[0]['Other'],
            "version": res_org[0]['Version'],
            "orgid": res_org[0]['orgid']
            }
        else:
            return {
            "kernel": 'Other',
            "version": 'Version',
            "orgid": 144
            }
    except Exception:
        return {
            "kernel": 'Other',
            "version": 'Version',
            "orgid": 144
            }


def version(ver_string):
    #regex för att fixa kernel/version
    get_ver_reg = re.compile('([0-9]*[0-9])')
    get_ver_fix = re.compile('(\d{3,5})')
    try:
        #list here in case of multiple versions kernels being enumarated.
        ver_strings_list = list()
        ver_strings_list.append(ver_string)
        #create a list here so we can test stuff later incase the function cant handle
        for idx, ver_to_fix in enumerate(ver_strings_list):
            joiner = str()
            verx = get_ver_reg.findall(ver_to_fix)
            if len(verx) <2:
                lenx = get_ver_fix.findall(verx[0])
                xd=str()
                for idx, jas in enumerate(lenx[0]):
                    if idx == 0:
                        xd += jas
                    if idx != 0:
                        xd += jas
                        if len(lenx[0]) != idx+1:
                            xd += '.'
                    if len(lenx[0]) == idx+1:
                        return xd

            else:
                for idx, ver_letter in enumerate(verx):
                    joiner += ver_letter
                    if len(verx) == idx+1:
                        return joiner[:6]
                    joiner += '.'
    except Exception:
        print('Något gick fel med att hämta ut applikationsversionen från strängen. Retunerar Other')
        return "N/A"


def kernel(kernel_string):
    #regex för att fixa kernel/version
    get_ver_reg = re.compile('([0-9]*[0-9])')
    get_ver_fix = re.compile('(\d{3,5})')
    #create a list here so we can test stuff later easy without adding code.
    try:
        kernel_strings = list()
        kernel_strings.append(kernel_string)
        for ker_to_fix in kernel_strings:
            kernel = str()

        
            #regex to get all numbers from kernel
            #1420 - list with 1 in length
            #1x4x2 - list with 3 in length (x means anything but a number)
            kerx = get_ver_reg.findall(ker_to_fix)
            if ker_to_fix[0] == 'R':
                R_first = ker_to_fix[0]
                
            #if kernel starts with 5 we know it's like 5xxx something
            #function has 3 if else because its easier to maintain incase i need to fix it in the future
            if kerx[0] == '5':
                for idx, kern in enumerate(kerx):
                    kernel += kern
                    kernel += '.'
                    if idx == len(kerx) -1:
                        #in case kernels are written really long, we need to take out 5.1 R1 for example only.
                        # R is part of it aswell
                        fix_dot_kernel = kernel[:3] + " R" + kernel[4:6]
                        r_over_ten = re.compile('R1[0-9]{0,1}')
                        r_whitespace_over_ten = re.compile('R 1[0-9]{0,1}')
                        match = r_over_ten.findall(fix_dot_kernel)
                        if len(match) > 0:
                             return fix_dot_kernel.strip()
                        else:
                            return fix_dot_kernel[:6].strip()


            #if length of list is longer than 1 and doesnt start with 5 know it's like 14.xxx with dots
            elif kerx[0] != '5' and len(kerx) > 1:                           
                for idx, kerne in enumerate(kerx):
                        kernel += kerne
                        kernel += "."
                        if (idx == len(kerx)-1):
                            #we need to take out only 14.0.1 for example, because otherwise too long
                            kernel_old = kernel[:5]
                            if kernel_old[-1] == '.':
                                return kernel_old + '0'
                            return kernel_old
            #if kernel doesnt start with 5 and length is only 1 we know it's like 1420.
            else:                          
                for idx, kerne in enumerate(kerx[0]):
                        kernel += kerne
                        kernel += "."
                        if (idx == len(kerx[0])-1):
                            kernel_no_dots = kernel[:5]
                            return kernel_no_dots
    except Exception as e:
        print('Något gick fel med hanteringen av Kernel/ver')
        return "Other"


#version_fixed = version(version_and_kernel['version'])
#print(version_fixed)
#kernel_fixed = kernel(version_and_kernel['kernel'])
#print(kernel_fixed)


'''
function: get_version_id
Hämtar rätt ID av versionen för att uppdatera ärendet med.
Använder en formaterad versions-sträng som den loopar igenom
en JSON och tar fram rätt ID

'''


def get_version_kernel_id(version_number, json_file, vk_version, verker):
    json_url = urllib.request.urlopen(json_file)
    loaded_json = json.loads(json_url.read().decode())
    #loaded_json = json.load(open(json_file, 'r', encoding='utf-8'))
    for count, version in enumerate(loaded_json[vk_version]):
        if version_number == version[verker]:
            return version['ID']
    if verker == 'Version':
        print(f'Kunde inte hitta {version[verker]}. Förmodligen inte kodad som : (Applikation 5.1 R2, 13.0.2 etc) (Kernel: 5.1 R2, 3.5.2)')
        return 58
    else:
        print(f'Kunde inte hitta {version[verker]}. Förmodligen inte kodad som : (Applikation 5.1 R2, 13.0.2 etc) (Kernel: 5.1 R2, 3.5.2)')
        return 109


'''
function: get_kernel_id
Hämtar rätt ID av versionen för att uppdatera ärendet med.
Använder en formaterad versions-sträng som den loopar igenom
en JSON och tar fram rätt ID
'''


"""def get_kernel_id(kernel_number, json_file, kernel_version, version):
    loaded_json = json.load(open(r'kernel_list.json', 'r', encoding='utf-8'))
    for count,version in enumerate(loaded_json['Kernel Version']):
		if kernel_number == version['Kernel']:
			return version['ID']
    return 109"""



def extract_cc_from_headers(incident):
    #regex för mailadresser
    regex_email_cc = '(?<=\<).*?(?=\>)'
    #om ingen mailHeader finns, har ärendet kommit in via portalen. Ingen CC behöves läggas in
    if incident['mailHeader'] is None:
        return False
    else:
        #tar ut en sträng från mailheadern
        msg = email.message_from_string(incident['mailHeader'])
        #tar ut CC från mailheader
        cc = msg['cc']
        #tar ut to från mailheader
        to = msg['to']
        #tar ut endast emails från to, filtrerar bort tecken som < > och Namn
        to_list = re.findall(regex_email_cc, to)

        if cc is None and len(to_list) > 1:
            for i, value in enumerate(to_list):
                if 'support@jeeves.se' in value:
                    #tar bort jeeves-support ur to.
                    del to_list[i]
            result = set(to_list)
            cc_final = ','
            #får en sträng med mails separerade vid ,
            cc_final = cc_final.join(result)
            return cc_final

        #om CC finns
            #tar ut endast emails från cc, filtrerar bort tecken som < > och Namn
            #om to-listan är mer än 1, vet vi att mer än bara jeeves-support finns där..
        if cc is not None:
            #tar ut endast emails från cc, filtrerar bort tecken som < > och Namn
            ccs = re.findall(regex_email_cc, cc)
            #om to-listan är mer än 1, vet vi att mer än bara jeeves-support finns där..
            if len(to_list) > 0:
                for i, value in enumerate(to_list):
                    if 'support@jeeves.se' in value:
                        #tar bort jeeves-support ur to.
                        del to_list[i]
                #slår ihop listorna
                result = list(set(ccs) | set(to_list))
                for i, value in enumerate(result):
                    if 'support@jeeves.se' in value:
                        #tar bort jeeves-support ur to.
                        del result[i]
                cc_final = ';'
                #får en sträng med mails separerade vid ,
                cc_final = cc_final.join(result)
                return cc_final
            elif len(to_list) == 1:
                #om to-list är == 1 och har CC vet vi att endast CC behöver slås ihop
                cc_single = ';'
                cc_single = cc_single.join(ccs)
                return cc_single
    return False


def analyze_text(text):
    try:
        is_forced = False
        words_list = list()
        text = text.replace('Jeeves confidentiality notice: This email and any attachment(s) are confidential and may be privileged and are intended only for the authorized recipients of the sender. The information contained in this email and any attachment(s) must not be published, copied, disclosed, or transmitted in any form to any person or entity unless expressly authorized by the sender. If you have received this email in error you are requested to delete it immediately and advise the sender by return email', ' ')
        placeholder_dict = dict()
        saved_category = str()
        #loads json data
        inci_keywords = keywords_json()
        for x, y in inci_keywords.items():
            #loops through teams
            for j, k in y.items():
                #loops through the different items for each index
                keyword_list = k["keywords"]
                #saves keyword items to a keyword_list
                for value in keyword_list:
                    #loops through every value of every keyword list in the keywords dict
                    keywords_to_check = value['keyword']
                    foreced = value['forced']
                    #saves value of the whole array in the keywords_to_check
                    if "," in keywords_to_check:
                        keyword_list = filter(None, keywords_to_check.split(', '))
                        #filters all the words that can be split with "word, "
                        for kw in keyword_list:
                            regex_match = re.compile(check_if_wild(value['wild'], kw.lower()))
                            #checks in the dict if the current value is wild
                            matches = regex_match.findall(text.lower())
                            #matches the current kw if it matches with the regex
                            length = len(matches)
                            if length > 0:
                                #if length > 0 we know its a match
                                if matches[0] == kw:
                                    value["matches"] += length
                                    placeholder_dict.update({j : {
                                    }})
                                    try:
                                        if k["count"] >= 0:
                                            words_list.append(kw)
                                            placeholder_dict[j].update({"words": words_list})
                                        else:
                                            words_list = []
                                            words_list.append(kw)
                                            placeholder_dict[j].update({"words": words_list})
                                        saved_text = text
                                    except Exception as e:
                                        print("Fel i Gissar-skriptet!", e)

                                    k["count"] += length


                                    placeholder_dict[j].update({"count": k['count'], "severity": value['severity'], "category": j, "Team": x, "head_category": k['head_category'], "rnid": k["rnid"]} )





                                    
                                if value['forced'] == 'G':
                                    return [j, k['count'], x, value['severity'], k['head_category'], k["rnid"]]
        category = max(placeholder_dict.items(), key=lambda x: x[1]['count'])
        highest_count = [category[1]['category'], category[1]['count'], category[1]['Team'], category[1]['severity'], category[1]['head_category'], category[1]['rnid'], category[1]['words']]
        return highest_count
    except Exception as e:
        print("Hittade inga Matchande Ord")
        return ["Other", 3, 2, 3, "doesnt matter", 144, []]



def check_if_wild(wild, kw):
    if wild == "N":
        return fr'(?i)\b{str(kw)}\b'
    else:
        return fr'(?i){str(kw)}'


def resource_path(relative_path):
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(os.path.abspath("."), relative_path)

def update_v_or_k(incident_id, category, id_of_v, current_v, force_change, setting, site, move_items, already_categorized):
    if (current_v is None and setting == True) or (force_change == True and setting == True) or (already_categorized == True and move_items == True and current_v is None):


        return {category:{
                        "id": int(id_of_v)
            }}


    return current_v
def update_severity(incident_id, severity, current_severity, force_change, setting, site):
    if (current_severity is None and setting == True) or (force_change == True and setting == True):

        if current_severity is not None:
            severity = current_severity

        return {
                "id": int(severity)
            }
    
    return current_severity

def update_product(incident_id, product, force_change, setting, site):
    if (product is None and setting == True) or (force_change == True and setting == True):
        return {
            "id": 1
        }
    
    return product


def update_inci_type(incident_id, org_id, inci_type, force_change, setting, site):
    if (inci_type is None and setting == True) or (force_change == True and setting == True):
        if int(org_id) == 976:
            new_inci_type = 30
        else:
            new_inci_type = 29

        return {
            "inc_types": {
                "id": new_inci_type
            }

            }
    
    return inci_type

def update_queue(incident_id, firstline_queue, queue, force_change, setting, site, move_items, already_categorized):
    if (int(firstline_queue) == 2 and setting == True) or (force_change == True and setting == True) or (already_categorized == True and move_items == True): 
        return {
            "id": int(queue)
            }
    
    return queue

def update_category(incident_id, category_update, current_category, force_change, setting, site):
    if (current_category is None and setting == True) or (force_change == True and setting == True):
        return { 
            "id": int(category_update)
             }
    return "current"

def update_cc(incident_id, cc, site):
    if cc == False:
        print("ingen CC!")
        return
    else:
        return {
                    "alternateemail": cc
                }
    return False




def get_queue_from_category(current_category, analyzed):
    if current_category is not None:
        categories = keywords_json()
        for category in categories:
            for cat in categories[category]:
                if int(current_category) == categories[category][cat]['rnid']:
                    number = convert_team_to_number(category)
                    return [number, cat]
    return [analyzed[2], analyzed[5]]



def show_toaster(subject, message, severity, check_sev, time):
    toaster = ToastNotifier()
    print(f"Varning skickad. ALla nuvarande ärende hanterade, kollar igen om {(time+30) / 60} min")
    if len(severity) > 0 and check_sev =="Y":
        toaster.show_toast("1:A I KÖN!", "1A I KÖN!", icon_path=resource_path('ICON.ico'), duration=30)
    else:
        toaster.show_toast(subject, message, icon_path=resource_path('ICON.ico'), duration=30)


def print_scanned_incident(date,single_incident):
    print(f"{date} - {single_incident['lookupName']} - {single_incident['subject']} - Severity: {single_incident['severity']}")







'''def send_teams_sev(queue, text, incident):
    p1 = SendToTeams()
    p1.connect()
    p1.addText()
    p1.sendText()   '''


def warn_for_severity_ones(rn_client, check_again, first_time):
    try:
        sev_one_list = OSvCPythonQueryResults().query(
                        query=f"SELECT DISTINCT I.ID, I.lookupName, min(I.threads.createdTime), I.subject, I.severity, I.threads.text, I.queue, I.statusWithType.*, I.assignedTo.* from INCIDENTS I where (I.assignedTo.account IS NULL or I.assignedTo.staffGroup.id = 100352) and I.severity=1 and (I.statusWithType.status=1 or I.statusWithType.status=8) and (I.queue=6 or I.queue=8 or I.queue=7) group by lookUpName",
                        client=rn_client,
                        annotation="Custom annotation",
                        exclude_null=True
        )


        for incident in sev_one_list:
            date_formatted = datetime.datetime.strptime(incident["min(I.threads.createdTime)"], "%Y-%m-%dT%H:%M:%SZ").strftime("%b %d %Y %H:%M:%S")
            if (int(incident['id']) in sev_ones_sent and check_again == True) or (int(incident['id']) not in sev_ones_sent and check_again == False and first_time == True):
                if check_again == True and first_time == False:
                    custom_message = f'Ingen har tagit Severity 1 ärendet'
                    print(f'Varning skickad igen för ärende: {incident["lookupName"]}\n')
                elif check_again == False and first_time == True:
                    custom_message = 'Severity 1 varning'
                    print(f'Varning skickad för första gången för ärende: {incident["lookupName"]}')
                text = cleanhtml(incident['text'])
                if check_again == False and first_time == True:
                    sev_ones_sent.append(int(incident['id']))
                text = text.replace('Jeeves confidentiality notice: This email and any attachment(s) are confidential and may be privileged and are intended only for the authorized recipients of the sender. The information contained in this email and any attachment(s) must not be published, copied, disclosed, or transmitted in any form to any person or entity unless expressly authorized by the sender. If you have received this email in error you are requested to delete it immediately and advise the sender by return email', ' ')
                sender = SendToTeams(int(incident['queue']), text, incident['lookupName'], convert_number_to_team(int(incident['queue'])), custom_message, incident['subject'], date_formatted)
                sender.get_web_hook()
                sender.connect()
                sender.add_text()
                sender.send_text()
    except Exception as e:
        sleep(60)
        print('Ingen teams varning skickad :(')




def format_incident_for_api(category, inci_type, product, severity, cc, app_fetched, kernel_fetched, queue):
    formatted_dict = {}
    if isinstance(category, dict):
        formatted_dict.update({"category": category})
    if isinstance(product, dict):
        formatted_dict.update({"product": product})
    if isinstance(severity, dict):
        formatted_dict.update({"severity": severity})
    if isinstance(kernel_fetched, dict):
        if "customFields" not in formatted_dict:
            formatted_dict.update({"customFields" : { "c": {
            }}})
        formatted_dict["customFields"]["c"].update(kernel_fetched)

    if isinstance(app_fetched, dict):
        if "customFields" not in formatted_dict:
            formatted_dict.update({"customFields" : { "c": {
            }}})
        formatted_dict["customFields"]["c"].update(app_fetched)

    if isinstance(cc, dict):
        if "customFields" not in formatted_dict:
            formatted_dict.update({"customFields" : { "c": {
            }}})
        formatted_dict["customFields"]["c"].update(cc)

    if isinstance(inci_type, dict):
        if "customFields" not in formatted_dict:
            formatted_dict.update({"customFields" : { "c": inci_type}})
        formatted_dict["customFields"]["c"].update(inci_type)
    if isinstance(queue, dict):
        formatted_dict.update({"queue": queue})

    return formatted_dict



def update_incident(formatted_incident, site, incident_id, automate):

    url = f"https://{site}/services/rest/connect/v1.3/incidents/{incident_id}"
    header = {"X-HTTP-Method-Override": "PATCH", 
        "Content-Type": "application/json"
    }

    
    if(bool(formatted_incident) == True and automate=='Y'):
        res = requests.post(url, json=formatted_incident, headers=header)
        if res.status_code == 200:
            print(f"Kategori uppdaterat! - statuskod: {res.status_code}")
            return True
        else:
            print(res.status_code)
            print(f'Något gick fel med uppdateringen av kategori - statuskod: {res.status_code}')
            return False
    else:
        print('Allt redan ifyllt. Ingen uppdatering gjordes')





def add_to_text(incident_to_add_idonly, incidents_to_add_full, settings_data, rn_client, username, password):
    site = get_test_or_prod_site(settings_data, username, password)
    already_categorized = False
    incident_send_to_api = dict();
    

    with open('incidents.txt', "a") as add_to_incident:
        add_to_incident.write('\n'.join(incident_to_add_idonly) +"\n")
        date = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        for single_incident in incidents_to_add_full:
            #om Organizationen är Jeeves hämtar vi från inmatningsfältet istället
            if int(single_incident['organization']) == 976 and single_incident['product_version'] is not None or single_incident['product_other'] is not None:
                version_fixed = version(single_incident['product_version'])
                kernel_fixed = kernel(single_incident['product_other'])
            else:
                version_and_kernel = get_v_and_k_from_org(int(single_incident['organization']), rn_client)
                version_fixed = version(version_and_kernel['version'])
                kernel_fixed = kernel(version_and_kernel['kernel'])

            #hämtar version från en versionslista formaterad
            version_id = get_version_kernel_id(version_fixed, r'https://dandelion-silky-lodge.glitch.me/appversion_v2.json', 'app version', 'Version')
            
            #hämtar version från en versionslista formaterad
            kernel_id = get_version_kernel_id(kernel_fixed, r'https://dandelion-silky-lodge.glitch.me/kernelversion_v2.json', 'Kernel Version', 'Kernel')

            #hämtar CC
            ccs = extract_cc_from_headers(single_incident)

            #returnerar kategori-nummer om vi har en kategori satt

            print_scanned_incident(date, single_incident)
            #retunerar en array med information om incidenten, te.x kategori
            text = cleanhtml(single_incident['subject'] + " " + single_incident['text'])
            analyzed_text = analyze_text(text)
            incident_category = analyzed_text[5]

            category_current = get_queue_from_category(single_incident['category'], analyzed_text)

            #om den inte har kategori.. måste den gissa
            if single_incident['category'] is None:
                print(f"Jag tror att denna är KATEGORI: {analyzed_text[0]} --- TEAM: {analyzed_text[2]} med Severity {analyzed_text[3]}")
                if len(analyzed_text[6]) > 0:
                    print(f"Unika Ord:{', '.join(analyzed_text[6])} med {analyzed_text[1]} matches")
                print(f"Applikationsversion:{version_fixed} Kernelversion: {kernel_fixed}")
                already_categorized = False
           #om den har kategori... då är det redan ifyllt
            if single_incident['category'] is not None:
                print(f"Denna är redan ifylld Kategori: {category_current[1]} --- med Severity {single_incident['severity']}")
                print(f"Med applikationsversion:{version_fixed} och kernelversion: {kernel_fixed}")
                print(f"---")
                print(f"Marcus Bot hade gissat: {analyzed_text[0]} --- TEAM: {analyzed_text[2]} med Severity {analyzed_text[3]}")
                if len(analyzed_text[6]) > 0:
                    print(f"Ord:{', '.join(analyzed_text[6])} med {analyzed_text[1]} matches")
                already_categorized = True

            inc_id = int(single_incident['id'])
            
            category_updated = update_category(int(single_incident['id']), analyzed_text[5], single_incident['category'], settings_data['force_setting']['force_category'], settings_data['auto_settings']['auto_category'], site)
            inci_type_fetched = update_inci_type(int(single_incident['id']), int(single_incident['organization']), single_incident['inc_types'], settings_data['force_setting']['force_type'], settings_data['auto_settings']['auto_type'], site)
            product_fetched = update_product(int(single_incident['id']), single_incident['product'], settings_data['force_setting']['force_product'],settings_data['auto_settings']['auto_product'], site)
            severity_fetched = update_severity(int(single_incident['id']), analyzed_text[3], single_incident['severity'], settings_data['force_setting']['force_severity'], settings_data['auto_settings']['auto_severity'], site)
            category_analyzed = convert_team_to_number(analyzed_text[2])
            cc_fetched = update_cc(int(single_incident['id']), ccs, site)
            current_new_cat = current_cat_or_new_cat(category_updated, category_current)
            #flyttar över till rätt kö. Om kategori redan finns, väljer den det istället
            move_items = settings_data["move_categorized_items"]
            app_fetched = update_v_or_k(int(single_incident['id']), "app_version", version_id, single_incident['app_version'], settings_data['force_setting']['force_app_v'], settings_data['auto_settings']['auto_ker_v'], site, move_items,already_categorized)
            kernel_fetched = update_v_or_k(int(single_incident['id']), "kernel_version", kernel_id, single_incident['kernel_version'], settings_data['force_setting']['force_ker_v'], settings_data['auto_settings']['auto_app_v'], site, move_items, already_categorized)
            queue = update_queue(int(single_incident['id']), single_incident['queue'], current_new_cat , settings_data['force_setting']['force_queue'],settings_data['auto_settings']['auto_queue'], site, move_items, already_categorized)
            
            formatted_incident = format_incident_for_api(category_updated, inci_type_fetched, product_fetched, severity_fetched, cc_fetched, app_fetched, kernel_fetched, queue)

            update_incident(formatted_incident, site, single_incident['id'], settings_data['automate'])
            print('Ärende hanterade och fixade!')
            #check for severity

            
            #Byggt en annan lösning för detta redan, vet om detta ska användas.
            #if isinstance(severity_fetched, dict) and int(severity_fetched['id']) == 1 and already_categorized:
                #warn_for_severity_ones(rn_client, False, True)



    
    add_to_incident.close()
'''
funktion: current_cat_or_new_cat
vad: kollar från requesten vad som har retunerats. Om
Kategorin var tom, kommer kön den har analyserats fram att användas.
Annars kommer nuvarande kö som finns användas.
'''

def current_cat_or_new_cat(category_updated, queue_current):
    if category_updated == 'current':
        print('here')
        return queue_current[0]
    else:
        return convert_team_to_number(queue_current[0])


def fix_to_list(incidents_to_txt): 
    for a in incidents_to_txt:
        yield a['lookupName']


def convert_team_to_number(team):
    if team == 'SCM':
        return 8
    if team  == 'Finance':
        return 6
    if team == 'Tech':
        return 7
    else:
        return 2

def convert_number_to_team(number):
    if number == 8:
        return 'SCM'
    if number == 6:
        return 'Finance'
    if number == 7:
        return 'Tech'
    else:
        return 2



def check_if_checked(results):
    with open('incidents.txt', "r") as checked_incidents:
        incidents = checked_incidents.read().splitlines()
        #retunerar alla (x) från incidents om de inte lookUpName, dvs incident finns i incidents
        #kan läsas som retunera endast x från results om loopade incidenten inte finns i incidents
        incidents_not_in_txt = [x for x in results if x['lookupName'] not in incidents]
        #retunerar 1orna från results
        severity = [x for x in results if x['severity'] == '1']
    return [incidents_not_in_txt, severity]


def fix_to_list(incidents_to_txt):
    for a in incidents_to_txt:
        yield a['lookupName']


def keywords_json():
    try:
        keywords = json.load(open(r'keywords_v2.json', 'r', encoding='utf-8'))
        return keywords
    except Exception:
        sleep(60)
        print('keywords felkodad. Få tag i en backup eller kontakta Marcus')

def settings_json():
    try:
        settings = json.load(open(r'settings.json', 'r', encoding='utf-8'))
        return settings

    except Exception:
        sleep(60)
        print('settings är felkodad. Få tag i en backup eller kontakta Marcus')

def cleanhtml(raw_html):
  re_html = re.compile('<.*?>|&([a-z0-9]+|#[0-9]{1,6}|#x[0-9a-f]{1,6});')
  cleantext = re.sub(re_html, '', raw_html)
  return cleantext


def get_user_pw(settings_data, pw_or_usr):
    if settings_data['default_account'] == "Y":
        if pw_or_usr == 'username':
            return 'marcus.lundgren'
        elif pw_or_usr == 'password':
            return 'hejhej123'
    else:
        return settings_data[pw_or_usr]

    


def convert_to_queue(number):
    if number == 7:
        return "Tech"
    if number == 6:
        return "Finance"
    if number == 2:
        return "First Line"
    if number == 8:
        return "SCM"



def tkShowFahlStrom():
    text_file = urllib.request.urlopen("https://dandelion-silky-lodge.glitch.me/fahlstrom.txt")
    asd = text_file.read()
    d = asd.split(b"\n")
    msg = base64.b64decode(random.choice(d))
    buf = BytesIO(msg)
    img = Image.open(buf)
    width, height = img.size
    a_root = Tk()
    a_root.attributes('-topmost',True)
    a_root.focus_force()

    a_root.title("❤️❤️❤️ FAHLSTRÖM ❤️❤️❤️")
    a_root.geometry(f"{width}x{height}")
    a_root.maxsize(1920, 1080)


    photo = ImageTk.PhotoImage(img)
    label = Label(a_root, text='', image=photo, compound='center')

    #img_label = Label(image=photo)
    #img_label.pack()
    label.pack()
    a_root.after(7000, lambda: a_root.destroy()) # Destroy the widget after 30 seconds
    a_root.mainloop()

def warnBUFABandDizparc(results, username):
    for orgCheck in results:
        if orgCheck['organization'] == '2020' or orgCheck['organization'] == '2021' or orgCheck['organization'] == '49' or orgCheck['organization'] == '2022' or orgCheck['organization'] == '2019' or orgCheck['organization'] == '2023' and username=='nils.myhrman':
            return True
        if(orgCheck['organization']== '25' and username=='nils.myhrman'):
            print('true')
            return True
    return False

def get_test_or_prod_site(settings_data, username, password):
    if settings_data['use_production_site'] == True:
        return f"{username}:{password}@jeeves.custhelp.com"
    else:
        return f"{username}:{password}@jeeves--tst1.custhelp.com"

    # SCM = 8
            # TECH = 7
            # FINANCE = 6
            # FIRSTLINE = 2

''' GLOBAL VARIABLE LOL CUZ WHO CARES'''


sev_ones_sent = list()


def main():
    try:
        time_warn = time()
        settings_data = settings_json()
        print("""
 __       __                                                   
/  \     /  |                                                  
$$  \   /$$ |  ______    ______    _______  __    __   _______ 
$$$  \ /$$$ | /      \  /      \  /       |/  |  /  | /       |
$$$$  /$$$$ | $$$$$$  |/$$$$$$  |/$$$$$$$/ $$ |  $$ |/$$$$$$$/ 
$$ $$ $$/$$ | /    $$ |$$ |  $$/ $$ |      $$ |  $$ |$$      \ 
$$ |$$$/ $$ |/$$$$$$$ |$$ |      $$ \_____ $$ \__$$ | $$$$$$  |
$$ | $/  $$ |$$    $$ |$$ |      $$       |$$    $$/ /     $$/ 
$$/      $$/  $$$$$$$/ $$/        $$$$$$$/  $$$$$$/  $$$$$$$/

marcus.lundgren@jeeves.se
Senast uppdaterad: 2022-03-17
--Senaste Ändringar--
    1. Fixade r10+ kärnor
    2. Fixade så att support@jeeves.se inte kommer med.
    3. Fixade så att hantering av kärnor med Rxx kan hanteras.
--OBS!--
Se till att sätta på notifikationer i Windows för att se varningar.
--Kända buggar--
    1. CC sätts inte på ärenden som skapats upp från gamla ärenden som "gått ut" och nytt ärende skapats upp"""
        )
        if datetime.datetime.today().weekday() == 4:
            print('Idag är det fredag! Wohooooo!')
            print('''
            ██▀█▀█████████████▀█▀█████
            ██─▀─█─▄─█─▄─█─▄─█─▀─█████
            ██─█─█─▄─█─▄▄█─▄▄██─██████
            ███▀█▀████▄███▄████▄██████
            ██─────███████████████████
            ██▄───▄███████▀▀▀█─▀─█████
            ████▄███▀█─▄▀█─▀─██─██████
            █─▄█─▄▀█─█─▀▄█▄█▄█████████
            █─▄█─▄▀█▄█████████████████
            █▄████████████████████████\n''')

        username = input(r'användare: ')
        password = getpass(r'lösenord: ')
        if(username == 'sara.ekstrom' or username == 'marcus.lundgren'):
            print('Laddar bilderna')
            text_file = urllib.request.urlopen("https://dandelion-silky-lodge.glitch.me/raya.txt")
            asd = text_file.read()
            d = asd.split(b"\n")
            msg = base64.b64decode(random.choice(d))
            print(msg)
            buf = BytesIO(msg)
            img = Image.open(buf)
            width, height = img.size
            a_root = Tk()
            a_root.attributes('-topmost',True)
            a_root.focus_force()

            a_root.title("❤️❤️❤️ RAYABAJA ❤️❤️❤️")
            a_root.geometry(f"{width}x{height}")
            a_root.maxsize(1920, 1080)


            photo = ImageTk.PhotoImage(img)
            label = Label(a_root, text='', image=photo, compound='center')

            #img_label = Label(image=photo)
            #img_label.pack()
            label.pack()
            a_root.mainloop()
            
        rn_client = rn_start(username, password)
        print(f"Bevakar {convert_to_queue(settings_data['queue'])}")
        print(f"Bevakningen sker nu var {settings_data['time'] / 60} min")
        if settings_data["warn_if_critical"] == "Y":
            print(f"Notifikation kommer varna om Critical-ärenden")
        else:
            print(f"Notifikation kommer INTE varna om Critical-ärenden")
        print(f'\n------------------------------------\n')
        while True:
            seconds = time()
            results = OSvCPythonQueryResults().query(
                query=f"SELECT DISTINCT I.ID, I.lookupName, I.product, min(I.threads.createdTime), I.category, I.subject, I.severity, I.organization, I.threads.text, I.threads.mailHeader, I.queue, I.statusWithType.*, I.assignedTo.*, I.customFields.c.app_version,I.customFields.c.kernel_version,I.customFields.c.product_version, I.customFields.c.product_other, I.customFields.c.inc_types, I.customFields.c.inc_types_jeeves_portals from INCIDENTS I where I.queue=2 and I.assignedTo.account IS NULL and I.statusWithType.status=1 group by lookUpName",
                client=rn_client,
                annotation="Custom annotation",
                exclude_null=True
            )


            #I.queue=2 and I.assignedTo.account IS NULL and I.statusWithType.status=1 

            if 'detail' in results:
                print('Fel anv/lösen förmodligen! Byt i JSON-filen eller kontakta Marcus')
                sleep(60)
                sys.exit()
            else:

            #query="SELECT I.lookupName, I.severity, I.subject, I.queue, I.statusWithType.*, I.assignedTo.* from INCIDENTS I where I.queue=8 and I.statusWithType.status=1 and I.assignedTo.account IS NULL",

            # STATUS NEW = 1
            # staffGroup = None
            # SCM = 8
            # TECH = 7
            # FINANCE = 6
            # FIRSTLINE = 2
                incidents_to_add_full = check_if_checked(results)
                incident_to_add_idonly = list(fix_to_list(incidents_to_add_full[0]))

                if len(incidents_to_add_full[0]) > 0:
                    add_to_text(incident_to_add_idonly, incidents_to_add_full[0], settings_data, rn_client, username, password)
                    checkToShowToaster = warnBUFABandDizparc(results, username)
                    if(checkToShowToaster == False):
                        show_toaster(settings_data['toast_subject_message'],settings_data['toast_message'], incidents_to_add_full[1], settings_data['warn_if_critical'], settings_data['time'])
                    else:
                        tkShowFahlStrom()

                else:
                    print(f"INGET NYTT I KÖN - KOLLAR IGEN OM {settings_data['time'] / 60} min")



                warn_for_severity_ones(rn_client, False, True);
                #funktion för att varna om severity one
                #om det har gått en viss tid, så kollar vi om någon har tagit 1orna
                #om ingen har gjort det, varna igen
                if time() - time_warn > 600:
                    print('Kollar om Team:en tagit deras 1or än')
                    warned = warn_for_severity_ones(rn_client, True, False)
                    time_warn = time()


                sleep(settings_data['time'])
    except Exception:
        print(traceback.format_exc())
        sleep(120)
        show_toaster("KRASCH", "STARTA OM PLZ", [], "N")
main()















