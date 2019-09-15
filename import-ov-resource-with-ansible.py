###
# Copyright (2016-2019) Hewlett Packard Enterprise Development LP
#
# Licensed under the Apache License, Version 2.0 (the "License");
# You may not use this file except in compliance with the License.
# You may obtain a copy of the License at
#
# http://www.apache.org/licenses/LICENSE-2.0
#
# Unless required by applicable law or agreed to in writing, software
# distributed under the License is distributed on an "AS IS" BASIS,
# WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
# See the License for the specific language governing permissions and
# limitations under the License.
###
### -----------------------History
##    04-SEP : Works up to LIG with Ethernet-FC
##    06-SEP :
##      - Scopes : ethernet-fc- lig - network set
##      - Enclosure group
##   08-SEP:
##      - Logical enclosure
##      - Update firmware
##      - Scopes to LE and EG
##   10-SEP: 
##      - server profiles



from pprint import pprint
import json
import copy
import csv

import os
from os import sys


from hpOneView.exceptions import HPOneViewException
from hpOneView.oneview_client import OneViewClient

import pandas as pd
from pandas import ExcelWriter
from pandas import ExcelFile

TABSPACE        = "    "
COMMA           = ','
CR              = '\n'
CRLF            = '\r\n'
IC_OFFSET       = 3         # Used to calculate bay number from InterconnectBaySet
                            # InterconnectBaySet = 1 ---> Bay 1  and Bay 4
                            # InterconnectBaySet = 2 ---> Bay 1  and Bay 5
                            # InterconnectBaySet = 3 ---> Bay 1  and Bay 6


# Definition of resource types
resource_type_ov4_20    = {
    'subnet'                        : 'Subnet',
    'range'                         : 'Range',
    'timeandlocale'                 : 'TimeAndLocale',
    'scope'                         : 'ScopeV3',
# network
    'ethernet'                      : 'ethernet-networkV4',
    'fcnetwork'                     : 'fc-networkV4',
    'fcoenetwork'                   : 'fcoe-networkV4',
    'ethernetsettings'              : 'EthernetInterconnectSettingsV4',
    'networkset'                    : 'network-setV4',
    'logicalinterconnectgroup'      : 'logical-interconnect-groupV5',
    'enclosuregroup'                : 'EnclosureGroupV7',
# storage
    'storagevolumetemplate'         : 'StorageVolumeTemplateV6',
    'storagevolume'                 : 'StorageVolumeV7',
# server
    'logicalenclosure'              : 'LogicalEnclosureV4',
    'serverprofiletemplate'         : 'ServerProfileTemplateV5',
    'serverprofile'                 : 'ServerProfileV9',
    'clusterprofile'                : 'HypervisorClusterProfileV3'
}

# Defintion of UplinkSet Ethernet Network type
uplinkSetEthNetworkType     = {
    'Ethernet'                      : 'Tagged',
    'Untagged'                      : 'Untagged',
    'Tunnel'                        : 'Tunnel',
    'ImageStreamer'                 : 'ImageStreamer'
}


# Defintion of UplinkSet Network type
uplinkSetNetworkType        = {
    'Ethernet'                      : 'Ethernet',
    'FibreChannel'                  : 'FibreChannel',
    'Untagged'                      : 'Ethernet',
    'Tunnel'                        : 'Ethernet',
    'ImageStreamer'                 : 'Ethernet'
}

# ================================================================================================
#
#   Build playbook headers
#
# ================================================================================================
def build_header(scriptCode):

    
    scriptCode.append("  hosts: localhost"                                                                                                          )       
    scriptCode.append("  vars:"                                                                                                                     )
    scriptCode.append("     config: \'oneview_config.json\'"                                                                                        )



# ================================================================================================
#
#   Write to file
#
# ================================================================================================
def write_to_file(scriptCode, filename):
    file                        = open(filename, "w")
    code                        = CR.join(scriptCode)
    file.write(code)

    file.close()

# ================================================================================================
#
#   Sort CSV based on column number
#
# ================================================================================================
def sort_csv(csvFile, column=0):

    ifile       = open(csvFile, 'r')
    reader      = csv.reader(ifile)

    header      = next(reader)
    sortedList  = sorted(reader)
    ifile.close()

    with open(csvFile, 'w') as f:
        writer = csv.writer(f)
        writer.writerow(header) 
        for row in sortedList:
            if row:
                writer.writerow(row)

    f.close()

# ================================================================================================
#
#  find_port_number_in_interconnect_type
#
# ================================================================================================
def find_port_number_in_interconnect_type(allInterconnectTypes, icName,  portName):

    portNumber          = 0
    this_ic             = None

    if allInterconnectTypes:
        # find interconnect type by name
        for ic in allInterconnectTypes:
            if icName == ic['name']:
                this_ic = ic


        if this_ic:
            # find portInfos of this ic
            portInfos       = this_ic['portInfos']
            for p in portInfos:
                if portName == p['portName'] :
                    portNumber = p['portNumber']
    
    return portNumber

# ================================================================================================
#
#   HELPER: generate_id_pools_ipv4_subnets
#
# ================================================================================================

def generate_id_pools_ipv4_subnets(values_in_dict,scriptCode):  
    # ---- Note:
    #       This is to define only common code


    subnet                      = values_in_dict
    name                        = subnet['name']
    networkId                   = subnet['networkId']   
    subnetmask                  = subnet['subnetmask'] 
    gateway                     = subnet['gateway']
    domain                      = subnet['domain']

    scriptCode.append("                                                     "                                                                       )
    scriptCode.append("     - name: Create subnet                       "                                                                           )
    scriptCode.append("       oneview_id_pools_ipv4_subnet:             "                                                                           )
    scriptCode.append("         config:     \'{{config}}\'              "                                                                           )
    scriptCode.append("         state:      present                     "                                                                           )
    scriptCode.append("         data:                                   "                                                                           )
    scriptCode.append("             name:                            {} ".format(name)                                                              )
    scriptCode.append("             networkId:                       {} ".format(networkId)                                                         )
    scriptCode.append("             subnetmask:                      {} ".format(subnetmask)                                                        )
    scriptCode.append("             gateway:                         {} ".format(gateway)                                                           )
    scriptCode.append("             domain:                          {} ".format(domain)                                                            )
    scriptCode.append("             type:                            {} ".format(rstype['subnet'])                                                  )
    

    return scriptCode


# ================================================================================================
#
#   HELPER: generate_id_pools_ipv4_ranges
#
# ================================================================================================

def generate_id_pools_ipv4_ranges(values_in_dict,scriptCode):  
    # ---- Note:
    #       This is to define only common code

    pool                        = values_in_dict
    name                        = pool['name']
    startAddress                = pool['startAddress']
    endAddress                  = pool['endAddress']

    scriptCode.append("                                                     "                                                                           )
    scriptCode.append("     - name: Create Id pools                     "                                                                           )
    scriptCode.append("       oneview_id_pools_ipv4_range: "                                                                           )
    scriptCode.append("         config:     \'{{config}}\'              "                                                                           )
    scriptCode.append("         state:      present                     "                                                                           )
    scriptCode.append("         data:                                   "                                                                           )
    scriptCode.append("             name:                         {}    ".format(name)                                                            )
    scriptCode.append("             enabled:                      True  "                                                         )
    scriptCode.append("             type:                         {}    ".format(rstype['range'])                                                          )
    scriptCode.append("             startAddress:                \'{}\' ".format(startAddress)                                                            )
    scriptCode.append("             endAddress:                  \'{}\' ".format(endAddress)                                                            )


    return scriptCode

# ================================================================================================
#
#   generate_id_pools_ipv4_ranges_subnets_ansible_script_from_csv
#
# ================================================================================================
def generate_id_pools_ipv4_ranges_subnets_ansible_script_from_csv(sheet, to_file):
    
    print('Creating ansible playbook   =====>           {}'.format(to_file))    
    scriptCode = []
    scriptCode.append("---"                                                                                                                         )
    scriptCode.append("- name:  Configure id pools ipv4 from csv"                                                                                     )    
    build_header(scriptCode)

    #####
    scriptCode.append("  tasks:"                                                                                                                    )
    sheet.dropna(how='all', inplace=True)
    for i in sheet.index:
        row                         = sheet.loc[i]
        name                        = row['name']
        startAddress                = row['startAddress']
        endAddress                  = row['endAddress']
        poolType                    = row['poolType']
        if 'IPV4' in poolType:
            generate_id_pools_ipv4_subnets(row, scriptCode)
            dnsServers                  = row['dnsServers'] 
            dnsServers                  = dnsServers.split('|')
            if dnsServers:
                scriptCode.append("             dnsServers:                         " )
                for dns in dnsServers:
                    scriptCode.append("                 - {}                            ".format(dns) )  
    
            scriptCode.append("")
            scriptCode.append("                                                     "                                                                           )
            scriptCode.append("     - name: Get uri for subnet from {}          ".format(name)                                                                           )
            scriptCode.append("       oneview_id_pools_ipv4_subnet_facts: "                                                                           )
            scriptCode.append("         config:     \'{{config}}\'              "                                                                           )
            scriptCode.append("         name:       \'{}\'                      ".format(name)                                                                           )
            scriptCode.append("     - set_fact:  " )
            var_name            = name.lower().replace('-','_').strip(' ')                         
            scriptCode.append("         subnet_{}_uri:".format(var_name) + " \'{{id_pools_ipv4_subnets[0].uri}}\'")

            generate_id_pools_ipv4_ranges(row,scriptCode)
            scriptCode.append("             subnetUri:                   \'{{" + "subnet_{}_uri".format(var_name)  + "}}\'"                 )



    # end of id pools
    scriptCode.append("       delegate_to: localhost                    "                                                                   )
    scriptCode.append(CR)


    # ============= Write scriptCode ====================
    write_to_file(scriptCode, to_file)



# ================================================================================================
#
#   HELPER: generate_time_locale
#
# ================================================================================================

def generate_time_locale(values_in_dict,scriptCode):  
    # ---- Note:
    #       This is to define only common code


    time                        = values_in_dict
    locale                      = time['locale']
    timezone                    = time['timezone']   
    ntpServers                  = time['ntpServers'] # []

    scriptCode.append("                                                     "                                                               )
    scriptCode.append("     - name: Create time and locale              "                                                                           )
    scriptCode.append("       oneview_appliance_time_and_locale_configuration: "                                                                    )
    scriptCode.append("         config: \'{{config}}\'                  "                                                                           )
    scriptCode.append("         state: present                          "                                                                           )
    scriptCode.append("         data:                                   "                                                                           )
    scriptCode.append("             locale:                          {} ".format(locale)                                                            )
    scriptCode.append("             timezone:                        {} ".format(timezone)                                                          )
    scriptCode.append("             type:                            {} ".format(rstype['timeandlocale'])                                                                   )
    





    

    return scriptCode

# ================================================================================================
#
#   generate_time_locale_ansible_script_from_csv
#
# ================================================================================================
def generate_time_locale_ansible_script_from_csv(sheet, to_file):
    
    print('Creating ansible playbook   =====>           {}'.format(to_file))    
    scriptCode = []
    scriptCode.append("---"                                                                                                                             )
    scriptCode.append("- name:  Configure time and locale from csv"                                                                                     )    
    build_header(scriptCode)

    scriptCode.append("  tasks:"                                                                                                                    )
    sheet.dropna(how='all', inplace=True)
    for i in sheet.index:
        row                         = sheet.loc[i]
        generate_time_locale(row, scriptCode)
        ntpServers                  = row['ntpServers'] 
        ntpServers                  = str(ntpServers).split('|')
        if ntpServers:
            scriptCode.append("             ntpServers:                 "                                                                           )
            for ntp in ntpServers:
                scriptCode.append("                 - {}                ".format(ntp)                                                               )  
    # end of time and locale
        scriptCode.append("       delegate_to: localhost                    "                                                                           )
        scriptCode.append(CR)
    #print(CR.join(scriptCode))

    # ============= Write scriptCode ====================
    write_to_file(scriptCode, to_file)

##
# ================================================================================================
#
#   generate_scope_for_resource
#
# ================================================================================================
def generate_scope_for_resource(name, varNameUri, scope, scriptCode):

    list_of_scopes      = scope.split('|')
    for sc in list_of_scopes:
        scope_name      = sc
        scriptCode.append("                                                         "                                                   )
        scriptCode.append("     - name: Update the scope {0} with new resource {1}  ".format(scope_name,name)                           )
        scriptCode.append("       oneview_scope:                                    "                                                   )
        scriptCode.append("         config:       '{{ config }}'                    "                                                   )
        scriptCode.append("         state:        resource_assignments_updated      "                                                   )
        scriptCode.append("         data:                                           "                                                   )
        scriptCode.append("             name:     {}                                ".format(scope_name)                                )
        scriptCode.append("             resourceAssignments:                        "                                                   )
        scriptCode.append("                 addedResourceUris:                      "                                                   )
        scriptCode.append("                     - {}                                ".format(varNameUri)                                    )    
##

# ================================================================================================
#
#   generate_firmware_bundle_ansible_script_from_csv
#
# ================================================================================================
def generate_firmware_bundle_ansible_script_from_csv(sheet, to_file):
    
    print('Creating ansible playbook   =====>           {}'.format(to_file))    
    scriptCode = []
    scriptCode.append("---"                                                                                                                             )
    scriptCode.append("- name:  Configure firmware bundle from csv"                                                                                     )    
    build_header(scriptCode)

    scriptCode.append("  tasks:"                                                                                                                        )
    sheet.dropna(how='all', inplace=True)
    for i in sheet.index:
        row                         = sheet.loc[i]
        name                        = row['name'] 
        filename                    = row['filename']

        scriptCode.append("                                                     "                                                                   )
        scriptCode.append("     - name: Upload firmware bundle   {}         ".format(name)                                                          )
        scriptCode.append("       oneview_firmware_bundle:                  "                                                                       )
        scriptCode.append("         config:         \'{{config}}\'          "                                                                       )
        scriptCode.append("         state: present                          "                                                                       )
        scriptCode.append("         file_path:      \'{}\'                  " .format(filename)                                                     )


    
    # end of firmware bundle
    scriptCode.append("       delegate_to: localhost                            "                                                                       )
    scriptCode.append(CR)

    # ============= Write scriptCode ====================
    write_to_file(scriptCode, to_file)



# ================================================================================================
#
#   generate_snmp_v1_ansible_script_from_csv
#
# ================================================================================================
def generate_snmp_v1_ansible_script_from_csv(sheet, to_file):
    
    print('Creating ansible playbook   =====>           {}'.format(to_file))    
    scriptCode = []
    scriptCode.append("---"                                                                                                                             )
    scriptCode.append("- name:  Configure snmp v1  from csv"                                                                                            )    
    build_header(scriptCode)

    scriptCode.append("  tasks:"                                                                                                                    )
    sheet.dropna(how='all', inplace=True)
    for i in sheet.index:
        row             = sheet.loc[i]
        destination                 = row['destination'] 
        communityString             = row['communityString']
        port                        = row['port']

        scriptCode.append("                                                     "                                                                           )
        scriptCode.append("     - name: Create trap destination {}         ".format(destination)                                                    )
        scriptCode.append("       oneview_appliance_device_snmp_v1_trap_destinations:                  "                                            )
        scriptCode.append("         config:                 \'{{config}}\'  "                                                                       )
        scriptCode.append("         state: present                          "                                                                       )
        scriptCode.append("         data:                                   "                                                                       )
        scriptCode.append("             communityString:    \'{}\'          " .format(communityString)                                              )
        scriptCode.append("             destination:        \'{}\'          " .format(destination)                                                  )
        scriptCode.append("             port:               \'{}\'          " .format(port)                                                         )
    
    
            # end of snmp_v1
    scriptCode.append("       delegate_to: localhost                            "                                                                       )
    scriptCode.append(CR)

    # ============= Write scriptCode ====================
    write_to_file(scriptCode, to_file)



# ================================================================================================
#
#   generate_ansible_configuration
#
# ================================================================================================
def generate_ansible_configuration(composerSheet, versionSheet, to_file):
     
    print('Creating ansible config     =====>           {}'.format(to_file)) 

    scriptCode = []
    composerSheet.dropna(how='all', inplace=True)
    row = composerSheet.iloc[0]
    Ip                       = row['Ip'].strip() 
    userName                 = row['userName'].strip()
    password                 = row['password'].strip()
    api_version              = row['api_version'].strip()
    scriptCode.append("{                                         "                                                                                                                           )
    scriptCode.append("     \"ip\":              \"{}\",         " .format(Ip)                                                                      )
    scriptCode.append("     \"credentials\" :    {               "                                                                                  )
    scriptCode.append("         \"userName\":    \"{}\",         " .format(userName)                                                                )
    scriptCode.append("         \"password\":    \"{}\"          " .format(password)                                                                )
    scriptCode.append("      },                                  "                                                                                  )
    scriptCode.append("     \"api_version\" :     \"{}\"         " .format(api_version)                                                             )
    scriptCode.append("}                                         "                                                                                  )

    scriptCode.append(CR)

    # ============= Write scriptCode ====================
    write_to_file(scriptCode, to_file)


    # ============== Generate prefix =====================
    versionSheet.dropna(how='all', inplace=True)
    row = versionSheet.iloc[0]
    pod                 = row['Pod']
    site                = row['Site']
    if pod:
        pod             = pod.lower().strip()
    if site:
        site            = site.lower().replace(',', '-')
    
    prefix              = site + '-' + pod + '-'

    return prefix

# ================================================================================================
#
#   generate_scopes_ansible_script_from_csv
#
# ================================================================================================

def generate_scopes_ansible_script_from_csv(sheet, to_file):


    print('Creating ansible playbook   =====>           {}'.format(to_file))
    scriptCode = []
    scriptCode.append("---"                                                                                                                     )
    scriptCode.append("- name:  Configure scopes from csv"                                                                                      )    
    build_header(scriptCode)


    scriptCode.append("  tasks:"                                                                                                                )
    sheet.dropna(how='all', inplace=True)
    sheet                       = sheet.applymap(str)                       # Convert data frame into string
    for i in sheet.index:
        row                     = sheet.loc[i]
        name                    = row["name"]
        description             = row["description"]

        scriptCode.append("                                                     "                                                               )
        scriptCode.append("     - name: Create scope  {}                    ".format(name)                                                      )
        scriptCode.append("       oneview_scope:                            "                                                                   )
        scriptCode.append("         config: \'{{config}}\'                  "                                                                   )
        scriptCode.append("         state: present                          "                                                                   )
        scriptCode.append("         data:                                   "                                                                   )
        scriptCode.append("             type:                       \'{}\'  ".format(rstype['scope'])                                           )
        scriptCode.append("             name:                       \'{}\'  ".format(name)                                                      )
        if 'nan' != description:
            scriptCode.append("             description:                \'{}\'  ".format(description)                                           )
        
    

    # end of scopes
    scriptCode.append("       delegate_to: localhost                    "                                                                       )
    scriptCode.append(CR) 

    # ============= Write scriptCode ====================
    write_to_file(scriptCode, to_file)



# ================================================================================================
#
#   HELPER: generate_ethernet_networks_ansible_script_from
#
# ================================================================================================

def generate_ethernet_networks(values_in_dict,scriptCode):

    #_________________________ NOT USED __________________________________________
    net                     = values_in_dict
    name                    = net["name"]
    if pd.notnull(net['description']):
        description         = net["description"]
    else:
        description         = ""
    purpose                 = net["purpose"]
    vlanId                  = net["vlanId"]
    smartLink               = net["smartLink"]
    privateNetwork          = net["privateNetwork"]
    ethernetNetworkType     = net['ethernetNetworkType']
    typicalBandwidth        = net['typicalBandwidth']
    maximumBandwidth        = net['maximumBandwidth']

    smartLink               = smartLink.lower().capitalize()
    privateNetwork          = privateNetwork.lower().capitalize()

    scriptCode.append("                                                     "                                                               )
    scriptCode.append("     - name: Create ethernet network {}          ".format(name)                                                      )
    scriptCode.append("       oneview_ethernet_network:                 "                                                                   )
    scriptCode.append("         config: \'{{config}}\'                  "                                                                   )
    scriptCode.append("         state: present                          "                                                                   )
    scriptCode.append("         data:                                   "                                                                   )
    scriptCode.append("             type:                       \'{}\'  ".format(rstype['ethernet'])                                                    )
    scriptCode.append("             name:                       \'{}\'  ".format(name)                                                      )
    if 'nan' not in description:
        scriptCode.append("             description:                \'{}\'  ".format(description)                                           )
    scriptCode.append("             ethernetNetworkType:        {}      ".format(ethernetNetworkType)                                       )
    scriptCode.append("             purpose:                    {}      ".format(purpose)                                                   )
    scriptCode.append("             smartLink:                  {}      ".format(smartLink)                                                 )
    scriptCode.append("             privateNetwork:             {}      ".format(privateNetwork)                                            )
    scriptCode.append("             vlanId:                     {}      ".format(vlanId)                                                    )
    scriptCode.append("             bandwidth:                          "                                                                   )
    scriptCode.append("                 typicalBandwidth:       {}      ".format(typicalBandwidth)                                          )
    scriptCode.append("                 maximumBandwidth:       {}      ".format(maximumBandwidth)                                          )


    return scriptCode


# ================================================================================================
#
#   generate_ethernet_networks_ansible_script_from_csv
#
# ================================================================================================
def generate_ethernet_networks_ansible_script_from_csv(sheet, to_file):

    print('Creating ansible playbook   =====>           {}'.format(to_file))
    scriptCode = []
    scriptCode.append("---"                                                                                                                     )
    scriptCode.append("- name:  Configure Ethernet networks from csv"                                                                           )    
    build_header(scriptCode)


    scriptCode.append("  tasks:"                                                                                                                )
    sheet.dropna(how='all', inplace=True)
    sheet                       = sheet.applymap(str)                       # Convert data frame into string
    for i in sheet.index:
        row                     = sheet.loc[i]
        name                    = row["name"]
        description             = row["description"]
        purpose                 = row["purpose"]
        vlanId                  = row["vlanId"]
        smartLink               = row["smartLink"]
        privateNetwork          = row["privateNetwork"]
        ethernetNetworkType     = row['ethernetNetworkType']
        typicalBandwidth        = row['typicalBandwidth']
        maximumBandwidth        = row['maximumBandwidth']
        scope                   = row['scope']


        smartLink               = smartLink.lower().capitalize()
        privateNetwork          = privateNetwork.lower().capitalize()
    
        scriptCode.append("                                                     "                                                               )
        scriptCode.append("     - name: Create ethernet network {}          ".format(name)                                                      )
        scriptCode.append("       oneview_ethernet_network:                 "                                                                   )
        scriptCode.append("         config: \'{{config}}\'                  "                                                                   )
        scriptCode.append("         state: present                          "                                                                   )
        scriptCode.append("         data:                                   "                                                                   )
        scriptCode.append("             type:                       \'{}\'  ".format(rstype['ethernet'])                                        )
        scriptCode.append("             name:                       \'{}\'  ".format(name)                                                      )
        if 'nan' != description:
            scriptCode.append("             description:                \'{}\'  ".format(description)                                           )
        scriptCode.append("             ethernetNetworkType:        {}      ".format(ethernetNetworkType)                                       )
        scriptCode.append("             purpose:                    {}      ".format(purpose)                                                   )
        scriptCode.append("             smartLink:                  {}      ".format(smartLink)                                                 )
        scriptCode.append("             privateNetwork:             {}      ".format(privateNetwork)                                            )
        scriptCode.append("             vlanId:                     {}      ".format(vlanId)                                                    )
        scriptCode.append("             bandwidth:                          "                                                                   )
        scriptCode.append("                 typicalBandwidth:       {}      ".format(typicalBandwidth)                                          )
        scriptCode.append("                 maximumBandwidth:       {}      ".format(maximumBandwidth)                                          )
    
        # Add scope here
        
        if 'nan' != scope:
            netVar              = 'var_' + name.lower().strip().replace(',', '_').replace('-', '_').replace(' ', '') 
            scriptCode.append("                                                     "                                                           )
            scriptCode.append("     - name: get ethernet network {}             ".format(name)                                                  )
            scriptCode.append("       oneview_ethernet_network_facts:           "                                                               )
            scriptCode.append("         config:     \'{{config}}\'              "                                                               )
            scriptCode.append("         name:       \'{}\'                      ".format(name)                                                  )
            scriptCode.append("     - set_fact:                                 "                                                               )
            scriptCode.append("          {}: ".format(netVar)  + "\'{{ethernet_networks.uri}}\' "                                               )
            netUri              = "\'{{" + '{}'.format(netVar) + "}}\'"  
            generate_scope_for_resource(name, netUri, scope, scriptCode)
            



    # end of ethernet networks
    scriptCode.append("       delegate_to: localhost                    "                                                                       )
    scriptCode.append(CR) 

    # ============= Write scriptCode ====================
    write_to_file(scriptCode, to_file)




# ================================================================================================
#
#   HELPER: generate_network_sets
#
# ================================================================================================

def generate_network_sets(values_in_dict,scriptCode):  
    # ---- Note:
    #       The helper function just generates code that are common to both csv and oneview
    #
#_________________________ NOT USED __________________________________________
    netset                  = values_in_dict
    name                    = netset["name"]

    if pd.notnull(netset['description']):
        description         = netset["description"]
    else:
        description          = ""
    
        nativeNetworkUri     = netset["nativeNetworkUri"]

    scriptCode.append("                                                     "                                                                       )
    scriptCode.append("     - name: Create network set {}               ".format(name)                                                              )
    scriptCode.append("       oneview_network_set:                      "                                                                           )
    scriptCode.append("         config: \'{{config}}\'                  "                                                                           )
    scriptCode.append("         state: present                          "                                                                           )
    scriptCode.append("         data:                                   "                                                                           )
    scriptCode.append("             type:                       \'{}\'  ".format(rstype['networkset'])                                                             )
    scriptCode.append("             name:                       \'{}\'  ".format(name)                                                              )
    if description:
        scriptCode.append("             description:                {}  ".format(description)                                                       )


    return scriptCode




# ================================================================================================
#
#   generate_network_sets_ansible_script_from_csv
#
# ================================================================================================
def generate_network_sets_ansible_script_from_csv(sheet, to_file):

    print('Creating ansible playbook   =====>           {}'.format(to_file))    
    scriptCode = []
    scriptCode.append("---"                                                                                                                         )
    scriptCode.append("- name:  Configure network sets from csv"                                                                                    )    
    build_header(scriptCode)


    scriptCode.append("  tasks:"                                                                                                                    )

    sheet.dropna(how='all', inplace=True)
    sheet                       = sheet.applymap(str)                       # Convert data frame into string

    for i in sheet.index:
        row                     = sheet.loc[i]
        #generate_network_sets(row,scriptCode) # get common code first
        ##
        name                    = row['name']
        description             = row["description"]
        networkUris             = row["networkUris"]
        typicalBandwidth        = row['typicalBandwidth']
        maximumBandwidth        = row['maximumBandwidth']
        scope                   = row['scope']
    
        scriptCode.append("                                                     "                                                                       )
        scriptCode.append("     - name: Create network set {}               ".format(name)                                                              )
        scriptCode.append("       oneview_network_set:                      "                                                                           )
        scriptCode.append("         config: \'{{config}}\'                  "                                                                           )
        scriptCode.append("         state: present                          "                                                                           )
        scriptCode.append("         data:                                   "                                                                           )
        scriptCode.append("             type:                       \'{}\'  ".format(rstype['networkset'])                                                               )
        scriptCode.append("             name:                       \'{}\'  ".format(name)                                                              )
        if 'nan' != description:
            scriptCode.append("             description:                {}  ".format(description)                                                       )
        if 'nan' != networkUris:
            networks = networkUris.split('|')

            if networks: 
                scriptCode.append("             networkUris:                    "                                                                       )
                for net in networks:
                        scriptCode.append("                 - {}                ".format(net)                                                           )   


        # Add scope here
        
        if 'nan' != scope:
            netVar              = 'var_' + name.lower().strip().replace(',', '_').replace('-', '_').replace(' ', '')
            scriptCode.append("                                                     "                                                           )
            scriptCode.append("     - name: get network set {}             ".format(name)                                                       )
            scriptCode.append("       oneview_network_set_facts:           "                                                                    )
            scriptCode.append("         config:     \'{{config}}\'              "                                                               )
            scriptCode.append("         name:       \'{}\'                      ".format(name)                                                  )
            scriptCode.append("     - set_fact:                                 "                                                               )
            scriptCode.append("          {}: ".format(netVar)  + "\'{{network_sets[0].uri}}\' "                                                 )
            netUri              = "\'{{" + '{}'.format(netVar) + "}}\'"  
            generate_scope_for_resource(name, netUri, scope, scriptCode)

 




    scriptCode.append("       delegate_to: localhost                    "                                                                               )
    scriptCode.append(CR                                                                                                                                )

# Bandwidth not working - Need to fin connection templates and set it
#scriptCode.append("             bandwidth:                          "                                                                   )
#scriptCode.append("                 typicalBandwidth:       {}      ".format(typicalBandwidth)                                          )
#scriptCode.append("                 maximumBandwidth:       {}      ".format(maximumBandwidth)                                          )







    #print('Creating ansible playbook   =====>           {}'.format(to_file))

    # ============= Write scriptCode ====================
    write_to_file(scriptCode, to_file)



# ================================================================================================
#
#   generate_fc_fcoe_networks_ansible_script_from_csv
#
# ================================================================================================
def generate_fc_fcoe_networks_ansible_script_from_csv(sheet, to_file):

    print('Creating ansible playbook   =====>           {}'.format(to_file))
    scriptCode = []
    scriptCode.append("---"                                                                                                                     )
    scriptCode.append("- name:  Configure fc/fcoe networks from csv"                                                                           )    
    build_header(scriptCode)


    scriptCode.append("  tasks:"                                                                                                                )
    sheet.dropna(how='all', inplace=True)
    sheet                       = sheet.applymap(str)                       # Convert data frame into string
    for i in sheet.index:
        row                     = sheet.loc[i]
        name                    = row["name"]
        description             = row["description"]
        autoLoginRedistribution = row["autoLoginRedistribution"]
        fabricType              = row["fabricType"]
        linkStabilityTime       = row["linkStabilityTime"]
        managedSanUri           = row["managedSanUri"]
        typicalBandwidth        = row['typicalBandwidth']
        maximumBandwidth        = row['maximumBandwidth']
        _type                   = row['type']
        vlanId                  = row['vlanId']
        scope                   = row['scope']
    
        _type                   = _type.lower()
        if 'fc' == _type:
            autoLoginRedistribution = autoLoginRedistribution.lower()
            isAuto                  = 'auto' in autoLoginRedistribution
            if 'nan' == linkStabilityTime:
                linkStabilityTime   = 30    # default is 30 sec

            scriptCode.append("                                                 "                                                                   )
            scriptCode.append("     - name: Create fc network {}                ".format(name)                                                      )
            scriptCode.append("       oneview_fc_network:                       "                                                                   )
            scriptCode.append("         config: \'{{config}}\'                  "                                                                   )
            scriptCode.append("         state: present                          "                                                                   )
            scriptCode.append("         data:                                   "                                                                   )
            scriptCode.append("             type:                       \'{}\'  ".format(rstype['fcnetwork'])                                       )
            scriptCode.append("             name:                       \'{}\'  ".format(name)                                                      )
            if 'nan' != description:
                scriptCode.append("             description:                \'{}\'  ".format(description)                                           )
            scriptCode.append("             autoLoginRedistribution:    {}      ".format(isAuto)                                                    )
            scriptCode.append("             linkStabilityTime:          {}      ".format(linkStabilityTime)                                         )
            scriptCode.append("             fabricType:                 {}      ".format(fabricType)                                                )
            if 'nan' != managedSanUri:
                scriptCode.append("             managedSanUri:              \'{}\'  ".format(managedSanUri)                                         )

 #       scriptCode.append("             bandwidth:                          "                                                                   )
 #       scriptCode.append("                 typicalBandwidth:       {}      ".format(typicalBandwidth)                                          )
 #       scriptCode.append("                 maximumBandwidth:       {}      ".format(maximumBandwidth)                                          )
        
        else:
            scriptCode.append("                                                 "                                                                   )
            scriptCode.append("     - name: Create fcoe network {}              ".format(name)                                                      )
            scriptCode.append("       oneview_fcoe_network:                     "                                                                   )
            scriptCode.append("         config: \'{{config}}\'                  "                                                                   )
            scriptCode.append("         state: present                          "                                                                   )
            scriptCode.append("         data:                                   "                                                                   )
            scriptCode.append("             type:                       \'{}\'  ".format(rstype['fcoenetwork'])                                     )
            scriptCode.append("             name:                       \'{}\'  ".format(name)                                                      )
            if 'nan' != description:
                scriptCode.append("             description:                \'{}\'  ".format(description)                                           )
            scriptCode.append("             vlanId:                     {}      ".format(vlanId)                                                    )
            if 'nan' != managedSanUri:
                scriptCode.append("             managedSanUri:              \'{}\'  ".format(managedSanUri)                                         )

 #       scriptCode.append("             bandwidth:                          "                                                                   )
 #       scriptCode.append("                 typicalBandwidth:       {}      ".format(typicalBandwidth)                                          )
 #       scriptCode.append("                 maximumBandwidth:       {}      ".format(maximumBandwidth)                                          )            
    

        # Add scope here
        
        if 'nan' != scope:
            netVar              = 'var_' + name.lower().strip().replace(',', '_').replace('-', '_').replace(' ', '')
            scriptCode.append("                                                             "                                                   )
            scriptCode.append("     - name: get fc or fcoe network {}                       ".format(name)                                      )
            if 'fc' == _type:
                scriptCode.append("       oneview_fc_network_facts:                         "                                                   )
            else:
                scriptCode.append("       oneview_fcoe_network_facts:                       "                                                   )
            scriptCode.append("         config:     \'{{config}}\'                          "                                                   )
            scriptCode.append("         name:       \'{}\'                                  ".format(name)                                      )
            scriptCode.append("     - set_fact:                                             "                                                   )
            if 'fc' == _type:
                scriptCode.append("          {}: ".format(netVar)  + "\'{{fc_networks[0].uri}}\' "                                                 )
            else:
                scriptCode.append("          {}: ".format(netVar)  + "\'{{fcoe_networks[0].uri}}\' "                                               )
            netUri              = "\'{{" + '{}'.format(netVar) + "}}\'"  
            generate_scope_for_resource(name, netUri, scope, scriptCode)
        


    # end of fc networks
    scriptCode.append("       delegate_to: localhost                    "                                                                       )
    scriptCode.append(CR) 

    # ============= Write scriptCode ====================
    write_to_file(scriptCode, to_file)




##

# ================================================================================================
#
#   HELPER: generate_logical_interconnect_groups
#
# ================================================================================================

def generate_logical_interconnect_groups(values_in_dict,scriptCode, isFC=False, comeFromOV=False):  
    # ---- Note:
    #       This is to define only common code

    lig                         = values_in_dict
    name                        = lig["name"]
    description                 = lig["description"] 
    enclosureType               = lig["enclosureType"]

    interconnectBaySet          = lig["interconnectBaySet"]
    redundancyType              = lig["redundancyType"]
    if not redundancyType:
        redundancyType          = 'Redundant'    # default value

    # ethernetSettings
    if not isFC:
        if comeFromOV:
            ethernetSettings        = lig["ethernetSettings"]
            interconnectType        = ethernetSettings["interconnectType"]
            ethernetSettingsType    = ethernetSettings["type"]            
        else:
            ethernetSettings        = lig
            ethernetSettingsType    = 'EthernetInterconnectSettingsV4' 



        igmpSnooping                = ethernetSettings["enableIgmpSnooping"].lower().capitalize()
        igmpIdleTimeout             = ethernetSettings["igmpIdleTimeoutInterval"]
        networkLoopProtection       = ethernetSettings["enableNetworkLoopProtection"].lower().capitalize()
        pauseFloodProtection        = ethernetSettings["enablePauseFloodProtection"].lower().capitalize()
        enableRichTLV               = ethernetSettings["enableRichTLV"].lower().capitalize()
        taggedLldp                  = ethernetSettings["enableTaggedLldp"].lower().capitalize()
        lldpIpv6Address             = ethernetSettings["lldpIpv6Address"]
        lldpIpv4Address             = ethernetSettings["lldpIpv4Address"]
        fastMacCacheFailover        = ethernetSettings["enableFastMacCacheFailover"].lower().capitalize()
        macRefreshInterval          = ethernetSettings["macRefreshInterval"]



    scriptCode.append("                                                     "                                                                       )
    scriptCode.append("     - name: Create logical interconnect group {}".format(name)                                                              )
    scriptCode.append("       oneview_logical_interconnect_group:       "                                                                           )
    scriptCode.append("         config: \'{{config}}\'                  "                                                                           )
    scriptCode.append("         state: present                          "                                                                           )
    scriptCode.append("         data:                                   "                                                                           )
    scriptCode.append("             name:                        \'{}\' ".format(name)                                                              )
    scriptCode.append("             type:                        \'{}\' ".format(rstype['logicalinterconnectgroup'])                                )
    if 'nan' != description:
        scriptCode.append("             description:                 {} ".format(description)                                                       )
    scriptCode.append("             enclosureType:               {}     ".format(enclosureType)                                                     )
    scriptCode.append("             interconnectBaySet:          {}     ".format(interconnectBaySet)                                                )
    scriptCode.append("             redundancyType:              {}     ".format(redundancyType)                                                    )


    if not isFC:
        scriptCode.append(CR)
        scriptCode.append("             ethernetSettings:                           "                                                               )
        scriptCode.append("                 type:                           {}      ".format(rstype['ethernetsettings'])                                   )
        scriptCode.append("                 enableIgmpSnooping:             {}      ".format(igmpSnooping)                                          )
        scriptCode.append("                 igmpIdleTimeoutInterval:        {}      ".format(igmpIdleTimeout)                                       )
        scriptCode.append("                 enableNetworkLoopProtection:    {}      ".format(networkLoopProtection)                                 )
        scriptCode.append("                 enablePauseFloodProtection:     {}      ".format(pauseFloodProtection)                                  )
        scriptCode.append("                 enableRichTLV:                  {}      ".format(enableRichTLV)                                         )
        scriptCode.append("                 enableTaggedLldp:               {}      ".format(taggedLldp)                                            )
        scriptCode.append("                 enableFastMacCacheFailover:     {}      ".format(fastMacCacheFailover)                                  )
        scriptCode.append("                 macRefreshInterval:             {}      ".format(macRefreshInterval)                                    )
        if 'nan' != lldpIpv4Address:    
            scriptCode.append("                 lldpIpv4Address:                {}      ".format(lldpIpv4Address)                                   )
        if 'nan' != lldpIpv6Address:
            scriptCode.append("                 lldpIpv6Address:                {}      ".format(lldpIpv6Address)                                   )



    

    return scriptCode



# ================================================================================================
#
#   HELPER: generate_sas_logical_interconnect_groups
#
# ================================================================================================

def generate_sas_logical_interconnect_groups(values_in_dict,scriptCode):  
    # ---- Note:
    #       This is to define only common code

    lig                         = values_in_dict
    name                        = lig["name"]
    description                 = lig["description"] 
    enclosureType               = lig["enclosureType"]

    interconnectBaySet          = lig["interconnectBaySet"]
    redundancyType              = lig["redundancyType"]

    scriptCode.append("                                                     "                                                                       )
    scriptCode.append("     - name: Create SAS logical interconnect group {}".format(name)                                                          )
    scriptCode.append("       oneview_sas_logical_interconnect_group:       "                                                                       )
    scriptCode.append("         config: \'{{config}}\'                  "                                                                           )
    scriptCode.append("         state: present                          "                                                                           )
    scriptCode.append("         data:                                   "                                                                           )
    scriptCode.append("             name:                        \'{}\' ".format(name)                                                              )

    if 'nan' != description:
        scriptCode.append("             description:                 {} ".format(description)                                                       )
    scriptCode.append("             enclosureType:               {}     ".format(enclosureType)                                                     )
    scriptCode.append("             interconnectBaySet:          {}     ".format(interconnectBaySet)                                                )



    

    return scriptCode


# ================================================================================================
#
#   generate_logical_interconnect_groups_ansible_script_from_csv
#
# ================================================================================================
def generate_logical_interconnect_groups_ansible_script_from_csv(sheet, uplsheet, to_file):
    
    print('Creating ansible playbook   =====>           {}'.format(to_file))
    scriptCode = []
    scriptCode.append("---"                                                                                                                         )
    scriptCode.append("- name:  Configure logical_interconnect_groups from csv"                                                                     )    
    build_header(scriptCode)

    # Dictionary to find the Interconnect name per bay
    # This will be used by uplinkset to find port number for each uplink from VC to production network and FC ( Q1, Q2....)
    ICTYPE_BAY                  = dict()


    ## Configure logical interconnect group
    scriptCode.append("  tasks:"                                                                                                                    )
    sheet.dropna(how='all', inplace=True)
    sheet       = sheet.applymap(str)                       # Convert data frame into string

    for i in sheet.index:
        row                         = sheet.loc[i]
        name                        = row["name"]
        bayConfig                   = row["bayConfig"]                       # []
        if 'FC' in bayConfig:
            isFC                    = True
        else:
            isFC                    = False
        if 'SAS' in bayConfig:
            isSAS                   = True
        else:
            isSAS                   = False

        if isSAS:
            generate_sas_logical_interconnect_groups(row,scriptCode) # get common code first
        else:
            generate_logical_interconnect_groups(row,scriptCode, isFC,comeFromOV=False) # get common code first

        # Scope
        scope                       = row['scope']


        # Number of frames
        frameCount                  = row["frameCount"]                     
        frameCount                  = int(frameCount)
        scriptCode.append(CR)
        scriptCode.append("             enclosureIndexes:                   "                                                                   )
    
        for index in range(1, frameCount+1):
            if 'FC' in bayConfig:
                index               = -1
            scriptCode.append("                 - {}                        ".format(index)                                                     )

        # Map entry template
        if bayConfig:
                scriptCode.append("             interconnectMapTemplate:            "                                                               )
                scriptCode.append("                 interconnectMapEntryTemplates:  "                                                               )
                
                bay_config_list         = bayConfig.split(CR)
                #bay_config_list         = bayConfig.split(CRLF)
                for bay_config in bay_config_list:
                    if '{' in bay_config:

                        frame, config       = bay_config.split('{')
                        config              = config.rstrip(' ').rstrip('}')        # remove ' }'
                        frame               = frame.replace(' ' , '').replace('=' , '')         # remove ' = '
                        frame_name          = frame.lower().strip() + '_'                                   # Format  is enclosure1_ --> will be used in ICTYPE_BAY dict

                        enclosureIndex      = frame.lower().replace('enclosure','')
                        bay_lists           = config.split('|')
                        for el in bay_lists:
                            bay, icType     = el.split('=')

                            # Used for ICTYPE_BAY
                            bay_number      = bay.lower().strip()                                           # format is bay3 --> will be used in ICTYPE_BAY dict
                            key_enc_bay     = frame_name + bay_number                                       # format is enclosure1_bay3 --> will be used in ICTYPE_BAY dict
                            ic_name         = icType                                                        # get Interconnect Type name
                            ICTYPE_BAY[key_enc_bay] = ic_name
                            # end of ICTYPE_BAY
                            
                            
                            
                            
                            if 'FC' in icType:
                                enclosureIndex  = -1
                            bay             = bay.lower().replace('bay','')
                            if isSAS:
                                icType      = icType.replace(' ' , '')
                                icType      = '/rest/sas-interconnect-types/' + icType
                                scriptCode.append("                     - permittedInterconnectTypeUri:     \'{}\'     ".format(icType)                 )
                            else:
                                scriptCode.append("                     - permittedInterconnectTypeName:    \'{}\'     ".format(icType)                 )
                            scriptCode.append("                       enclosureIndex:                   {}         ".format(enclosureIndex)             )
                            scriptCode.append("                       logicalLocation:                             "                                    )
                            scriptCode.append("                         locationEntries:                           "                                    )
                            scriptCode.append("                             - relativeValue:            {}         ".format(bay)                        )
                            scriptCode.append("                               type:                     \'{}\'     ".format('Bay')                      )
                            scriptCode.append("                             - relativeValue:            {}         ".format(enclosureIndex)             )
                            scriptCode.append("                               type:                     \'{}\'     ".format('Enclosure')                )

                

        # Add scope here
        if not isSAS:        
            if 'nan' != scope:
                netVar              = 'var_' + name.lower().strip().replace(',', '_').replace('-', '_').replace(' ', '')
                scriptCode.append("                                                             "                                                   )
                scriptCode.append("     - name: get LIG                {}                       ".format(name)                                      )
                scriptCode.append("       oneview_logical_interconnect_group_facts:             "                                                   )
                scriptCode.append("         config:     \'{{config}}\'                          "                                                   )
                scriptCode.append("         name:       \'{}\'                                  ".format(name)                                      )
                scriptCode.append("     - set_fact:                                             "                                                   )
                scriptCode.append("          {}: ".format(netVar)  + "\'{{logical_interconnect_groups[0].uri}}\' "                                  )

                netUri              = "\'{{" + '{}'.format(netVar) + "}}\'"  
                generate_scope_for_resource(name, netUri, scope, scriptCode)


            
    ## -------------------------------------------------
    ## Configure uplink set per LIG

    
    # Sort based on LIG name 
    
    columns_names               = uplsheet.columns.tolist()
    uplsheet                    = uplsheet.sort_values(columns_names[0])
    uplsheet.dropna(how='all', inplace=True)
    uplsheet                    = uplsheet.applymap(str)                       # Convert data frame into string

    # 1 - open uplinkset sheet to set vars 

    for i in uplsheet.index:
        row                     = uplsheet.loc[i]
        nativeNetworkUri        = row["nativeNetworkUri"]
        networkUris             = row["networkUris"]                # Ex: rhosp3_storage|Vlan3100-1|Vlan_501|rhosp3_storage_mgmt
        upl_logicalPortConfigs  = row["logicalPortConfigInfos"]     #[]
        networkType             = row['networkType']

        networkType             = uplinkSetNetworkType[networkType]

        ## Define var for networks first 
        if 'nan' != networkUris:
            networkUris             = networkUris.split('|')            
            for net in networkUris:
                netName             = net.strip(' ')
                netVar              = 'var_' + netName.lower().strip().replace('-', '_') 
                if 'Ethernet' == networkType:
                    scriptCode.append("                                                     "                                                                   )
                    scriptCode.append("     - name: Get uri for network {0}             ".format(netName)                                                       )
                    scriptCode.append("       oneview_ethernet_network_facts:           "                                                                       )
                    scriptCode.append("         config: \'{{config}}\'                  "                                                                       )
                    scriptCode.append("         name:   \'{}\'                          ".format(netName)                                                       )    
                    scriptCode.append("     - set_fact:                                 "                                                                       )
                    scriptCode.append("          {}: ".format(netVar)  + "\'{{ethernet_networks.uri}}\' "                                                       )  
                    scriptCode.append(CR)

                if 'FibreChannel' == networkType:
                    scriptCode.append("                                                     "                                                                   )
                    scriptCode.append("     - name: Get uri for network {0}             ".format(netName)                                                       )
                    scriptCode.append("       oneview_fc_network_facts:                 "                                                                       )
                    scriptCode.append("         config: \'{{config}}\'                  "                                                                       )
                    scriptCode.append("         name:   \'{}\'                          ".format(netName)                                                       )    
                    scriptCode.append("     - set_fact:                                 "                                                                       )
                    scriptCode.append("          {}: ".format(netVar)  + "\'{{fc_networks[0].uri}}\' "                                                          )  
                    scriptCode.append(CR)
#HKD - to check for fcoe
                if 'fcoe' == networkType:
                    scriptCode.append("                                                     "                                                                   )
                    scriptCode.append("     - name: Get uri for network {0}             ".format(netName)                                                       )
                    scriptCode.append("       oneview_fcoe_network_facts:               "                                                                       )
                    scriptCode.append("         config: \'{{config}}\'                  "                                                                       )
                    scriptCode.append("         name:   \'{}\'                          ".format(netName)                                                       )    
                    scriptCode.append("     - set_fact:                                 "                                                                       )
                    scriptCode.append("          {}: ".format(netVar)  + "\'{{fcoe_networks.uri}}\' "                                                           )  
                    scriptCode.append(CR)

        # Define var for native network #####  Do I need this?
        # Excel panda

        if 'nan' != nativeNetworkUri:
            netName             = nativeNetworkUri.strip(' ')
            netVar              = 'var_' + netName.lower().strip().replace('-', '_')  

            scriptCode.append("                                                     "                                                                           )
            scriptCode.append("     - name: Get uri for native network {0}      ".format(netName)                                                       )
            scriptCode.append("       oneview_ethernet_network_facts:           "                                                                       )
            scriptCode.append("         config:     \'{{config}}\'              "                                                                       )
            scriptCode.append("         name:       \'{}\'                      ".format(netName)                                                       )    
            scriptCode.append("     - set_fact:                                 "                                                                       )  
            scriptCode.append("          {}= ".format(netVar)  + "\'{{ethernet_networks.uri}}\' "                                                       )


        # Define var for port number
        if 'nan' != upl_logicalPortConfigs:
            upl_logicalPortConfigs      = upl_logicalPortConfigs.replace(CRLF, '|').replace(CR,'|')
            arr_logicalPortConfigs      = upl_logicalPortConfigs.split('|')
            for config in arr_logicalPortConfigs:                               # format expected is--->  Enclosure1:Bay3:Q1|Enclosure1:Bay6:Q1
                if ':' in config:
                    enclosure,bay,portName  = config.split(':')
                    ic_location             = enclosure + ':' + bay                 # re-build enclosure1:Bay3
                    
                    ic_location             = ic_location.lower().replace(':', '_') # recycle ic_location to be used in var

                    # Use ICTYPE_BAY to find the IC name
                    key_enc_bay             = ic_location
                    interconnectName        = ICTYPE_BAY[key_enc_bay]

                    # We have interconnect name and portName. Now go find port Number from interconnect type
                    # Note: interconnect_types come from the query to OV

                    if 'FibreChannel' == networkType:
                        portName            = portName.replace('Q', '')                                         # for FC, remove Q in front

                    portNumber              = find_port_number_in_interconnect_type(interconnect_types, interconnectName, portName )
                    scriptCode.append("     - set_fact:        "                                                                                             )
                    scriptCode.append("          var_{0}_{1}:   {2}".format(ic_location, portName, portNumber)                                               )


                    #scriptCode.append("     - set_fact:                                 "                                                                    )
                    #scriptCode.append("          port_name: \'{}\'                      ".format(portName)                                                   )
                    #scriptCode.append("                                                 "                                                                    )
                    #scriptCode.append("     - name: Get port number from  {0}           ".format(interconnectName)                                           )
                    #scriptCode.append("       oneview_interconnect_type_facts:          "                                                                    )
                    #scriptCode.append("         config:    \'{{config}}\'               "                                                                    )
                    #scriptCode.append("         name:      \'{}\'                       ".format(interconnectName)                                           )
                    #scriptCode.append("     - set_fact:         "                                                                                            )
                    #scriptCode.append("          var_portInfos: \'{{interconnect_types[0].portInfos}}\'"                                                     )
                    #scriptCode.append("     - set_fact:        "                                                                                             )
                    #scriptCode.append("          var_{0}_{1}: ".format(ic_location,portName)   + " \'{{item.portNumber}}\' "                                 )
                    #scriptCode.append("       loop: \'{{var_portInfos}}\'               "                                                                    )
                    #scriptCode.append("       when: item.portName==\'{{port_name}}\'    "                                                                    )
                    #scriptCode.append(CR)


    # 2 -  open uplink set sheet to create uplink set

    currentLig                  = ""
    uplsheet.dropna(how='all', inplace=True)
    uplsheet                       = uplsheet.applymap(str)                       # Convert data frame into string

    for i in uplsheet.index:
        row                     = uplsheet.loc[i]
        name                    = row["name"]
        ligName                 = row['ligName']
        upl_logicalPortConfigs  = row["logicalPortConfigInfos"]     #[]
        desiredSpeed            = row['desiredSpeed']
        nativeNetworkUri        = row["nativeNetworkUri"]
        networkUris             = row["networkUris"]                # Ex: rhosp3_storage|Vlan3100-1|Vlan_501|rhosp3_storage_mgmt
        networkType             = row["networkType"]
        lacpTimer               = row["lacpTimer"]
        mode                    = row["mode"]
        trunking                = row['trunking']


        networkUris             = networkUris.split('|')               


        if 'nan' == desiredSpeed:
            desiredSpeed        = 'Auto'

        if 'nan' ==  mode:
            mode                = 'Auto'
        

        
        # Configure Tagged/unTagged.Tunnel for Ethernet
        if 'FibreChannel' != networkType:
            networkType         = uplinkSetNetworkType[networkType]
            ethernetNetworkType = uplinkSetEthNetworkType[networkType]

        
        if currentLig != ligName:
            scriptCode.append("                                                     "                                                                       )    
            scriptCode.append("     - name: Create uplink set {0} for LIG --> {1}".format(name, ligName)                                                    )
            scriptCode.append("       oneview_logical_interconnect_group:       "                                                                           )
            scriptCode.append("         config:     \'{{config}}\'              "                                                                           )
            scriptCode.append("         state:      present                     "                                                                           )
            scriptCode.append("         data:                                   "                                                                           )
            scriptCode.append("             name:                            {} ".format(ligName)                                                           )
            scriptCode.append("             uplinkSets:                         "                                                                           )
            # set new lig Name to be current
            currentLig          = ligName

        scriptCode.append("               - name:                        {} ".format(name)                                                                  )
        #scriptCode.append("                 type:                        \'uplink-setV300\' "                            )
        scriptCode.append("                 mode:                        {} ".format(mode)                                                                  )
        scriptCode.append("                 networkType:                 {} ".format(networkType)                                                           )

        # LACP for Ethernet networks
        if 'FibreChannel' != networkType:
            scriptCode.append("                 ethernetNetworkType:         {} ".format(ethernetNetworkType)                                               )

            if 'Auto' in mode:
                if 'nan' != lacpTimer:
                    scriptCode.append("                 lacpTimer:                   {} ".format(lacpTimer)                                                 )
        else:
            if 'Enabled' == trunking:
                scriptCode.append("                 fcMode:                        TRUNK "                                                                  )                

        # List of networks
        if 'nan' != networkUris:
            scriptCode.append("                 networkUris:                "                                                                               )
            for net in networkUris:
                netName         = net.strip(' ')
                netVar          = 'var_' + netName.lower().strip().replace('-', '_') 
                netUri          = "\'{{" + '{}'.format(netVar) + "}}\'"
                scriptCode.append("                     - {}                        ".format(netUri)                                                        )

        # nativeNetworkUri

        if 'nan' != nativeNetworkUri:

            netName             = nativeNetworkUri.strip(' ')
            netVar              = 'var_' + netName.lower().strip().replace('-', '_') 
            netUri          = "\'{{" + '{}'.format(netVar) + "}}\'"
            scriptCode.append("                 nativeNetworkUri:               {} ".format(netUri)                                                     )


        # logicalPortInfos
        if 'nan' != upl_logicalPortConfigs:
            scriptCode.append(CR)
            scriptCode.append("                 logicalPortConfigInfos:                "                                                                 )
            upl_logicalPortConfigs      = upl_logicalPortConfigs.replace(CRLF, '|').replace(CR,'|')
            arr_logicalPortConfigs      = upl_logicalPortConfigs.split('|')
            for config in arr_logicalPortConfigs:
                if ':' in config:
                    enclosure,bay,port      = config.split(':')
                    ic_location             = enclosure + ':' + bay                                         # re-build enclosure1:Bay3
                    ic_location             = ic_location.lower().replace(':', '_')                         # recycle ic_location to be used in var                    
                    enclosure               = enclosure.lower().strip('enclosure')
                    bay                     = bay.lower().strip('bay')
                    port                    = port.strip(' ')

                    if 'FibreChannel' == networkType:
                        port                = port.replace('Q', '')                                         # for FC, remove Q in front
                        enclosure           = -1

                    port_number             = "\'{{var_" + '{0}_{1}'.format(ic_location,port) + "}}\'"     # port = '{{var_enclosure1_bay_3_portNumber}}' 

                    scriptCode.append("                     - desiredSpeed:             {}     ".format(desiredSpeed)                                        )
                    scriptCode.append("                       logicalLocation:                 "                                                             )
                    scriptCode.append("                         locationEntries:               "                                                             )    
                    scriptCode.append("                             - relativeValue:    {}              ".format(bay)                                        )
                    scriptCode.append("                               type:             \'Bay\'         "                                                    )
                    scriptCode.append("                             - relativeValue:    {}              " .format(port_number)                               )
                    scriptCode.append("                               type:             \'Port\'        "                                                    )
                    scriptCode.append("                             - relativeValue:    {}              " .format(enclosure)                                 )
                    scriptCode.append("                               type:             \'Enclosure\'   "                                                    )




    # end of LIG / uplinkset
    scriptCode.append("       delegate_to: localhost                    "                                                                               )
    scriptCode.append(CR)
    #print(CR.join(scriptCode))




    # ============= Write scriptCode ====================
    write_to_file(scriptCode, to_file)





# ================================================================================================
#
#   HELPER: generate_enclosure_groups
#
# ================================================================================================

def generate_enclosure_groups(values_in_dict,scriptCode):  
    # ---- Note:
    #       This is to define only common code
    # 
    
    eg                      = values_in_dict
    name                    = eg["name"]
    description             = eg['description']
    enclosureCount          = eg["enclosureCount"]
    powerRedundantMode      = eg["powerRedundantMode"]


    scriptCode.append("                                                     "                                                                                )
    scriptCode.append("     - name: Create enclosure group {}           ".format(name)                                                                       )
    scriptCode.append("       oneview_enclosure_group:                  "                                                                                    )
    scriptCode.append("         config: \'{{config}}\'                  "                                                                                    )
    scriptCode.append("         state: present                          "                                                                                    )
    scriptCode.append("         data:                                   "                                                                                    )
    scriptCode.append("             name:                        \'{}\' ".format(name)                                                                       )
              



    return scriptCode

# ================================================================================================
#
#   generate_enclosure_groups_ansible_script_from_csv
#
# ================================================================================================
def generate_enclosure_groups_ansible_script_from_csv(sheet, to_file):
    
    print('Creating ansible playbook   =====>           {}'.format(to_file))    
    scriptCode = []
    scriptCode.append("---"                                                                                                                                  )
    scriptCode.append("- name:  Configure enclosure groups from csv"                                                                                         )    
    build_header(scriptCode)

    ## Configure logical interconnect group
    scriptCode.append("  tasks:"                                                                                                                    )
    sheet.dropna(how='all', inplace=True)
    sheet       = sheet.applymap(str)                       # Convert data frame into string

    for i in sheet.index:
        row                         = sheet.loc[i]
        name                        = row["name"]

        enclosureCount              = row["enclosureCount"]
        powerRedundantMode          = row["powerRedundantMode"]
        ligMappings                 = row['logicalInterConnectGroupMapping']    # []

        ipAddressingMode            = row['ipAddressingMode']
        ipRangeUris                 = row['ipRangeUris']                        # []
        scope                       = row['scope']

        if ligMappings:
            ligMappings         = ligMappings.split('|')
            for element in ligMappings:                                     # Frame1=LIG-ETH,LIG-SAS,LIG-FC
                frame,lig_list  = element.split('=')
                frame           = frame.strip(' ').lower()
                lig_list        = lig_list.split(',')
                for ligName in lig_list:
                    ligName     = ligName.strip(' ') 
                    lig_name    = ligName.replace('-', '_')                 # to be used in vars
                    scriptCode.append("     - set_fact:                                 "                                                                )
                    scriptCode.append("          IC_OFFSET:     3                       "                                                                )

                    scriptCode.append("                                                 "                                                                )
                    scriptCode.append("     - name: Get lig uri from  {0}               ".format(ligName)                                                )
                    scriptCode.append("       oneview_logical_interconnect_group_facts: "                                                                )
                    scriptCode.append("         config:    \'{{config}}\'               "                                                                )
                    scriptCode.append("         name:      \'{}\'                       ".format(ligName)                                                )
                    scriptCode.append("     - set_fact:         "                                                                                        )
                    scriptCode.append("         lig:       \'{{logical_interconnect_groups}}\' "                                                         )       
                    # if it's not lig then try with sas_lig
                    scriptCode.append("     - name: Get sas lig uri from  {0}               ".format(ligName)                                            )
                    scriptCode.append("       oneview_sas_logical_interconnect_group_facts: "                                                            )
                    scriptCode.append("         config:    \'{{config}}\'               "                                                                )
                    scriptCode.append("         name:      \'{}\'                       ".format(ligName)                                                )
                    scriptCode.append("     - set_fact:         "                                                                                        )
                    scriptCode.append("         lig:       \'{{sas_logical_interconnect_groups}}\' "                                                     )   
                    scriptCode.append("       when: (lig|length == 0)                   "                                                                )

                    scriptCode.append("     - set_fact:         "                                                                                        )
                    scriptCode.append("          var_{0}_{1}_uri:      ".format(frame, lig_name) + "\'{{lig[0].uri}}\'"                                  )
                    scriptCode.append("          var_{0}_{1}_bay_prim: ".format(frame, lig_name) + "\'{{lig[0].interconnectBaySet}}\'"                   )
                    scriptCode.append("          var_{0}_{1}_bay_sec:  ".format(frame, lig_name) + " \'{{lig[0].interconnectBaySet + IC_OFFSET}}\'"      )


        # build ID pools uri if used
        var_range_names = []
        if ipAddressingMode == 'IpPool' and ipRangeUris:
            range_names     = ipRangeUris.split('|')

            for r_name in range_names: 
                idpool_name     = r_name.lower().strip(' ').replace('-', '_')                 # to be used in vars        
                var_name        = "var_{0}_uri   ".format(idpool_name)
                var_range_names.append(var_name)
                
                scriptCode.append("                                                 "                                                             )
                scriptCode.append("     - name: Get uri for range {}          ".format(r_name)                                                    )
                scriptCode.append("       oneview_id_pools_ipv4_range_facts: "                                                                    )
                scriptCode.append("         config:     \'{{config}}\'              "                                                             )
                scriptCode.append("         name:       \'{}\'                      ".format(r_name)                                              )
                scriptCode.append("     - set_fact:  " )
                scriptCode.append("          {}:   ".format(var_name) + "\'{{id_pools_ipv4_ranges[0].uri}}\'"                                     )


        scriptCode.append("                                                     "                                                                 )
        scriptCode.append("     - name: Create enclosure group {0}".format(name)                                                                  )
        scriptCode.append("       oneview_enclosure_group:                  "                                                                     )
        scriptCode.append("         config:    \'{{config}}\'               "                                                                     )
        scriptCode.append("         state:      present                     "                                                                     )
        scriptCode.append("         data:                                   "                                                                     )
        #scriptCode.append("             type:                            {} ".format(rstype['enclosuregroup'])                                    )
        scriptCode.append("             name:                            {} ".format(name)                                                        )
        scriptCode.append("             enclosureCount:                  {} ".format(enclosureCount)                                              )
        scriptCode.append("             powerMode:                       {} ".format(powerRedundantMode)                                          )
    
        # Build interconnectBayMappings
        if ligMappings:
            scriptCode.append("             interconnectBayMappings: "                                                                            )

            for element in ligMappings:                                     # Frame1=LIG-ETH,LIG-SAS,LIG-FC
                frame,lig_list  = element.split('=')
                frame           = frame.strip(' ').lower()
                lig_list        = lig_list.split(',')
                for ligName in lig_list:
                    ligName     = ligName.strip(' ') 
                    lig_name    = ligName.replace('-', '_')                 # to be used in vars


                    lig_uri         = "{{" + "var_{0}_{1}_uri".format(frame, lig_name) + "}}"
                    var_bay_prim    = "{{" + "var_{0}_{1}_bay_prim".format(frame, lig_name) + "}}"
                    var_bay_sec     = "{{" + "var_{0}_{1}_bay_sec".format(frame, lig_name) + "}}"

                    enclosureIndex  = frame.strip('frameFrameenclosureEnclosure')

                    scriptCode.append("                 - interconnectBay:              \'{}\'      ".format(var_bay_prim)                                    )
                    scriptCode.append("                   logicalInterconnectGroupUri:  \'{}\'      ".format(lig_uri)                                         )
                    scriptCode.append("                   enclosureIndex:               {}          ".format(enclosureIndex)                                  )
                    scriptCode.append(""                                                                                                                      )
                    scriptCode.append("                 - interconnectBay:               \'{}\'     ".format(var_bay_sec)                                     )
                    scriptCode.append("                   logicalInterconnectGroupUri:  \'{}\'      ".format(lig_uri)                                         )
                    scriptCode.append("                   enclosureIndex:               {}          ".format(enclosureIndex)                                  )
                    
        # Build ipv4 ranges

        scriptCode.append("")
        scriptCode.append("             ipAddressingMode:                {} ".format(ipAddressingMode)                                            )
        if var_range_names:
            scriptCode.append("             ipRangeUris: " )
            for idpool_name in var_range_names:
                scriptCode.append("                 - \'{{" + "{0}".format(idpool_name.strip(' ')) +  "}}\' "                                     )

        # Add scope here
        
        if 'nan' != scope:
            netVar              = 'var_' + name.lower().strip().replace(',', '_').replace('-', '_').replace(' ', '')
            scriptCode.append("                                                             "                                                   )
            scriptCode.append("     - name: get logical enclosure                {}         ".format(name)                                      )
            scriptCode.append("       oneview_enclosure_group_facts:                        "                                                   )
            scriptCode.append("         config:     \'{{config}}\'                          "                                                   )
            scriptCode.append("         name:       \'{}\'                                  ".format(name)                                      )
            scriptCode.append("     - set_fact:                                             "                                                   )
            scriptCode.append("          {}: ".format(netVar)  + "\'{{enclosure_groups.uri}}\' "                                                )

            netUri              = "\'{{" + '{}'.format(netVar) + "}}\'"  

            generate_scope_for_resource(name, netUri, scope, scriptCode)




        # end of enclosure group
        scriptCode.append("       delegate_to: localhost                    "                                                                                 )
        scriptCode.append(CR)

    #print(CR.join(scriptCode))

    # ============= Write scriptCode ====================
    write_to_file(scriptCode, to_file)
                    


# ================================================================================================
#
#   generate_logical_enclosures_ansible_script_from_csv
#
# ================================================================================================
def generate_logical_enclosures_ansible_script_from_csv(sheet, to_file):
    
    print('Creating ansible playbook   =====>           {}'.format(to_file))    
    scriptCode = []
    scriptCode.append("---"                                                                                                                                 )
    scriptCode.append("- name:  Configure logical enclosures from csv"                                                                                      )    
    build_header(scriptCode)

    scriptCode.append("  tasks:"                                                                                                                    )
    sheet.dropna(how='all', inplace=True)
    sheet       = sheet.applymap(str)                       # Convert data frame into string

    for i in sheet.index:
        row                         = sheet.loc[i]
        name                        = row['logicalEnclosureName']
        enclosureName               = row['enclosureName']          # []
        enclosureNewname            = row['enclosureNewname']       # []
        enclosureGroup              = row['enclosureGroup']
        fwBaseline                  = row['fwBaseline']     
        fwInstall                   = row['fwInstall']
        scope                       = row['scope']

        # Get enclosure Uris
        if 'nan' != enclosureName:
            list_enclosure          = enclosureName.split('|')
            for enclosure in list_enclosure:
                var_encl_name       = enclosure.replace('-','_')
                scriptCode.append("                                                     "                                                                       )
                scriptCode.append("     - name: Get uri of enclosure {}                 ".format(enclosure)                                                     )
                scriptCode.append("       oneview_enclosure_facts:                      "                                                                       )
                scriptCode.append("         config:            \'{{ config }}\'         "                                                                       )
                scriptCode.append("     - set_fact:                                     "                                                                       )          
                scriptCode.append("         list_enclosures : \'{{enclosures}}\'        "                                                                       )                               
                scriptCode.append("     - set_fact:                                     "                                                                       )                                    
                scriptCode.append("         var_{}_uri:       ".format(var_encl_name)   + "\'{{item.uri}}\'"                                                    )       
                scriptCode.append("       loop: '{{list_enclosures}}'                   "                                                                       )
                scriptCode.append("       when: item.name== \'{}\'                      ".format(enclosure)                                                     ) 


            # Get enclosure group Uris
            scriptCode.append("                                                     "                                                                       )
            scriptCode.append("     - name: Get uri of enclosure group {}           ".format(enclosureGroup)                                                )
            scriptCode.append("       oneview_enclosure_group_facts:              "                                                                         )
            scriptCode.append("         config:   \'{{ config }}\'                  "                                                                       )
            scriptCode.append("         name:     {}                                ".format(enclosureGroup)                                                )
            scriptCode.append("     - set_fact:                                     "                                                                       )
            scriptCode.append("         var_{}_uri:   ".format(enclosureGroup) + "\'{{enclosure_groups.uri}}\' "                                            )
            scriptCode.append("")



            # Create logical enclosure
            scriptCode.append("                                                     "                                                                       )
            scriptCode.append("     - name: Create logical enclosure {}             ".format(name)                                                          )
            scriptCode.append("       oneview_logical_enclosure:                    "                                                                       )
            scriptCode.append("         config:   \'{{ config }}\'                  "                                                                       )
            scriptCode.append("         state:     present                          "                                                                       )
            scriptCode.append("         data:                                       "                                                                       )
            scriptCode.append("             name:   \'{}\'                          ".format(name)                                                          )
            scriptCode.append("             enclosureGroupUri:  \'{{" + "var_{}_uri".format(enclosureGroup) + "}}\'"                                        )                                               
            scriptCode.append("             enclosureUris:                          "                                                                       )
            if 'nan' != enclosureName:
                list_enclosure          = enclosureName.split('|')
                for enclosure in list_enclosure:
                    var_encl_uri        = 'var_{}_uri'.format(enclosure.replace('-','_'))                                                                                      
                    scriptCode.append("                 -  \'{{" + var_encl_uri + "}}\'"                                                                    )                                        
            scriptCode.append("") 

            ## Add firmware 
            # Get fw Baseline Uris -
            if 'nan' != fwBaseline and 'true' in fwInstall.lower() :
                fw                  = fwBaseline.strip().replace(' ','_').replace('-','_').lower()
                fwInstall           = 'true' in fwInstall.lower()
                scriptCode.append("                                                     "                                                                      )
                scriptCode.append("     - name: Get uri of fw baseline {}               ".format(fwBaseline)                                                   )
                scriptCode.append("       oneview_firmware_driver_facts:                "                                                                      )
                scriptCode.append("         config:       \'{{ config }}\'              "                                                                      )
                scriptCode.append("         name:         \'{}\'                        ".format(fwBaseline)                                                   )
                scriptCode.append("     - set_fact:                                     "                                                                      )
                scriptCode.append("         var_fw_{}_uri:   ".format(fw) + "\'{{firmware_drivers[0].uri}}\' "                                                 )
                scriptCode.append("")


                scriptCode.append("                                                     "                                                                      )
                scriptCode.append("     - name: Update firmware on logical enclosure {} ".format(name)                                                         )
                scriptCode.append("       oneview_logical_enclosure:                    "                                                                      )
                scriptCode.append("         config:   \'{{ config }}\'                  "                                                                      )
                scriptCode.append("         state:     firmware_updated                 "                                                                      )
                scriptCode.append("         data:                                       "                                                                      )
                scriptCode.append("             name:   \'{}\'                          ".format(name)                                                         )
                scriptCode.append("             firmware:                               "                                                                      )
                scriptCode.append("                 firmwareBaselineUri:      \'{{" + "var_fw_{}_uri".format(fw) +  "}}\'"                                     )
                scriptCode.append("                 forceInstallFirmware:      {}".format(fwInstall)                                                           )
                scriptCode.append("                 firmwareUpdateOn:          \'EnclosureOnly\' "                                                             )
                scriptCode.append("             custom_headers:                         "                                                                      )                           
                scriptCode.append("                 if-Match: \'*\'                     "                                                                      )

    
        # Add scope here
        
        if 'nan' != scope:
            netVar              = 'var_' + name.lower().strip().replace(',', '_').replace('-', '_').replace(' ', '')
            scriptCode.append("                                                             "                                                   )
            scriptCode.append("     - name: get logical enclosure                {}         ".format(name)                                      )
            scriptCode.append("       oneview_logical_enclosure_facts:                      "                                                   )
            scriptCode.append("         config:     \'{{config}}\'                          "                                                   )
            scriptCode.append("         name:       \'{}\'                                  ".format(name)                                      )
            scriptCode.append("     - set_fact:                                             "                                                   )
            scriptCode.append("          {}: ".format(netVar)  + "\'{{logical_enclosures.uri}}\' "                                              )

            netUri              = "\'{{" + '{}'.format(netVar) + "}}\'"  

            generate_scope_for_resource(name, netUri, scope, scriptCode)

    

            # end of logical enclosure 
            scriptCode.append("       delegate_to: localhost                        "                                                                           )
            scriptCode.append(CR)
        else:
            print(' No enclosure specified--> cennot create logical enclosure. Skip creating playbook for logical enclosure {}'.format(name))
    #print(CR.join(scriptCode))

    # ============= Write scriptCode ====================
    write_to_file(scriptCode, to_file)



# ================================================================================================
#
#   generate_profile_or_templates
#
# ================================================================================================
def generate_profile_or_template(values_in_dict,scriptCode):

    # Coomon code for server profiles and templates

    prof                                = values_in_dict
    serverHardwareTypeName              = prof['serverHardwareTypeName']
    enclosureGroupName                  = prof['enclosureGroupName']
    affinity                            = prof["affinity"]
    wwnType                             = prof["wwnType"]
    macType                             = prof["macType"]
    serialNumberType                    = prof["serialNumberType"]
    iscsiInitiatorNameType              = prof["iscsiInitiatorNameType"]
    hideUnusedFlexNics                  = prof["hideUnusedFlexNics"]
    manageMode                          = prof["manageMode"]    
    mode                                = prof['mode']
    pxeBootPolicy                       = prof['pxeBootPolicy']
    secureBoot                          = prof['secureBoot']
    manageBoot                          = prof["manageBoot"]    
    order                               = prof['order']                           
    manageBios                          = prof['manageBios']
    overriddenSettings                  = prof['overriddenSettings']
    manageFirmware                      = prof['manageFirmware']
    firmwareBaselineName                = prof['firmwareBaselineName']
    firmwareInstallType                 = prof['firmwareInstallType']
    forceInstallFirmware                = prof['forceInstallFirmware']
    firmwareActivationType              = prof['firmwareActivationType']
    hideUnusedFlexNics                  = prof['hideUnusedFlexNics']
    scope                               = prof['scope']



    # Generate code
    scriptCode.append("             serverHardwareTypeName:       {}            ".format(serverHardwareTypeName)                                            )
    scriptCode.append("             enclosureGroupName:           {}            ".format(enclosureGroupName)                                                )

    scriptCode.append("             affinity:                    {}             ".format(affinity)                                                           )

    if 'nan' != hideUnusedFlexNics:
        scriptCode.append("             hideUnusedFlexNics:          {}         ".format(hideUnusedFlexNics)                                                 )
    if 'nan' != iscsiInitiatorNameType:
        scriptCode.append("             iscsiInitiatorNameType:      {}         ".format(iscsiInitiatorNameType)                                             )

    # types region
    if 'nan' != wwnType:
        scriptCode.append("             wwnType:                     {}         ".format(wwnType)                                                            )
    if 'nan' != macType:
        scriptCode.append("             macType:                     {}         ".format(macType)                                                            )
    if 'nan' != serialNumberType:
        scriptCode.append("             serialNumberType:            {}         ".format(serialNumberType)                                                   )
    if 'nan' != iscsiInitiatorNameType:
        scriptCode.append("             iscsiInitiatorNameType:      {}         ".format(iscsiInitiatorNameType)                                             )

    # bootMode region

    if 'nan' == pxeBootPolicy:
        pxeBootPolicy                       = 'Auto'
    if 'nan' == secureBoot:
        secureBoot                          = 'Disabled'
    if 'nan' != manageMode:
        manageMode                          = manageMode.lower().capitalize() 
    if 'True' in manageMode:
        scriptCode.append("             bootMode:                               "                                                                            )
        scriptCode.append("                 manageMode:              {}         ".format(manageMode)                                                         )
        scriptCode.append("                 mode:                    {}         ".format(mode)                                                               )
        if 'BIOS' in mode.upper():
            scriptCode.append("                 secureBoot:              Disabled  "                                                                         )            
        if 'UEFI' == mode.upper():
            scriptCode.append("                 secureBoot:              Disabled   "                                                                        )    
            scriptCode.append("                 pxeBootPolicy:                   ".format(pxeBootPolicy)                                                     )   
        if 'optimized' in mode.lower():
            scriptCode.append("                 secureBoot:                      ".format(secureBoot)                                                        )    
            scriptCode.append("                 pxeBootPolicy:                   ".format(pxeBootPolicy)                                                     )                

        # boot order is allowed ONLY if managedMode is True
        if 'nan' != manageBoot:
            manageBoot              = manageBoot.lower().capitalize()
        if 'True' in manageBoot and 'nan' != order:
            scriptCode.append("             boot:                                    "                                                                          )
            scriptCode.append("                 manageBoot:              {}          ".format(manageBoot)                                                       )
            scriptCode.append("                 order:                               "                                                                          )

            for boot_order in order.split('|'):
                boot_order      = boot_order.strip(' ')
                scriptCode.append("                     - {}                         ".format(boot_order)                                                       )

    # BIOS settings
    if 'nan' != manageBios:
        manageBios              = manageBios.lower().capitalize()
    if 'True' in manageBios and 'nan' != overriddenSettings:
        scriptCode.append("             bios:                                    "                                                                          )
        scriptCode.append("                 manageBios:               {}         ".format(manageBios)                                                       )
        scriptCode.append("                 overriddenSettings:                                    "                                                        )
        for setting in overriddenSettings.split('|'):                               # format is id=EnergyEfficientTurbo;value=Disabled|id=PowerRegulator;value=StaticHighPerf
            if setting.startswith('id'):
                _id,value                       = setting.split(';')
                id_name,id_attribute            = _id.split('=')
                id_name                         = id_name.strip(' ')
                id_attribute                    = id_attribute.strip(' ')
                value_name, value_attribute     = value.split('=')
                value_name                      = value_name.strip(' ')
                value_attribute                 = value_attribute.strip(' ')

                scriptCode.append("                 - {0}:         {1}              ".format(id_name,id_attribute)                                          )
                scriptCode.append("                   {0}:      {1}                 ".format(value_name,value_attribute)                                                       )

    # Firmware
    if 'nan' != manageFirmware:
        manageFirmware                  = manageFirmware.lower().capitalize()
        if 'nan' == firmwareActivationType:
            firmwareActivationType  = 'NotScheduled'      # Values are: Immediate, Scheduled, NotScheduled

        if 'True' in manageFirmware:
            forceInstallFirmware        = forceInstallFirmware.lower().capitalize()
            scriptCode.append("             firmware:                               "                                                                       )
            scriptCode.append("                 manageFirmware:          {}         ".format(manageFirmware)                                                )
            scriptCode.append("                 firmwareInstallType:     {}         ".format(firmwareInstallType)                                                                )
            scriptCode.append("                 forceInstallFirmware:    {}         ".format(forceInstallFirmware)                                                                )
            scriptCode.append("                 firmwareBaselineName:    {}         ".format(firmwareBaselineName)                                                                )
            scriptCode.append("                 firmwareActivationType:  {}         ".format(firmwareActivationType)                                                                 )


# ================================================================================================
#
#   HELPER  for profile and templates -Add LOCAL Storage - Add Connections
#
# ================================================================================================

def generate_connection_storage_for_profile(connectionsheet,localstoragesheet,thisProfilename,isProfile,scriptCode):



    # -------------------------------- Local storage CSV
    currentProfileName              = ""
    columns_names                   = localstoragesheet.columns.tolist()
    localstoragesheet               = localstoragesheet.sort_values(columns_names[0])
    localstoragesheet.dropna(how='all', inplace=True)
    localstoragesheet               = localstoragesheet.applymap(str)                       # Convert data frame into string
    subsetlocalstoragesheet         = localstoragesheet[localstoragesheet[columns_names[0]]== thisProfilename]

    for i in subsetlocalstoragesheet.index:
        row                         = subsetlocalstoragesheet.loc[i]

        profileName                 = row['profileName']
        deviceSlot                  = row['deviceSlot']
        driveWriteCache             = row['driveWriteCache']
        initialize                  = row['initialize']
        ld_name                     = row['logicalDiskName']                # []
        driveTechnology             = row['driveTechnology']                # []
        bootable                    = row['bootable']                       # []
        numPhysicalDrives           = row['numPhysicalDrives']              # []
        raidLevel                   = row['raidLevel']
        accelerator                 = row['accelerator']
        mode                        = row['mode']

        # Generate Code
        if 'nan' != deviceSlot:
            if 'nan' == driveWriteCache:
                driveWriteCache     = 'Unmanaged'
            driveWriteCache         = driveWriteCache.lower().capitalize()
            if  'nan' == initialize:
                initialize          = 'False'
            initialize              = initialize.lower().capitalize()

            if 'nan' != ld_name:
                ldname_list                 = ld_name.split('|')
                if 'nan' != driveTechnology:
                    driveTechnology_list    = driveTechnology.split('|')
                if 'nan' != bootable:
                    bootable_list           = bootable.split('|')
                if 'nan' != numPhysicalDrives:
                    numPhysicalDrives_list  = numPhysicalDrives.split('|')
                if 'nan' != raidLevel:
                    raidLevel_list          = raidLevel.split('|')
                if 'nan' != accelerator:
                    accelerator_list        = accelerator.split('|')

            if currentProfileName != profileName:
                scriptCode.append("                                                     "                                                                   )
                scriptCode.append("     - name: Add local storage \'{0}\' to server profile or template {1}    ".format(deviceSlot, profileName)            )
                if isProfile:
                    scriptCode.append("       oneview_server_profile:                   "                                                                   )
                else:
                    scriptCode.append("       oneview_server_profile_template:          "                                                                   )
                scriptCode.append("         config:     \'{{ config }}\'                "                                                                   )
                scriptCode.append("         state:      present                         "                                                                   )
                scriptCode.append("         data:                                       "                                                                   )
                scriptCode.append("             name:                          \'{}\'   ".format(profileName)                                               ) 
                if isProfile:                 
                    scriptCode.append("             type:                          \'{}\'  ".format(rstype['serverprofile'])                                )
                else:
                    scriptCode.append("             type:                          \'{}\'  ".format(rstype['serverprofiletemplate'])                        )
                scriptCode.append("             localStorage:                           "                                                                   )
                scriptCode.append("                 controllers:                        "                                                                   )
                # set new profile name to be current
                currentProfileName      = profileName
                
            scriptCode.append("                     - deviceSlot:           {}      ".format(deviceSlot)                                                ) 
            scriptCode.append("                       driveWriteCache:      {}      ".format(driveWriteCache)                                           )
            scriptCode.append("                       initialize:           {}      ".format(initialize)                                                )
            scriptCode.append("                       mode:                 {}      ".format(mode)                                                      )

            if 'nan' != ld_name:
                scriptCode.append("                       logicalDrives:                "                                                               )
                for ld in ldname_list:
                    index   = ldname_list.index(ld)
                    dt      = driveTechnology_list[index]
                    bt      = bootable_list[index].lower().capitalize()
                    ac      = accelerator_list[index].lower().capitalize()
                    ra      = raidLevel_list[index]
                    np      = numPhysicalDrives_list[index]

                    scriptCode.append("                         - name:                 {}   ".format(ld)                                               )
                    if dt:            
                        scriptCode.append("                           driveTechnology:      {}".format(dt)                                              )
                    scriptCode.append("                           bootable:             {}  ".format(bt)                                                )
                    scriptCode.append("                           numPhysicalDrives:    {}  ".format(np)                                                )
                    scriptCode.append("                           raidLevel:            {}  ".format(ra)                                                )
                    if ac:
                        scriptCode.append("                           accelerator:          {} ".format(ac)                                             )


    
    # -------------------------------- Network Connection CSV
    currentProfileName              = ""

    columns_names                   = connectionsheet.columns.tolist()
    connectionsheet                 = connectionsheet.sort_values(columns_names[0])
    connectionsheet.dropna(how='all', inplace=True)
    connectionsheet                 = connectionsheet.applymap(str)                       # Convert data frame into string

    subsetconnectionsheet           = connectionsheet[connectionsheet[columns_names[0]]== thisProfilename]

    for i in subsetconnectionsheet.index:
        row                         = subsetconnectionsheet.loc[i]

        name                        = row['name']
        profileName                 = row['profileName']
        id                          = row['id']
        functionType                = row['functionType'].capitalize()
        networkUri                  = row['networkUri']
        portId                      = row['portId']
        requestedMbps               = row['requestedMbps']
        requestedVFs                = row['requestedVFs']
        lagName                     = row['lagName']

        boot                        = row['boot']
        priority                    = row['priority']

        userDefined                 = row['userDefined']
        userDefined                 = userDefined.lower().capitalize()
        if 'True' in userDefined:
            mac                     = row['mac']
            wwwnn                   = row['wwnn']   
            wwpn                    = row['wwpn']

        if 'nan' != name:
            manageConnections       = True
        if 'nan' == boot:
            boot                    = 'False'
        boot                        = boot.lower().capitalize()
        


        if currentProfileName != profileName:
            # create new play for connections
            scriptCode.append("                                                     "                                                                       )
            scriptCode.append("     - name: Add network connection \'{0}\'  to server profile or template {1}     ".format(networkUri, profileName)         )
            if isProfile:
                scriptCode.append("       oneview_server_profile:                   "                                                                       )
            else:
                scriptCode.append("       oneview_server_profile_template:          "                                                                       )
            scriptCode.append("         config:     \'{{ config }}\'                "                                                                       )
            scriptCode.append("         state:      present                         "                                                                       )
            scriptCode.append("         data:                                       "                                                                       )
            scriptCode.append("             name:                          \'{}\'   ".format(profileName)                                                   )    
            if isProfile:              
                scriptCode.append("             type:                          \'{}\'   ".format(rstype['serverprofile'])                                   )
            else:
                scriptCode.append("             type:                          \'{}\'   ".format(rstype['serverprofiletemplate'])                           )
            scriptCode.append("             connectionSettings:                     "                                                                       )
            if not isProfile:           # Only template has it
                scriptCode.append("                 manageConnections:          {}      ".format(manageConnections)                                             )
            scriptCode.append("                 connections:                        "                                                                       )
            # set new profile name to be current
            currentProfileName      = profileName


        scriptCode.append("                     - id:                   {}      ".format(id)                                                                ) 
        scriptCode.append("                       portId:               {}      ".format(portId)                                                            ) 
        scriptCode.append("                       name:                 {}      ".format(name)                                                              ) 
        scriptCode.append("                       functionType:         {}      ".format(functionType)                                                      ) 
        scriptCode.append("                       networkName:          {}      ".format(networkUri)                                                        ) 

        if 'nan' != lagName:
            scriptCode.append("                       lagName:              {}      ".format(lagName)                                                       ) 
        if 'nan' != requestedVFs:
            scriptCode.append("                       requestedVFs:         {}      ".format(requestedVFs)                                                  ) 
        if 'nan' != requestedMbps:
            scriptCode.append("                       requestedMbps:        {}      ".format(requestedMbps)                                                 )

        # Bootable connection
        if 'True'in boot:
            scriptCode.append("                       boot:                     "                                                                           )
            scriptCode.append("                         priority:           {}      ".format(priority)                                                      )
            if 'Ethernet' in functionType:
                scriptCode.append("                         ethernetBootType:   PXE  "                                                                    )

        ### mac/wwn
        if 'True' in userDefined:
            if 'nan' != mac and 'Ethernet' in functionType:
                scriptCode.append("                       mac:                     \'{}\'      ".format(macAddress)                                        )
    
            if 'nan' != wwwnn and 'FibreChannel' in functionType:
                scriptCode.append("                       wwwn:                      \'{}\'      ".format(wwwn)                                            )
            if 'nan' != wwwpn and 'FibreChannel' in functionType:
                scriptCode.append("                       wwpn:                      \'{}\'      ".format(wwpn)                                            )






# ================================================================================================
#
#   generate_server_profile_templates_ansible_script_from_csv
#
# ================================================================================================
def generate_server_profile_templates_ansible_script_from_csv(sheet,connectionsheet,localstoragesheet, to_file):
    
    print('Creating ansible playbook   =====>           {}'.format(to_file))    
    scriptCode = []
    scriptCode.append("---"                                                                                                                                 )
    scriptCode.append("- name:  Creating server profile templates from csv"                                                                                 )    
    build_header(scriptCode)

    scriptCode.append("  tasks:"                                                                                                                    )
    sheet.dropna(how='all', inplace=True)
    sheet       = sheet.applymap(str)                       # Convert data frame into string

    for i in sheet.index:
        row                         = sheet.loc[i]

        name                        = row['name']
        description                 = row['description']
        serverProfileDescription    = row['serverProfileDescription']
        scope                       = row['scope']




        # Create server profile template
        scriptCode.append("                                                     "                                                                       )
        scriptCode.append("     - name: Create server profile template {}       ".format(name)                                                          )
        scriptCode.append("       oneview_server_profile_template:              "                                                                       )
        scriptCode.append("         config:     \'{{ config }}\'                "                                                                       )
        scriptCode.append("         state:      present                         "                                                                       )
        scriptCode.append("         data:                                       "                                                                       )

        scriptCode.append("             name:                       \'{}\'      ".format(name)                                                          )
        scriptCode.append("             type:                       \'{}\'      ".format(rstype['serverprofiletemplate'])                               )
        if 'nan' != description:
            scriptCode.append("             description:                \'{}\'  ".format(description)                                                   )
        if 'nan' != serverProfileDescription:
            scriptCode.append("             serverProfileDescription:   \'{}\'  ".format(serverProfileDescription)                                      )

        # Add profile attributes
        generate_profile_or_template(row,scriptCode)

        # Add network connections and storage
        isProfile   = False
        generate_connection_storage_for_profile(connectionsheet,localstoragesheet,name,isProfile,scriptCode) 

    # Add scope
    if 'nan' != scope:
        netVar              = 'var_' + name.lower().strip().replace(',', '_').replace('-', '_').replace(' ', '') 
        scriptCode.append("                                                             "                                                   )
        scriptCode.append("     - name: get spt                {}                       ".format(name)                                      )
        scriptCode.append("       oneview_server_profile_template_facts:                "                                                   )
        scriptCode.append("         config:     \'{{config}}\'                          "                                                   )
        scriptCode.append("         name:       \'{}\'                                  ".format(name)                                      )
        scriptCode.append("     - set_fact:                                             "                                                   )
        scriptCode.append("          {}: ".format(netVar)  + "\'{{server_profile_templates[0].uri}}\' "                                     )

        netUri              = "\'{{" + '{}'.format(netVar) + "}}\'"  
        generate_scope_for_resource(name, netUri, scope, scriptCode)
    

        



    # end of server profile template
    scriptCode.append("       delegate_to: localhost                    "                                                                                  )
    scriptCode.append(CR)
    #print(CR.join(scriptCode))

    # ============= Write scriptCode ====================
    write_to_file(scriptCode, to_file)



# ================================================================================================
#
#   generate_server_profiles_ansible_script_from_csv
#
# ================================================================================================
def generate_server_profiles_ansible_script_from_csv(sheet,connectionsheet,localstoragesheet, to_file):
    
    print('Creating ansible playbook   =====>           {}'.format(to_file))    
    scriptCode = []
    scriptCode.append("---"                                                                                                                                      )
    scriptCode.append("- name:  Creating server profiles from csv"                                                                                               )    
    build_header(scriptCode)

    ##
    scriptCode.append("  tasks:"                                                                                                                                )
    sheet.dropna(how='all', inplace=True)
    sheet       = sheet.applymap(str)                       # Convert data frame into string

    for i in sheet.index:
        row                         = sheet.loc[i]

        name                        = row['name']
        description                 = row['description']
        serverProfileTemplateUri    = row['serverProfileTemplateUri']   # in CSV, it's a name instead of URI
        serverHardwareUri           = row['serverHardwareUri']          # in CSV, it's a name instead of URI
        scope                       = row['scope']

        if 'nan'  == serverHardwareUri:
            serverHardwareUri       = 'unassigned'


        # Get server hardware URI from name
        var_hardware_uri            = 'unassigned'
        if 'unassigned' != serverHardwareUri.lower():
            var_hardware_name        = serverHardwareUri.strip().replace(',', '_').replace('-', '_').replace(' ', '') 
            var_hardware_uri         = "var_{}_uri".format(var_hardware_name)
            scriptCode.append("                                                     "                                                                       )
            scriptCode.append("     - name: Get server hardware uri   {}            ".format(serverHardwareUri)                                             )
            scriptCode.append("       oneview_server_hardware_facts:                "                                                                       )
            scriptCode.append("         config:             \'{{ config }}\'        "                                                                       )
            scriptCode.append("         name:                \'{}\'                 ".format(serverHardwareUri)                                             )    
            scriptCode.append("     - set_fact:                                     "                                                                       )
            scriptCode.append("         {}:                 ".format(var_hardware_uri) + "\'{{server_hardwares.uri}}\' "                                    )  
            scriptCode.append("                                                     "                                                                       )                



            # Power off the server 
            scriptCode.append("     - name: Power off server {}                     ".format(serverHardwareUri)                                                 )
            scriptCode.append("       oneview_server_hardware:                      "                                                                           )
            scriptCode.append("         config:     \'{{ config }}\'                "                                                                           )
            scriptCode.append("         state:      power_state_set                 "                                                                           )
            scriptCode.append("         data:                                       "                                                                           )
            scriptCode.append("             name:                           \'{}\'  ".format(serverHardwareUri)                                                 )
            scriptCode.append("             powerStateData:                         "                                                                           ) 
            scriptCode.append("                 powerState: \'Off\'                 "                                                                           )
            scriptCode.append("                 powerControl: \'MomentaryPress\'    "                                                                           )
            scriptCode.append("                                                     "                                                                           )

        if 'nan' == serverProfileTemplateUri:
            # Create standalone profile
            scriptCode.append("     - name: Create server profile {0}               ".format(name)                                                              )
            scriptCode.append("       oneview_server_profile:                       "                                                                           )
            scriptCode.append("         config:     \'{{ config }}\'                "                                                                           )
            scriptCode.append("         state:      present                         "                                                                           )
            scriptCode.append("         data:                                       "                                                                           )
            scriptCode.append("             name:                           \'{}\'  ".format(name)                                                              )
            scriptCode.append("             type:                           \'{}\'  ".format(rstype['serverprofile'])                                           )
            if 'nan' != description:
                scriptCode.append("             description:                    \'{}\'      ".format(description)                                               )
            
            # Add attributes to profile
            generate_profile_or_template(row, scriptCode)
            # Addd network connections and storage
            isProfile   = True
            generate_connection_storage_for_profile(connectionsheet,localstoragesheet,name,isProfile,scriptCode) 

        # Create profile from template
        else:               
            var_template_name           = serverProfileTemplateUri.strip().replace(',', '_').replace('-', '_').replace(' ', '') 
        
            # Get template uri from name
            scriptCode.append("                                                     "                                                                           )
            scriptCode.append("     - name: Get profile template uri   {}           ".format(serverProfileTemplateUri)                                          )
            scriptCode.append("       oneview_server_profile_template_facts:        "                                                                           )
            scriptCode.append("         config:             \'{{ config }}\'        "                                                                           )
            scriptCode.append("         name:                \'{}\'                 ".format(serverProfileTemplateUri)                                          )     
            scriptCode.append("     - set_fact:                                     "                                                                           )
            scriptCode.append("         var_{}_uri:   ".format(var_template_name) + "\'{{server_profile_templates[0].uri}}\' "                                  )    
            scriptCode.append("                                                     "                                                                           )


            # Create server profile from template
            scriptCode.append("     - name: Create server profile {0} from template {1} ".format(name, serverProfileTemplateUri)                                )
            scriptCode.append("       oneview_server_profile:                       "                                                                           )
            scriptCode.append("         config:     \'{{ config }}\'                "                                                                           )
            scriptCode.append("         state:      present                         "                                                                           )
            scriptCode.append("         data:                                       "                                                                           )
            scriptCode.append("             name:                           \'{}\'  ".format(name)                                                              )
            scriptCode.append("             type:                           \'{}\'  ".format(rstype['serverprofile'])                                           )
            if 'nan' != description:
                scriptCode.append("             description:                    \'{}\'      ".format(description)                                               )

            var_template_uri                = "var_{}_uri".format(var_template_name) 
            scriptCode.append("             serverProfileTemplateUri:       \'{{"   + "{}".format(var_template_uri)  + "}}\'"                                   )
            scriptCode.append("             serverHardwareUri:              \'{{"   + "{}".format(var_hardware_uri)  + "}}\'"                                   )
            scriptCode.append("                                                     "                                                                           )


        
        # Add scope
        if 'nan' != scope:
            netVar              = 'var_' + name.lower().strip().replace(',', '_').replace('-', '_').replace(' ', '') 
            scriptCode.append("                                                             "                                                   )
            scriptCode.append("     - name: get server profile  {}                          ".format(name)                                      )
            scriptCode.append("       oneview_server_profile_facts:                         "                                                   )
            scriptCode.append("         config:     \'{{config}}\'                          "                                                   )
            scriptCode.append("         name:       \'{}\'                                  ".format(name)                                      )
            scriptCode.append("     - set_fact:                                             "                                                   )
            scriptCode.append("          {}: ".format(netVar)  + "\'{{server_profiles[0].uri}}\' "                                              )

            netUri              = "\'{{" + '{}'.format(netVar) + "}}\'"  
            generate_scope_for_resource(name, netUri, scope, scriptCode)



        # Power on the server 
        if 'unassigned' != serverHardwareUri.lower():
            scriptCode.append("     - name: Power on server {}                     ".format(serverHardwareUri)                                                  )
            scriptCode.append("       oneview_server_hardware:                      "                                                                           )
            scriptCode.append("         config:     \'{{ config }}\'                "                                                                           )
            scriptCode.append("         state:      power_state_set                 "                                                                           )
            scriptCode.append("         data:                                       "                                                                           )
            scriptCode.append("             name:                           \'{}\'  ".format(serverHardwareUri)                                                 )
            scriptCode.append("             powerStateData:                         "                                                                           ) 
            scriptCode.append("                 powerState: \'On\'                  "                                                                           )
            scriptCode.append("                 powerControl: \'MomentaryPress\'    "                                                                           )
            scriptCode.append("                                                     "                                                                           )


    
    # end of server profile
    scriptCode.append("       delegate_to: localhost                    "                                                                                       )
    scriptCode.append(CR)

    # ============= Write scriptCode ====================
    write_to_file(scriptCode, to_file)


# ================================================================================================
#
#   MAIN
#
# ================================================================================================


if __name__ == "__main__":
    _home = (os.environ['SYNERGY_AUTO_HOME'])

    # Read excel file
    if len(sys.argv) >= 2:
        excelfile        = sys.argv[1]
        excelfile        = _home + '/csv/' + excelfile
    else:
        print('No Excel file specified. Exiting now....')
        sys.exit()

    print(CR)




    #xl = pd.read_excel(excelfile, None)
    xl                      = pd.ExcelFile(excelfile)
    xlsheets                = dict()
    for sheet in xl.sheet_names:
        sheet_name          = sheet.lower()
        if 'version' == sheet_name:
            xlsheets['version']                 = pd.read_excel(excelfile, sheet   ,comment='#' , dtype=str)
        if 'composer' in sheet_name:
            xlsheets['composer']                = pd.read_excel(excelfile, sheet   ,comment='#' , dtype=str)
            
            

        if 'ilo' in sheet_name:
            xlsheets['ilo']                     = pd.read_excel(excelfile, sheet   ,comment='#' , dtype=str)
        if 'timelocale' in sheet_name:
            xlsheets['timelocale']              = pd.read_excel(excelfile, sheet   ,comment='#' , dtype=str)
        if 'backup' in sheet_name:
            xlsheets['backupconfig']            = pd.read_excel(excelfile, sheet   ,comment='#' , dtype=str)
        if 'firmware' in sheet_name:
            xlsheets['firmware']                = pd.read_excel(excelfile, sheet   ,comment='#' , dtype=str)
        if 'snmpv1' in sheet_name:
            xlsheets['snmpv1']                  = pd.read_excel(excelfile, sheet   ,comment='#' , dtype=str)
        if 'addresspool' in sheet_name:
            xlsheets['addresspool']             = pd.read_excel(excelfile, sheet   ,comment='#' , dtype=str)

        if 'scope' in sheet_name:
            xlsheets['scope']                   = pd.read_excel(excelfile, sheet   ,comment='#' , dtype=str)
        if 'ethernetnetwork' in sheet_name:
            xlsheets['ethernetnetwork']         = pd.read_excel(excelfile, sheet   ,comment='#' , dtype=str)
        if 'fcnetwork' in sheet_name:
            xlsheets['fcnetwork']               = pd.read_excel(excelfile, sheet   ,comment='#' , dtype=str)
        if 'networkset' in sheet_name:
            xlsheets['networkset']              = pd.read_excel(excelfile, sheet   ,comment='#' , dtype=str)
        if 'logicalinterconnectgroup' in sheet_name:
            xlsheets['logicalinterconnectgroup'] = pd.read_excel(excelfile, sheet  ,comment='#' , dtype=str)
        if 'uplinkset' in sheet_name:
            xlsheets['uplinkset']               = pd.read_excel(excelfile, sheet   ,comment='#' , dtype=str)   
        if 'enclosuregroup' in sheet_name:
            xlsheets['enclosuregroup']          = pd.read_excel(excelfile, sheet   ,comment='#' , dtype=str)
        if 'logicalenclosure' in sheet_name:
            xlsheets['logicalenclosure']        = pd.read_excel(excelfile, sheet   ,comment='#' , dtype=str)
        
        if 'profiletemplate' == sheet_name:
            xlsheets['profiletemplate']         = pd.read_excel(excelfile, sheet   ,comment='#' , dtype=str)
        if 'templateconnection' in sheet_name:
            xlsheets['templateconnection']      = pd.read_excel(excelfile, sheet   ,comment='#' , dtype=str)
        if 'templatelocalstorage' in sheet_name:
            xlsheets['templatelocalstorage']    = pd.read_excel(excelfile, sheet   ,comment='#' , dtype=str)

        if 'profile' == sheet_name:
            xlsheets['profile']                 = pd.read_excel(excelfile, sheet   ,comment='#' , dtype=str)
        if 'profileconnection' in sheet_name:
            xlsheets['profileconnection']       = pd.read_excel(excelfile, sheet   ,comment='#' , dtype=str)
        if 'profilelocalstorage' in sheet_name:
            xlsheets['profilelocalstorage']    = pd.read_excel(excelfile, sheet   ,comment='#' , dtype=str)

    #  OneView sconfig json
    print('#---------------- Generate Oneview config json')
    prefix  = generate_ansible_configuration(                       xlsheets['composer'] , xlsheets['version'], _home + '/configFiles/oneview_config.json')

    # Create YML folder
    if not os.path.isdir(prefix):
        os.makedirs(prefix)

    ymlFolder   = _home + '/configFiles' + '/prefix/'

    # Connect to new OneView instance to collect interconnect types information
    print(CR)
    print('#---------------- Connect to Oneview instance')
    config_file = _home + '/configFiles/oneview_config.json'
    with open(config_file) as json_data:
        config = json.load(json_data)
    oneview_client = OneViewClient(config)

    # load resource type
    print('X-API version used ---> {}'.format(config["api_version"]) )
    rstype = resource_type_ov4_20 

    print(CR)
    print('#---------------- Query interconnect types ' )
    interconnect_types = oneview_client.interconnect_types.get_all(sort='name:descending')
    #print(json.dumps(interconnect_types, indent = 4))



    # OneView settings
    print(CR)
    print('#---------------- Generate playbooks for Oneview settings')
    generate_firmware_bundle_ansible_script_from_csv(               xlsheets['firmware']        , _home + '/configFiles/' + prefix +'firmwarebundle.yml')
    generate_time_locale_ansible_script_from_csv(                   xlsheets['timelocale']      , _home + '/configFiles/' + prefix +'timelocale.yml')
    generate_id_pools_ipv4_ranges_subnets_ansible_script_from_csv(  xlsheets['addresspool']     , _home + '/configFiles/' + prefix +'addresspool.yml')
    generate_snmp_v1_ansible_script_from_csv(                       xlsheets['snmpv1']          , _home + '/configFiles/' + prefix +'snmpv1.yml')

    print(CR)
    print('#---------------- Generate playbooks for Oneview resources')
    # OneView resources
    generate_scopes_ansible_script_from_csv(                        xlsheets['scope'] ,                 _home + '/configFiles/' + prefix +'scope.yml')
    generate_ethernet_networks_ansible_script_from_csv(             xlsheets['ethernetnetwork'] ,       _home + '/configFiles/' + prefix +'ethernetnetwork.yml')
    generate_fc_fcoe_networks_ansible_script_from_csv(              xlsheets['fcnetwork'] ,             _home + '/configFiles/' + prefix +'fcnetwork.yml')
    generate_network_sets_ansible_script_from_csv(                  xlsheets['networkset']      ,       _home + '/configFiles/' + prefix +'networkset.yml')
    generate_logical_interconnect_groups_ansible_script_from_csv(   xlsheets['logicalinterconnectgroup'] , xlsheets['uplinkset'], _home + '/configFiles/' + prefix +'logicalinterconnectgroup.yml')
    generate_enclosure_groups_ansible_script_from_csv(              xlsheets['enclosuregroup'],         _home + '/configFiles/' + prefix +'enclosuregroup.yml')
    generate_logical_enclosures_ansible_script_from_csv(            xlsheets['logicalenclosure'],       _home + '/configFiles/' + prefix +'logicalenclosure.yml')

    generate_server_profile_templates_ansible_script_from_csv(      xlsheets['profiletemplate'], xlsheets['templateconnection'], xlsheets['templatelocalstorage'], _home + '/configFiles/' + prefix +'profiletemplate.yml')
    generate_server_profiles_ansible_script_from_csv(               xlsheets['profile'],         xlsheets['profileconnection'],  xlsheets['profilelocalstorage'], _home + '/configFiles/' + prefix +'profile.yml')

    print(CR)
