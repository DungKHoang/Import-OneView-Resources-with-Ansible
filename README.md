# Generate ansible playbooks from Excel

import-ov-resource-with-ansible.py is a python script that generates ansible playbooks to configure OneView resources and settings from an Excel file.
The Excel file provides OV setting values and OV resources values.
The script generates the following ansible playbooks:

    * addresspool.yml
    * firmwarebundle.yml
    * snmpv1.yml
    * timelocale.yml

    * scope.yml
    * ethernetnetwork.yml
    * fcnetwork.yml
    * networkset.yml
    * logicalinterconnectgroup.yml (with uplinkset)
    * enclosuregroup.yml
    * logicalenclosure.yml
    * profiletemplate.yml (with local storage and network connections)
    * profile.yml (with local storage and network connections)


    Note: The playbooks work with OneView 4.20

## Prerequisites
    * Virtual machine running Ubuntu 18.09
    * python 2.7.15+
    * ansible 2.8.4
    * pandas library ( for reading/writing Excel files)
    * pip 
    * oneview-python SDK 5.0.0
    * oneview-ansible library 

## Setup
In the Ubuntu machine, perform the following operations:

    * sudo apt upgrade
    * sudo apt install python-pip
    * sudo pip install pandas
    * sudo pip install xlrd
    * sudo pip install requests

    * Install oneview python SDK
        **  git clone https://github.com/HewlettPackard/python-hpOneView.git --single-branch -b release/5.0.0-beta
        ** 	cd python-hpOneView
	    ** sudo python setup.py install 
        ** pip install -r requirements.txt
    
    * Install oneview ansible library
        ** git clone  https://github.com/HewlettPackard/oneview-ansible.git --single-branch  -b enhancement/pass_by_name/network_set
	    ** cd oneview-ansible
        ** sudo pip install -r requirements.txt   
    
    * Create folders from your home directory
        ** bin              --> python script
        ** configFiles      --> YML files and oneview_config.json
        ** csv              --> Excel configuration file
        ** iso              --> Synergy SPP iso
    
## Configuration
Create the following environment variables

    * export ANSIBLE_LIBRARY=/home-folder/oneview-ansible/library/
    * export ANSIBLE_MODULE_UTILS=/home-folder/oneview-ansible/library/module_utils/
    * export SYNERGY_AUTO_HOME=home-folder

## OV configuration Excel file
In the Excel file, configure the following tabs:

    * Version          
        ** I:12     ---> pod mumber used to create prefix for YML files
        ** I:13     ---> site name used to create prefix for YML files
    YML files will br created with prefix <site>-<POD>

    * Composer_OneView
        ** Ip           --> IP address of the OneView instance. The OneView instance needs to be online as the python script will connect to it to collect information
        ** userName     --> administrator name
        ** password     --> administrator password
    
    * iLO               - NOT USED
    * TimeLocale
    * Backup_config     - NOT USED
    * firmwareBundle
    * snmpV1
    * AddressPool
    * Scope
    * EthernetNetwork
    * FCNetwork
    * NetworkSet
    * LogicalInterconnectGroup
    * UplinkSet
    * EnclosureGroup
    * LogicalEnclosure
    * ProfileTemplate
    * Profile
    * ProfileTemplateConnection
    * ProfileTemplateLOCALStorage
    * ProfileConnection
    * ProfileLOCALStorage



## To generate ansible playbooks

```
    import-ov-resource-with-ansible.py ov-configuration.xlsx

```

## To run ansible playbooks

```
    ansible-playbook configFiles/<YML files>

```
