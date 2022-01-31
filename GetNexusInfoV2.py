#!/bin/env python
import cisco, json, datetime
from cisco import clid
#from nxos import *

#make switch name using switch mgmt ip and replacing '.' with '_' and appending a timestamp
sw_name = str(json.loads(clid('show int mgmt0 brief'))["TABLE_interface"]["ROW_interface"]["ip_addr"]).replace('.','_')

#make timestamp using current datetime on the switch in format YYYY_MM_DD_HH_MM_SS  - example 2022_01_29_17_33_00
timestamp = str(datetime.datetime.now()).replace(" ", "_").replace("-", "_").replace(".", "_").replace(":", "_")[:19]

#make json output file name
json_file = sw_name + '_op_' + timestamp + '.json'


sh_ver = json.loads(clid('show version'))
sh_mod = json.loads(clid('show module'))
sh_inv = json.loads(clid('show inventory'))
sh_env = json.loads(clid('show environment'))

final_json = json.dumps({"sh_ver": sh_ver, "sh_mod": sh_mod, "sh_inv": sh_inv, "sh_env": sh_env})

print final_json
