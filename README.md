# cisco-switch-parse-interfaces

Summary:

Parse Cisco Catalyst switch running config and paste interface information into an Excel spreadsheet. One switch per 
sheet.

Specifics:

1) Tested primarily in virtual Cisco environment. 
2) Support for interfaces from Ethernet to 10Gb.
3) Able to recognize ether-channels, and their modes.
4) Parameters that are recorded in the spreadsheet: interface number, description, shut/no shut, type (access/trunk), 
   access vlan, voice vlan, allowed vlans, portfast, stpguard, link negotiation.
   
How to run:

1) Paste desired switch config in the parse_switch_config.py directory. Ensure that all switch files end with 
   _switchconfig.txt All files ending in with _switchconfig.txt will be parsed further.
2) Open terminal and run python3 parse_switch_config.py. Upon completion, Switches and Ports.xlsx file will be created 
   in the script directory.