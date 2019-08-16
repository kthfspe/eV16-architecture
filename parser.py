import xml.etree.ElementTree as ET
import csv
import xlsxwriter



Verbose = 0 # 1: To print processing details or not, 0: For minimal
id = dict()
ecus = dict()
sensors = dict()
actuators = dict()
hmi = dict()
otsc = dict()
enclosures = dict()
connectors = dict()
npc = dict()
pc = dict()
swc = dict()
f_cansignals = dict()
f_analogsignals = dict()
f_digitalsignals = dict()
f_internalsignals = dict()
f_lvsignals = dict()
f_hvsignals = dict()
f_gndsignals = dict()
f_signals = dict()

p_cansignals = dict()
p_analogsignals = dict()
p_digitalsignals = dict()
p_internalsignals = dict()
p_lvsignals = dict()
p_hvsignals = dict()
p_gndsignals = dict()
p_signals = dict()

idcount = 1
ecucount = 1
sensorcount = 1
actuatorcount = 1
otsccount = 1
enclosurecount = 1
hmicount = 1
npccount = 1
pccount = 1
swccount = 1
f_cansignalcount = 1
f_digitalsignalcount = 1
f_analogsignalcount = 1
f_internalsignalcount = 1
f_hvsignalcount = 1
f_lvsignalcount = 1
f_gndsignalcount = 1
f_signalcount = 1

p_cansignalcount = 1
p_digitalsignalcount = 1
p_analogsignalcount = 1
p_internalsignalcount = 1
p_hvsignalcount = 1
p_lvsignalcount = 1
p_gndsignalcount = 1
p_signalcount = 1


# Read data store by BlockType
# Each block is stored in its own dict.
# All signals are stored in one dict
def readid(a):
    global id
    global idcount
    b = dict((k, a[k]) for k in ('Name', 'BlockType', 'id', 'Parent'))
    id.update({idcount:b})
    idcount += 1

def p_readsignal(child):
    global p_signals
    global p_signalcount
    global id
    s = dict()
    b = dict((k, child.attrib[k]) for k in ('Name', 'BlockType')) #Add from connector to connector
    for item in id:
        if id[item]['id'] == child[0].attrib['source']:
            s = {'Source' : id[item]['Name']}
        if id[item]['id'] == child[0].attrib['target']:
            t = {'Target' : id[item]['Name']}
    z = b.copy()
    z.update(s)
    d = z.copy()
    d.update(t)
    p_signals.update({p_signalcount:d})
    p_signalcount += 1

def f_readsignal(child):
    global f_signals
    global f_signalcount
    global id

    b = dict((k, child.attrib[k]) for k in ('Name', 'BlockType')) #Add from connector to connector
    for item in id:
        if id[item]['id'] == child[0].attrib['source']:
            s1 = {'Source' : id[item]['Name']}
        if id[item]['id'] == child[0].attrib['target']:
            t = {'Target' : id[item]['Name']}
    z = b.copy()
    z.update(s1)
    d = z.copy()
    d.update(t)
    f_signals.update({f_signalcount:d})
    f_signalcount += 1

root = ET.parse('physical.xml').getroot()
for child in root.findall('diagram/mxGraphModel/root/object'):
    a = dict(child.attrib)
    #print(child.attrib["BlockType"])
    if child.attrib["BlockType"].lower() == 'sen':
        if sensorcount == 1:
            sensors.update({sensorcount:a})
            sensorcount += 1
        else:
            for item in range(1,sensorcount):
                if (sensors[item]["Name"].lower() != child.attrib["Name"].lower()):
                    sensors.update({count:a})
                    sensorcount += 1
        readid(a)

    elif child.attrib["BlockType"].lower() == 'act':
        if actuatorcount == 1:
            actuators.update({actuatorcount:a})
            actuatorcount += 1
        else:
            for item in range(1,actuatorcount):
                if (actuators[item]["Name"].lower() != child.attrib["Name"].lower()):
                    actuators.update({actuatorcount:a})
                    actuatorcount += 1
        readid(a)

    elif child.attrib["BlockType"].lower() == 'hmi':
        if hmicount == 1:
            hmi.update({hmicount:a})
            hmicount += 1
        else:
            for item in range(1,hmicount):
                if (hmi[item]["Name"].lower() != child.attrib["Name"].lower()):
                    hmi.update({hmicount:a})
                    hmicount += 1
        readid(a)

    elif child.attrib["BlockType"].lower() == 'otsc':
        if otsccount == 1:
            otsc.update({otsccount:a})
            otsccount += 1
        else:
            for item in range(1,otsccount):
                if (otsc[item]["Name"].lower() != child.attrib["Name"].lower()):
                    otsc.update({otsccount:a})
                    otsccount += 1
        readid(a)


    elif child.attrib["BlockType"].lower() == 'ecu':
        if ecucount == 1:
            ecus.update({ecucount:a})
            ecucount += 1
        else:
            for item in range(1,ecucount):
                if (ecus[item]["Name"].lower() != child.attrib["Name"].lower()):
                    ecus.update({ecucount:a})
                    ecucount += 1
        readid(a)

    elif child.attrib["BlockType"].lower() == 'enc':
        if enclosurecount == 1:
            enclosures.update({enclosurecount:a})
            enclosurecount += 1
        else:
            for item in range(1,enclosurecount):
                if (enclosures[item]["Name"].lower() != child.attrib["Name"].lower()):
                    enclosures.update({enclosurecount:a})
                    enclosurecount += 1
        readid(a)

    elif child.attrib["BlockType"].lower() == 'con':
        if connectorcount == 1:
            connectors.update({connectorcount:a})
            connectorcount += 1
        else:
            for item in range(1,connectorcount):
                if (connectors[item]["Name"].lower() != child.attrib["Name"].lower()):
                    connectors.update({count:a})
                    connectorcount += 1
        readid(a)


root = ET.parse('functional.xml').getroot()
Verbose = 0 # 1: To print processing details or not, 0: For minimal

# Read data store by BlockType
# Each block is stored in its own dict.
# All signals are stored in one dict

for child in root.findall('diagram/mxGraphModel/root/object'):
    a = dict(child.attrib)

    if child.attrib["BlockType"].lower() == 'npc':
        if npccount == 1:
            npc.update({npccount:a})
            npccount += 1
        else:
            for item in range(1,npccount):
                if (npc[item]["Name"].lower() != child.attrib["Name"].lower()):
                    npc.update({npccount:a})
                    npccount += 1
        readid(a)

    elif child.attrib["BlockType"].lower() == 'pc':
        if pccount == 1:
            pc.update({pccount:a})
            pccount += 1
        else:
            for item in range(1,pccount):
                if (pc[item]["Name"].lower() != child.attrib["Name"].lower()):
                    pc.update({pccount:a})
                    pccount += 1
        readid(a)

    elif child.attrib["BlockType"].lower() == 'swc':
        if swccount == 1:
            swc.update({swccount:a})
            swccount += 1
        else:
            for item in range(1,swccount):
                if (swc[item]["Name"].lower() != child.attrib["Name"].lower()):
                    swc.update({swccount:a})
                    swccount += 1
        readid(a)



for child in root.findall('diagram/mxGraphModel/root/object'):
    a = dict(child.attrib)

    if child.attrib["BlockType"].lower() == 'can':
        f_cansignals.update({f_cansignalcount:a})
        f_cansignalcount += 1
        f_readsignal(child)

    elif child.attrib["BlockType"].lower() == 'dig':
        f_digitalsignals.update({f_digitalsignalcount:a})
        f_digitalsignalcount += 1
        f_readsignal(child)

    elif child.attrib["BlockType"].lower() == 'ana':
        f_analogsignals.update({f_analogsignalcount:a})
        f_analogsignalcount += 1
        f_readsignal(child)

    elif child.attrib["BlockType"].lower() == 'iv':
        f_internalsignals.update({f_internalsignalcount:a})
        f_internalsignalcount += 1
        f_readsignal(child)

    elif child.attrib["BlockType"].lower() == 'lvs':
        f_lvsignals.update({f_lvsignalcount:a})
        f_lvsignalcount += 1
        f_readsignal(child)

    elif child.attrib["BlockType"].lower() == 'hvs':
        f_hvsignals.update({f_hvsignalcount:a})
        f_hvsignalcount += 1
        f_readsignal(child)

    elif child.attrib["BlockType"].lower() == 'gnd':
        f_gndsignals.update({f_gndsignalcount:a})
        f_gndsignalcount += 1
        f_readsignal(child)

root = ET.parse('physical.xml').getroot()
for child in root.findall('diagram/mxGraphModel/root/object'):
    a = dict(child.attrib)
    if child.attrib["BlockType"].lower() == 'can':
        p_cansignals.update({p_cansignalcount:a})
        p_cansignalcount += 1
        p_readsignal(child)

    elif child.attrib["BlockType"].lower() == 'dig':
        p_digitalsignals.update({p_digitalsignalcount:a})
        p_digitalsignalcount += 1
        p_readsignal(child)

    elif child.attrib["BlockType"].lower() == 'ana':
        p_analogsignals.update({p_analogsignalcount:a})
        p_analogsignalcount += 1
        p_readsignal(child)

    elif child.attrib["BlockType"].lower() == 'iv':
        p_internalsignals.update({p_internalsignalcount:a})
        p_internalsignalcount += 1
        p_readsignal(child)

    elif child.attrib["BlockType"].lower() == 'lvs':
        p_lvsignals.update({p_lvsignalcount:a})
        p_lvsignalcount += 1
        p_readsignal(child)

    elif child.attrib["BlockType"].lower() == 'hvs':
        p_hvsignals.update({p_hvsignalcount:a})
        p_hvsignalcount += 1
        p_readsignal(child)

    elif child.attrib["BlockType"].lower() == 'gnd':
        p_gndsignals.update({p_gndsignalcount:a})
        p_gndsignalcount += 1
        p_readsignal(child)

# If checks passed, export to CSV, print total number of signals, units etc.


# Else print error message


workbook = xlsxwriter.Workbook('SysArch.xlsx')
idsheet = workbook.add_worksheet('ID')
for item in id:
    col = 0
    for key in id[item].keys():
        idsheet.write(0, col, key)
        idsheet.write(item, col, id[item][key])
        col += 1

f_signalsheet = workbook.add_worksheet('Signals P')
for item in f_signals:
    col = 0
    for key in f_signals[item].keys():
        f_signalsheet.write(0, col, key)
        f_signalsheet.write(item, col, f_signals[item][key])
        col += 1
p_signalsheet = workbook.add_worksheet('Signals F')
for item in p_signals:
    col = 0
    for key in p_signals[item].keys():
        p_signalsheet.write(0, col, key)
        p_signalsheet.write(item, col, p_signals[item][key])
        col += 1

ecusheet = workbook.add_worksheet('ECUs')
for item in ecus:
    col = 0
    for key in ecus[item].keys():
        ecusheet.write(0, col, key)
        ecusheet.write(item, col, ecus[item][key])
        col += 1

sensorsheet = workbook.add_worksheet('Sensors')
for item in sensors:
    col = 0
    for key in sensors[item].keys():
        sensorsheet.write(0, col, key)
        sensorsheet.write(item, col, sensors[item][key])
        col += 1

actuatorsheet = workbook.add_worksheet('Actuators')
for item in actuators:
    col = 0
    for key in actuators[item].keys():
        actuatorsheet.write(0, col, key)
        actuatorsheet.write(item, col, actuators[item][key])
        col += 1

hmisheet = workbook.add_worksheet('HMI')
for item in hmi:
    col = 0
    for key in hmi[item].keys():
        hmisheet.write(0, col, key)
        hmisheet.write(item, col, hmi[item][key])
        col += 1

otscsheet = workbook.add_worksheet('OTSC')
for item in otsc:
    col = 0
    for key in otsc[item].keys():
        otscsheet.write(0, col, key)
        otscsheet.write(item, col, otsc[item][key])
        col += 1

connectorsheet = workbook.add_worksheet('Connectors')
for item in connectors:
    col = 0
    for key in connectors[item].keys():
        connectorsheet.write(0, col, key)
        connectorsheet.write(item, col, connectors[item][key])
        col += 1

npcsheet = workbook.add_worksheet('NPC')
for item in npc:
    col = 0
    for key in npc[item].keys():
        npcsheet.write(0, col, key)
        npcsheet.write(item, col, npc[item][key])
        col += 1

pcsheet = workbook.add_worksheet('PC')
for item in pc:
    col = 0
    for key in pc[item].keys():
        pcsheet.write(0, col, key)
        pcsheet.write(item, col, pc[item][key])
        col += 1

swcsheet = workbook.add_worksheet('SWC')
for item in swc:
    col = 0
    for key in swc[item].keys():
        swcsheet.write(0, col, key)
        swcsheet.write(item, col, swc[item][key])
        col += 1

dig = workbook.add_worksheet('Digital Signals')
for item in p_digitalsignals:
    col = 0
    for key in p_digitalsignals[item].keys():
        dig.write(0, col, key)
        dig.write(item  , col, p_digitalsignals[item][key])
        col += 1
net = item
for item in f_digitalsignals:
    col = 0
    for key in f_digitalsignals[item].keys():
        dig.write(0, col, key)
        dig.write(item+net  , col, f_digitalsignals[item][key])
        col += 1

analog = workbook.add_worksheet('Analog Signals')
for item in p_analogsignals:
    col = 0
    for key in p_analogsignals[item].keys():
        analog.write(0, col, key)
        analog.write(item  , col, p_analogsignals[item][key])
        col += 1
net = item
for item in f_analogsignals:
    col = 0
    for key in f_analogsignals[item].keys():
        analog.write(0, col, key)
        analog.write(item+net  , col, f_analogsignals[item][key])
        col += 1

can = workbook.add_worksheet('CAN')
for item in f_cansignals:
    col = 0
    for key in f_cansignals[item].keys():
        can.write(0, col, key)
        can.write(item  , col, f_cansignals[item][key])
        col += 1
net = item
for item in p_cansignals:
    col = 0
    for key in p_cansignals[item].keys():
        can.write(0, col, key)
        can.write(item+net  , col, p_cansignals[item][key])
        col += 1

internal = workbook.add_worksheet('Internal Signals')
for item in f_internalsignals:
    col = 0
    for key in f_internalsignals[item].keys():
        internal.write(0, col, key)
        internal.write(item  , col, f_internalsignals[item][key])
        col += 1
net = item
for item in p_internalsignals:
    col = 0
    for key in p_internalsignals[item].keys():
        internal.write(0, col, key)
        internal.write(item+net  , col, p_internalsignals[item][key])
        col += 1
lv = workbook.add_worksheet('LV Supply')
for item in f_lvsignals:
    col = 0
    for key in f_lvsignals[item].keys():
        lv.write(0, col, key)
        lv.write(item  , col, f_lvsignals[item][key])
        col += 1
net = item
for item in p_lvsignals:
    col = 0
    for key in p_lvsignals[item].keys():
        lv.write(0, col, key)
        lv.write(item+net  , col, p_lvsignals[item][key])
        col += 1

hv = workbook.add_worksheet('HV Supply')
for item in p_hvsignals:
    col = 0
    for key in p_hvsignals[item].keys():
        hv.write(0, col, key)
        hv.write(item  , col, p_hvsignals[item][key])
        col += 1
net = item
for item in f_hvsignals:
    col = 0
    for key in f_hvsignals[item].keys():
        hv.write(0, col, key)
        hv.write(item+net  , col, f_hvsignals[item][key])
        col += 1

gnd = workbook.add_worksheet('GND')
for item in f_gndsignals:
    col = 0
    for key in f_gndsignals[item].keys():
        gnd.write(0, col, key)
        gnd.write(item , col, f_gndsignals[item][key])
        col += 1
net = item
for item in p_gndsignals:
    col = 0
    for key in p_gndsignals[item].keys():
        gnd.write(0, col, key)
        gnd.write(item+net  , col, p_gndsignals[item][key])
        col += 1
workbook.close()
#print digitalsignals
