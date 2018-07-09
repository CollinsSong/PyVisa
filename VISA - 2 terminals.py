import  xdrlib ,sys
import xlrd
from xlwt import Workbook
from xlrd import xldate_as_tuple
import visa

wb=Workbook()
sheet1=wb.add_sheet('Sheet 1')

visa_dll = 'c:/windows/system32/visa32.dll'
#tcp_addr = 'TCPIP::192.168.1.1::inst0::INSTR'

gpib_addr1 = 'GPIB0::9::INSTR'
gpib_addr2 = 'GPIB0::22::INSTR'
# Create an object of visa_dll
rm = visa.ResourceManager(visa_dll)

# Create an instance of certain interface(GPIB and TCPIP)

gpib_inst1 = rm.open_resource(gpib_addr1)
gpib_inst2 = rm.open_resource(gpib_addr2)


#gpib_inst1.write('VOLTage {<voltage>|2|4}')
gpib_inst1.write('INST:SEL OUT1')
gpib_inst1.write('VOLT 4')
#gpib_inst1.write('OUTPUT:ON')

gpib_inst1.write('INST:SEL OUT2')
gpib_inst1.write('VOLT 0')
#gpib_inst1.write('OUTPut:ON')
gpib_inst1.write('VOLT:STEP 0.1')
gpib_inst2.write('CURRent:DC:NPLC 100')
#print("NPLC:" + gpib_inst2.query('CURRent:DC:NPLC?'))
n=1
while n<11:
    #print("OUTPUT:" + gpib_inst1.read())
    #have some problems on NPLC setting
    gpib_inst2.write('CURRent:DC:NPLC 100')
    print("OUTPUT 2 Voltage:" + gpib_inst1.query('VOLT?'))
    print("Current:" + gpib_inst2.query('MEASure:CURRent:DC?'))
    sheet1.write(n,0,float(gpib_inst1.query('VOLT?')))
    sheet1.write(n,1,1000000*float(gpib_inst2.query('MEASure:CURRent:DC?')))
    gpib_inst1.write('VOLT UP')
    n=n+1
else:
    print("finished")


gpib_inst1.write('INST:SEL OUT1')
print(gpib_inst1.query('VOLT?'))
gpib_inst1.write('INST:SEL OUT2')
print(gpib_inst1.query('VOLT?'))

wb.save('testing.xls')

#print(gpib_inst1.query('*IDN?'))
#print(gpib_inst1.query('MEASure:VOLTage:DC?'))
#print(gpib_inst2.query('*IDN?'))
#print(gpib_inst2.query('MEASure:VOLTage:DC?'))

# Using write()/read()/query() function to make communication with device

# according to the command type
#gpib_inst1.write("*IDN?")
#print  ("OUTPUT:" + gpib_inst1.read())




