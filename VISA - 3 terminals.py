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

m=1#channel 1
n=1#channel 2

#gpib_inst1.write('VOLTage {<voltage>|2|4}')
gpib_inst1.write('INST:SEL OUT1')
gpib_inst1.write('VOLT 0')
gpib_inst1.write('VOLT:STEP 1')
#gpib_inst1.write('OUTPUT:ON')
while m<=4:
        print("OUTPUT 1 Voltage:" + gpib_inst1.query('VOLT?'))  
        gpib_inst1.write('INST:SEL OUT2')
        gpib_inst1.write('VOLT 0')
        #gpib_inst1.write('OUTPut:ON')
        gpib_inst1.write('VOLT:STEP 0.1')
        #gpib_inst2.write('CURRent:DC:NPLC 100')
        #print("NPLC:" + gpib_inst2.query('CURRent:DC:NPLC?'))
        gpib_inst1.write('INST:SEL OUT1')
        sheet1.write(0,0+2*m,float(gpib_inst1.query('VOLT?')))
        while n<42:
                #print("OUTPUT:" + gpib_inst1.read())
                #have some problems on NPLC setting
                gpib_inst1.write('INST:SEL OUT2')
                gpib_inst2.write('CURRent:DC:NPLC 100')
                print("OUTPUT 2 Voltage:" + gpib_inst1.query('VOLT?'))
                print("Current:" + gpib_inst2.query('MEASure:CURRent:DC?'))
                sheet1.write(n,0+2*m,float(gpib_inst1.query('VOLT?')))
                sheet1.write(n,1+2*m,1000000*float(gpib_inst2.query('MEASure:CURRent:DC?')))
                gpib_inst1.write('VOLT UP')
                n=n+1
        else:
                    gpib_inst1.write('INST:SEL OUT1') 
                    print("Finished measurments of output 1 voltage =" + gpib_inst1.query('VOLT?'))
                
        gpib_inst1.write('INST:SEL OUT1')   
        
        gpib_inst1.write('VOLT UP')
        print("OUTPUT 1 Voltage:" + gpib_inst1.query('VOLT?'))  
        m=m+1
        n=1
        gpib_inst1.write('INST:SEL OUT2')
        gpib_inst1.write('VOLT 0')
else:
        print("Finished")

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
#gpib_inst1.timeout = 1000
#ms
# according to the command type
#gpib_inst1.write("*IDN?")
#print  ("OUTPUT:" + gpib_inst1.read())




