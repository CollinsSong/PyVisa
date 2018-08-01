#ID-VD testing

import  xdrlib ,sys
import xlrd
from xlwt import Workbook
from xlrd import xldate_as_tuple
import visa
import time
wb=Workbook()
sheet1=wb.add_sheet('Sheet 1')
visa_dll = 'c:/windows/system32/visa32.dll'
gpib_addr1 = 'GPIB0::9::INSTR'
gpib_addr2 = 'GPIB0::22::INSTR'
rm = visa.ResourceManager(visa_dll)
gpib_inst1 = rm.open_resource(gpib_addr1)#PS
gpib_inst2 = rm.open_resource(gpib_addr2)#DMM
gpib_inst1.timeout=5000
gpib_inst2.timeout=5000
m=1
n=1                                                                                                                       
gpib_inst1.write('INST:SEL OUT2')
gpib_inst1.write('VOLT 4')                      #1.Set VG starting value
gpib_inst1.write('VOLT:STEP 1')                 #2.Set VG changing interval

while n<=3:                                     #3.setting VG ending value
        gpib_inst1.write('INST:SEL OUT2')
        print("Now starting the measurements of VG=" + gpib_inst1.query('VOLT?'))
        gpib_inst1.write('INST:SEL OUT1')
        gpib_inst1.write('VOLT 0')              #4.Set VD starting value
        gpib_inst1.write('VOLT:STEP 0.1')       #5.Set VD changing interval
        
        gpib_inst1.write('INST:SEL OUT2')
        sheet1.write(0,0+2*n,float(gpib_inst1.query('VOLT?')))
        
        while m<52:                             #6.Set VD ending value

                gpib_inst1.write('INST:SEL OUT1')
                print("VD=" + gpib_inst1.query('VOLT?'))
                #print("ID=" + gpib_inst2.query('MEASure:CURRent:DC?'))
                sheet1.write(m,0+2*n,float(gpib_inst1.query('VOLT?')))
                sheet1.write(m,1+2*n,float(gpib_inst2.query('MEASure:CURRent:DC? MIN,MIN')))
                gpib_inst1.write('VOLT UP')
                m=m+1
        else:
                    gpib_inst1.write('INST:SEL OUT2') 
                    print("Finished measurments of VG =" + gpib_inst1.query('VOLT?'))
                
        gpib_inst1.write('INST:SEL OUT2')           
        gpib_inst1.write('VOLT UP')
        m=1
        n=n+1
        gpib_inst1.write('INST:SEL OUT1')
        gpib_inst1.write('VOLT 0')
        time.sleep(2)
else:
        print("Finished")

gpib_inst1.write('INST:SEL OUT1')
gpib_inst1.write('VOLT 0')
gpib_inst1.write('INST:SEL OUT2')
gpib_inst1.write('VOLT 0')
print("Error 2:" + gpib_inst2.query('SYSTem:ERRor?'))
print("Error 1:" + gpib_inst1.query('SYSTem:ERRor?'))
wb.save('ID-VD testing precise.xls')

#tcp_addr = 'TCPIP::192.168.1.1::inst0::INSTR'



