#ID-VG Leakage disconnecting source
import  xdrlib ,sys
import xlrd
from xlwt import Workbook
from xlrd import xldate_as_tuple
import visa

wb=Workbook()
sheet1=wb.add_sheet('Sheet 1')
visa_dll = 'c:/windows/system32/visa32.dll'
gpib_addr1 = 'GPIB0::9::INSTR'
gpib_addr2 = 'GPIB0::22::INSTR'
rm = visa.ResourceManager(visa_dll)
gpib_inst1 = rm.open_resource(gpib_addr1)#PS
gpib_inst2 = rm.open_resource(gpib_addr2)#DMM
m=1
n=1
gpib_inst1.timeout=5000
gpib_inst2.timeout=5000
gpib_inst1.write('INST:SEL OUT1')
gpib_inst1.write('VOLT 0')                                                                                                                       
gpib_inst1.write('INST:SEL OUT2')
gpib_inst1.write('VOLT 0')#1. Set VG starting value
gpib_inst1.write('VOLT:STEP 0.1')#2. Set VG changing interval
#gpib_inst2.write('CURRent:DC:NPLC 100')

while n<=52: #3. Set VG ending value
        gpib_inst1.write('INST:SEL OUT2')
        print("VG=" + gpib_inst1.query('VOLT?'))
        #print("ID=" + gpib_inst2.query('MEASure:CURRent:DC?'))
        sheet1.write(n,1,float(gpib_inst1.query('VOLT?')))
        sheet1.write(n,2,float(gpib_inst2.query('MEASure:CURRent:DC? MIN,MIN')))
        gpib_inst1.write('VOLT UP')
        n=n+1
else:
        print("Finished measurments of Gate leakage" )
        
gpib_inst1.write('INST:SEL OUT2')
gpib_inst1.write('VOLT 0')
wb.save('Gate Leakage testing precise.xls')



