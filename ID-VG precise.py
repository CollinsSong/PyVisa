#ID-VG testing
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
gpib_inst1.timeout=5000
gpib_inst2.timeout=5000
m=1
n=1                                                                                                                       
gpib_inst1.write('INST:SEL OUT1')
gpib_inst1.write('VOLT 1')                      #1. Set VD starting value
gpib_inst1.write('VOLT:STEP 1')                 #2. Set VD changing interval
gpib_inst2.write('CURRent:DC:NPLC 100')

while m<=1:                                     #3. Set VD ending value
        gpib_inst1.write('INST:SEL OUT1')
        print("Now starting the measurements of VD=:" + gpib_inst1.query('VOLT?'))  
        gpib_inst1.write('INST:SEL OUT2')
        gpib_inst1.write('VOLT 0')              #4. Set VG starting value
        gpib_inst1.write('VOLT:STEP 0.1')       #5. Set VG changing interval
        gpib_inst1.write('INST:SEL OUT1')
        sheet1.write(0,0+2*m,float(gpib_inst1.query('VOLT?')))
        while n<52:                             #6. Set VG ending value

                gpib_inst1.write('INST:SEL OUT2')
                print("VG=" + gpib_inst1.query('VOLT?'))
                #print("ID=" + gpib_inst2.query('MEASure:CURRent:DC?'))
                sheet1.write(n,0+2*m,float(gpib_inst1.query('VOLT?')))
                sheet1.write(n,1+2*m,float(gpib_inst2.query('MEASure:CURRent:DC? MIN,MIN')))
                gpib_inst1.write('VOLT UP')
                n=n+1
        else:
                    gpib_inst1.write('INST:SEL OUT1') 
                    print("Finished measurments of VD =" + gpib_inst1.query('VOLT?'))
                
        gpib_inst1.write('INST:SEL OUT1')           
        gpib_inst1.write('VOLT UP')
        m=m+1
        n=1
        gpib_inst1.write('INST:SEL OUT2')
        gpib_inst1.write('VOLT 0')
else:
        print("Finished")

gpib_inst1.write('INST:SEL OUT1')
gpib_inst1.write('VOLT 0')
gpib_inst1.write('INST:SEL OUT2')
gpib_inst1.write('VOLT 0')

wb.save('ID-VG testing precise.xls')




