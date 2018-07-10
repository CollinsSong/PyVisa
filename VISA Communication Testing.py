import visa
visa_dll = 'c:/windows/system32/visa32.dll'
#tcp_addr = 'TCPIP::192.168.1.1::inst0::INSTR'

gpib_addr1 = 'GPIB0::9::INSTR'
gpib_addr2 = 'GPIB0::22::INSTR'
# Create an object of visa_dll
rm = visa.ResourceManager(visa_dll)

# Create an instance of certain interface(GPIB and TCPIP)

gpib_inst1 = rm.open_resource(gpib_addr1)
gpib_inst2 = rm.open_resource(gpib_addr2)


gpib_inst1.write("*IDN?")
print  ("OUTPUT:" + gpib_inst1.read())

gpib_inst2.write("*IDN?")
print  ("OUTPUT:" + gpib_inst2.read())

