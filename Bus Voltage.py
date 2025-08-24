
import excelpy

SS_Name = 'Normal Condition Bus Voltage.xlsx'
SS_Event = 'Case name'
xl = excelpy.workbook(SS_Name,SS_Event,False)
xl.show_alerts(True)


xl.set_cell('A1', 'Bus Number',fontStyle='Bold')
xl.set_cell('C1', 'Bus Per-Unit Voltage',fontStyle='Bold')
xl.set_cell('E1', 'Bus Actual Voltage',fontStyle='Bold')

import psspy

#input bus number

buses=(181403,151411,101,301,201,401,501,601,701,181404,151408,801,151403,901,1001) 

for i,bus in enumerate(buses):
  err, voltage =  psspy.busdat(bus,"PU") 
  err1, actvoltage =  psspy.busdat(bus,"KV")  
  if err==0:
    xl.set_cell('A' + str(i+2), bus)
    xl.set_cell('C' + str(i+2), round (voltage,3))
    xl.set_cell('E' + str(i+2), round (actvoltage,1))
  else:
    xl.set_cell('A' + str(i+2), bus)
    xl.set_cell('C' + str(i+2), " ")
    xl.set_cell('E' + str(i+2), " ")

xl.save()
xl.show()

