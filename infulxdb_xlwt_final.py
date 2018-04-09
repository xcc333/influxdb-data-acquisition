
import xlwt
from influxdb import InfluxDBClient

#have connection with influxdb
def connection(SQL_message):
    client = InfluxDBClient('10.203.96.26', 8086, 'influx', 'patac2016', 'patac_eim') 
    result = client.query(SQL_message)
    for data_origial in result:
        data = data_origial
    return data
#transfor a list to a two-dimention list
def list_transfor(data):
    b=[]
    c=[]
    for i in range(len(data)):
        for k in data[i]:
            b.append(data[i][k])

    for i in range(0,len(data)*len(data[0]),len(data[0])):
        c=c+[b[i:i+len(data[0])]]
    return c
#export data into a excel
def output_excel(data,sheet_name,start_no):
    a=[]
    table=workbook.add_sheet(sheet_name,cell_overwrite_ok=True)
    # add first column----serial number
    for i in range(len(data)):
        table.write(0,0,'序号',style)
        table.write(i+1,0,i+start_no,style)
    # add first row------channel title    
    for j in range(len(data[0])):
        for k in data[0]:
            a.append(k)
            table.write(0,j+1,a[j],style)
            
    c=list_transfor(data)
    #write data into cells
    for i in range(len(data)):
        for j in range(len(data[0])):
            table.write(i+1,j+1,c[i][j],style)
    
if __name__== '__main__':
    
    workbook=xlwt.Workbook()
    
    style = xlwt.XFStyle() 
    font=xlwt.Font()
    font.bold=False
    font.italic=True
    font.name='Calibr'
    style.font=font
    #using SQL parameters to get different data
    data1=connection('select * from "AVL_engine_durability" where "eqpt_no"=\'PEC0-5W03\' and time > \'2018-04-04 11:40:01\' limit 60000')
    output_excel(data1,'CAN_AO_1',1)
    
    data=connection('select * from "AVL_engine_durability" where "eqpt_no"=\'PEC0-5W03\' and time > \'2018-04-05 06:00:00\' limit 60000' )
    output_excel(data,'CAN_AO_2',60001)
    #export data in one excel file
    workbook.save(r'D:\sgmuserprofile\shstm8\Desktop\influxdb_5W03.xls')
    