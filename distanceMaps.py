import googlemaps, openpyxl
from openpyxl.styles import Font, Alignment

gmaps = googlemaps.Client(key='YOUR_GOOGLE_MAPS_API_KEY')

#open an excel file.

wb=openpyxl.load_workbook('database.xlsx')
sheet=wb.active
num=0
#loop by all the rows

for i in range(2,(sheet.max_row+1)):
    
#read the cell with departure place adress.

    org=sheet['H'+str(i)].value
    
#read the cell with arrival place adress.
    
    dst=sheet['I'+str(i)].value
    
#put the variables to distance matrix function, and assign to dictionary.
    
    try:
        matrix=gmaps.distance_matrix(org, dst, mode='driving')
    
#take the distace from dict.
    
        distValue=matrix['rows'][0]['elements'][0]['distance']['value']
    
#put distace value into excel.

        sheet['K'+str(i)]=round((distValue/1000),2)
    except:
        sheet['J'+str(i)]='Puste'
    sheet['K'+str(i)].alignment=Alignment(horizontal='center',wrapText=True,vertical='center')
    sheet['K'+str(i)].number_format = '0.00'
    sheet['L'+str(i)].number_format = '0.00'
    num+=1
    
#save na excel file.
    
wb.save('Database_with_Distance.xlsx')
wb.close()
print('Gotowe!')
