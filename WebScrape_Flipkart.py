from urllib.request import urlopen as uReq
from bs4 import BeautifulSoup as soup
import xlsxwriter
import locale

my_url = 'https://www.flipkart.com/home-kitchen/home-appliances/refrigerators/double-door~type/pr?sid=j9e%2Cabm%2Chzg&otracker=nmenu_sub_TVs%20and%20Appliances_0_Double%20Door'

uClient2 = uReq(my_url)
page_html = uClient2.read()
uClient2.close()

page_soup = soup(page_html, "html.parser")
#locale.setlocale(locale.LC_ALL, 'en_US.UTF8')


containers11 = page_soup.find_all("div", {"class": "_25b18c"})
specific = page_soup.find_all("div", {"class": "col col-7-12"})
EMO = page_soup.find_all("div", {"class": "fMghEO"})

workbook = xlsxwriter.Workbook('Refrigerators.xlsx')

bold_format = workbook.add_format({'bold':True})

cell_format = workbook.add_format()
#cell_format.set_text_wrap()
cell_format.set_align('top')
cell_format.set_align('left=')
cell_format.set_font_size(7)

#moneyFormat = workbook.add_format({'num_format': '$#,##0.00'})


#moneyRedFormat.set_font_color('red')
#moneyRedFormat1 = workbook.add_format({'num_format': '#,##0.00'})
#moneyRedFormat1.set_font_color('green')





worksheet = workbook.add_worksheet('Refrigerators')



#chart1 = workbook.add_chart({'type': 'bar'})

worksheet.write('A1', 'Product', bold_format)
worksheet.write('B1', 'Price',bold_format)
worksheet.write('C1', 'Actual_Price', bold_format)
worksheet.write('D1', 'Detail_Specification', bold_format)


rowIndex = 1

for price, container, price1, EMI in zip(containers11, specific, containers11, EMO):
    #title_container = container.find_all("span", {"class": "B_NuCI"})
    #product_name = title_container[0].text


    try:
        price_con = price.find_all("div", {"class": "_30jeq3 _1_WHN1"})
        price = price_con[0].text




    except:
        price = 'No Data'

    #print("Product: " + product_name)
    print("Price: " + price)

    #f.write(price.replace(",", "|") + "\n")


#f.close()

#for speci in specific:
    try:
        product_name1 = container.find_all("div", {"class": "_4rR01T"})
        product_name = product_name1[0].text

    except:
        'No Spec'

    try:
        actual_price1 = price1.find_all("div", {"class": "_3I9_wc _27UcVY"})
        actual_price = actual_price1[0].text
    except:
        'No Spec'

    try:
        EMI_options = EMI.find_all("ul", {"class": "_1xgFaf"})
        emp = EMI_options[0].text
    except:
        price = 'No Data'

    #print("EMI: " + emp)

    #print("Product: " + product_name)
    #print("Price: " + price)



    rowIndex+=1



    worksheet.write('A' + str(rowIndex), product_name.split()[0].lower(),cell_format)
    worksheet.write('B' + str(rowIndex), price)
    worksheet.write('C' + str(rowIndex),actual_price,cell_format)
    worksheet.write('D' + str(rowIndex), emp,cell_format)
        #worksheet.ignore_errors({'number_stored_as_text': 'B1:XFD1048576'})

chart1= workbook.add_chart({'type': 'bar'})
chart1.add_series({
    'name': '=Refrigerators!$A$1',
    'categories': '=Refrigerators!$A$2:$A$25',
    'values': '=Refrigerators!$B$2:$B$25',
        'gap': 10,
})



chart1.set_title({'name': 'Price Analysis'})
chart1.set_x_axis({'name': 'Price'})
chart1.set_y_axis({'name': 'Product'})
chart1.set_style(11)
#chart1.drawing.top = 50
#chart1.drawing.left = 100
#chart1.drawing.width = 50
#chart1.drawing.height = 20

worksheet.insert_chart('O2', chart1)


workbook.close()
