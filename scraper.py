import requests
import json
import xlsxwriter
import time
from datetime import datetime

# Import the byte stream handler.
from io import BytesIO

# Import urlopen() for either Python 2 or 3.
try:
    from urllib.request import urlopen
except ImportError:
    from urllib2 import urlopen

from forex_python.converter import CurrencyRates


##########################################
# Function to Add row entry
##########################################
def addEntry(worksheet,row,details,image_data):
    print "Adding",details[1],"to row",row
    col=0
    for item in details:
        worksheet.write(row, col, item)
        col+=1

    # Write the byte stream image to a cell. Note, the filename must be
    # specified. In this case it will be read from url string.
    worksheet.insert_image(row, col, 'Image Data', {'image_data': image_data})

##########################################
# Function to return image data
##########################################
def getimg(targetURL):
    #print targetURL
    image_data = BytesIO(urlopen(targetURL).read())
    return image_data
    

##########################################
# Main Function
##########################################
# sheet = openWorksheet()
# initializeWorksheet(sheet)
try:
    print "Opening Excel File"
    workbook = xlsxwriter.Workbook('kickstarter.xlsx')
    print "Adding Excel Sheet"
    worksheet = workbook.add_worksheet()

    print "Initializing Worksheet"
    bold = workbook.add_format({'bold': True})
    worksheet.write(0, 0, "Rank",bold)
    worksheet.write(0, 1, "Name",bold)
    worksheet.write(0, 2, "Creator's Name",bold)
    worksheet.write(0, 3, "Goal (USD)",bold)
    worksheet.write(0, 4, "Pledged Amount (USD)",bold)
    worksheet.write(0, 5, "Percentage Fulfilled",bold)
    worksheet.write(0, 6, "Status",bold)
    worksheet.write(0, 7, "Backers",bold)
    worksheet.write(0, 8, "Launch Date",bold)
    worksheet.write(0, 9, "Deadline",bold)
    worksheet.write(0, 10, "URL",bold)
    worksheet.write(0, 11, "Duration(days)",bold)
    worksheet.write(0, 12, "Funding Goal vs. Time",bold)
except:
    print "Unable to Open Excel File. Please close the Excel File if already open."
else:
    page = 1
    total = 0
    ctr = 1
    
    try:
        pages = int(raw_input("Enter Number of pages to scrape (1 page = 12 entries): "))
    except:
        print "Please enter Valid Numeric value"
    else:
        if pages<=0:
            print "Enter Valid number of pages"
        else:
            while page<=pages:
                print "Collecting Data for page",page
                r = requests.get('https://www.kickstarter.com/discover/advanced.json?category_id=334&woe_id=0&sort=popularity&seed=2508285&page=' + str(page))
                #print r.status_code
                if r.status_code!=200:
                    print "Connection Error! Status:",r.status_code
                    break
                data = r.json()

                total+=len(data["projects"])
                for index in range(len(data["projects"])):
                    #print "%s,%f,%s,%f" % (data["projects"][index]["name"], data["projects"][index]["goal"], data["projects"][index]["currency"], data["projects"][index]["pledged"])
                    details = []

                    details.append(ctr)
                    details.append(data["projects"][index]["name"])
                    details.append(data["projects"][index]["creator"]["name"])
                    
                    cur_conv = CurrencyRates()
                    currency = data["projects"][index]["currency"]
                    if not currency == 'USD':
                        goal_usd = cur_conv.convert(currency, 'USD', int(data["projects"][index]["goal"]))
                        pledged_usd = cur_conv.convert(currency, 'USD', int(data["projects"][index]["pledged"]))
                    else:
                        # No need to convert since already in USD.
                        goal_usd = int(data["projects"][index]["goal"])
                        pledged_usd = int(data["projects"][index]["pledged"])

                    details.append(goal_usd)
                    details.append(pledged_usd)
                    details.append(float(data["projects"][index]["pledged"]/data["projects"][index]["goal"])*100)
                        
                    launch_date = time.strftime('%c', time.localtime(data["projects"][index]["launched_at"]))
                    deadline = time.strftime('%c', time.localtime(data["projects"][index]["deadline"]))

                    details.append(data["projects"][index]["state"])
                    details.append(data["projects"][index]["backers_count"])
                    details.append(launch_date)
                    details.append(deadline)
                    details.append(data["projects"][index]["urls"]["web"]["project"])

                    #durationInfo(data["projects"][index]["urls"]["web"]["project"])
                    l_date = datetime.strptime(launch_date,'%c')
                    d_date = datetime.strptime(deadline,'%c')
                    details.append(abs((l_date-d_date).days))

                    slug = data["projects"][index]["slug"]
                    proj_id = data["projects"][index]["creator"]["id"]
                    tracker_url = 'http://www.kicktraq.com/projects/'+str(proj_id)+'/'+str(slug)+'/dailychart.png'
                    img = getimg(tracker_url)

                    addEntry(worksheet,ctr,details,img)
                    ctr+=1

                print ""
                page+=1

    try:
        print "Closing Workbook"
        workbook.close()
        print "Closed Workbook"
    except:
        print "Unable to close Excel File. Please close the file if open."
        print "Entries not written to file"
    else:
        print "Added",total,"entries"
