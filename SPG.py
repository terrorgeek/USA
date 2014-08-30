#########################################################################
#########################################################################
#
#  spg.py
#
#  Updated by: Jason Cluck (jcluck@gwmail.gwu.edu)
#
#  Description: searches spg.com for availability in the cash and points
#  area
#
#  Input: The text file locations.txt is required but not passed in as an argument.  Locations.txt iterates through all of the cities and states listed,
#  iterating through each day in dates.xls for every city.
#
#  Output: city.xls, where city is a list of hardcoded cities. Each Excel
#  file has the following columns: check-in date, hotel name, cash, and
#  points price.
#
#  Procedure: the program accesses spg.com and checks their cash and
#  points price from the results for a given check-in date and city name.
#
#########################################################################
##########################################################################
#spg.py
#Hotel Reconnaissance Agent

from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
import time
from datetime import date, timedelta, datetime
import xlrd
import xlwt
import selenium.webdriver.support.ui as ui
#acquire the current time!
now=time.strftime('%m/%d/%y',time.localtime(time.time()))
def main():
    #user input - commented out for running as a chronjob. cities and states are pulled in through a text file
    # city = str(raw_input("City: "))
    # state =str(raw_input("State: "))
    # country = str(raw_input("Country: "))
    # limit = int(raw_input("Limit search to first how many hotels? "))

    country = "United States"
    limit = 12

    #open up the weekend_days excel book
    book = xlrd.open_workbook("dates.xls")
    sheet = book.sheet_by_index(0)

    #start firefox web driver - requires firefox
    browser = webdriver.Firefox()
    wait = ui.WebDriverWait(browser,30)

   # try:
    #main branch that uses the workbook and iterates through the dates inside of it
    #iterate through the rows of weekend_days.xls
    #open the text file
    file1=open("locations.txt").read().split('; ')
    #for each city in the text file...
    for location in file1:
        print location
        #create a workbook
        wbook = xlwt.Workbook()
        #add a new sheet that will have the hotel rates
        wsheet = wbook.add_sheet("Hotel Rates")
        wsheet.write(0, 0, "Check-in Date")
        wsheet.write(0, 1, "Cash Only")
        wsheet.write(0, 2, "Cash")
        wsheet.write(0, 3, "Points")
        wsheet.write(0, 4, "Hotel Name")

        #split up the city and state using a comma delimiter
        location = location.split(',')
        #get the city, should be the first argument (Fort Lauderdale, Florida)
        city = location[0]
        #get the state, second argument
        state = location[1]

        #now construct the search query using the city, state, and country
        search_query = city+", "+state+", "+country

        #setup the output file name (ie: New York.xls)
        output_filename = city + ".xls"

        #set up webdriver wait for timeout

        for row_number in xrange(0, sheet.nrows, 1):
            #pull the first date from the excel file and the organize the tuple like a date
            rowValues = sheet.row_values(row_number, 0, 3)
            dateA = "%d/%d/%d" % tuple(rowValues)
            #convert the date to a datetime object
            dateA = datetime.strptime(dateA, '%m/%d/%Y')

            #for the second date, just add a time delta of one day
            dateB = dateA + timedelta(days=1)

            #convert the dates back to strings for the SPG homepage
            dateA = dateA.strftime("%m/%d/%y")
            dateB = dateB.strftime("%m/%d/%y")
            if dateA<=now:
               continue
            #navigate to the homepage for starwoodhotels
            browser.get("http://www.starwoodhotels.com/preferredguest/index.html")

            #wait for the required fields to be ready before doing anything
            wait.until(lambda browser: browser.find_element_by_name("complexSearchField"))
            wait.until(lambda browser: browser.find_element_by_name("arrivalDate"))
            wait.until(lambda browser: browser.find_element_by_name("departureDate"))

            arrivalDate = browser.find_element_by_name("arrivalDate")
            arrivalDate.clear()


            departureDate = browser.find_element_by_name("departureDate")
            departureDate.clear()


            complexSearchField = browser.find_element_by_name("complexSearchField")
            complexSearchField.clear()
            complexSearchField.send_keys(search_query)
            arrivalDate.send_keys(dateA)
            departureDate.send_keys(dateB + Keys.RETURN)
            try:
                WebDriverWait(browser, 60).until(lambda browser : browser.find_element_by_link_text("Website Terms of Use"))
            except:
                browser.quit()
            cheapest_hotel = ""
            cheapest_hotel_cash_only = ""
            cheapest_starpoints = None
            cheapest_cash = None
            cheapest_cash_only = None
            cash_only = None

            WebDriverWait(browser, 60).until(lambda browser : browser.find_element_by_link_text("Compare Rates"))
            browser.find_element_by_link_text("Compare Rates").click()
            #set up the xpath to determine if the Compare Rates page has loaded
            xpath = "/html/body/div[2]/div[2]/div[2]/div[2]/div[2]/div/table/tbody/tr[2]/td/div/div[2]/div/a"
            wait.until(lambda browser: browser.find_element_by_xpath(xpath))
            #reset the cash/points flag
            CASH_AND_POINTS_FLAG = False
            #iterate through the table on the SPG site with the values
            for counter in range(2,(limit*2)+1,2):

                #Get the hotel name
                try:
                    xpath = "/html/body/div[2]/div[2]/div[2]/div[2]/div[2]/div/table/tbody/tr[%d]/td/div/div[2]/div/a" % counter
                    hotel_name = browser.find_element_by_xpath(xpath).text
                except:
                    print "No hotel path found"
                    continue

                #Cash
                try:
                    xpath = "/html/body/div[2]/div[2]/div[2]/div[2]/div[2]/div/table/tbody/tr[%d]/td[2]/div/div[@class='bookingLink']" % counter
                    #strip off the other stuff included in the link text and get the cash integer
                    cash_only = int(str(browser.find_element_by_xpath(xpath).text).strip(' >').strip('USD '))
                    if (cheapest_cash_only == None ) or (cash_only < cheapest_cash_only):
                        cheapest_hotel_cash_only = hotel_name
                        cheapest_cash_only = cash_only
                    elif cheapest_cash_only == cash_only:
                        cheapest_hotel_cash_only += ", " + hotel_name
                except:
                    print "No cash path found"


                #Cash & Points
                try:
                    xpath = "/html/body/div[2]/div[2]/div[2]/div[2]/div[2]/div/table/tbody/tr[%d]/td[4]/div/div[@class='currencyAmount']" % counter
                    spg_cash_and_points = browser.find_element_by_xpath(xpath)

                    value_vector = [int(s) for s in spg_cash_and_points.text.split() if s.isdigit()]
                    starpoints_value = value_vector[0]
                    cash = value_vector[1]
                    if (cheapest_starpoints == None) or (starpoints_value < cheapest_starpoints):
                        #set the flag so this is printed out and not the cash
                        CASH_AND_POINTS_FLAG = True
                        cheapest_hotel = hotel_name
                        cheapest_starpoints = starpoints_value
                        cheapest_cash = cash
                    elif starpoints_value == cheapest_starpoints:
                        cheapest_hotel += ", "+hotel_name
                except:
                    continue


            #print out the date range
            print dateA+" - "+dateB
            #print out the cheapest starpoints and cash value then save to .xls file
            print str(cheapest_hotel)+" ("+str(cheapest_starpoints)+" Starpoints)\n"

            #print out the cheapest cash only then save to .xls file
            print str(cheapest_hotel_cash_only.encode("utf-8"))+" ("+str(cheapest_cash_only).encode("utf-8")+" USD)\n"

            #write the data to the worksheet
            wsheet.write((((row_number))+1), 0, dateA)
            #if there was a cash and points option, print that
            if not CASH_AND_POINTS_FLAG:
                wsheet.write((((row_number))+1), 1, cheapest_cash_only)
                wsheet.write((((row_number))+1), 4, cheapest_hotel_cash_only)
            #if there wasn't a cash and points option, print the cash option
            else:
                wsheet.write((((row_number))+1), 2, cheapest_cash)
                wsheet.write((((row_number))+1), 3, cheapest_starpoints)
                wsheet.write((((row_number))+1), 4, cheapest_hotel)

            wbook.save(output_filename)


    #except:
        #browser.close()

if __name__ == "__main__":
    main()
