"""
Website credits to www.worldweatheronline.com. Parsing the data provided by them as it provides a good documentation of data.
"""
from selenium import webdriver
import xlsxwriter
from xlwt import Workbook
from calendar import monthrange
from datetime import date, datetime
import os
import sys

class WebAutomation():
    def __init__(self):
        self.firefox_path = "C:/Program Files/Mozilla Firefox"
        self.wb = Workbook()
        self.web = webdriver.Firefox(self.firefox_path)
        self.place = input("Enter the place you want to record the time for. Should be a single word. ")
        self.state = input("Enter the state to which the city belongs. ")
        self.country = input("Enter the country abbreviation. Example:  India=in, Germany=de, etc. ")
        self.link = "https://www.worldweatheronline.com/" + self.place + "-weather-history/"+self.state+"/" + self.country + ".aspx"
        self.web.get(self.link)

        try:
            year = int(input("Enter year for which dataset is to be created "))
        except:
            print("Error in conversion. Please enter only numbers")
            sys.exit(0)
        todayDate = date.today()
        td = todayDate.strftime("%d %m %Y").split(" ")
        for month in range(1, 13):
            for day in range(1, ((monthrange(year, (month))[1]) + 1)):
                self.date = datetime.strptime(str(day), '%d').strftime('%d')
                self.month = datetime.strptime(str(month), '%m').strftime('%m')
                self.year = datetime.strptime(str(year), '%Y').strftime('%Y')
                if td[0] == self.date and td[1] == self.month and td[2] == self.year:
                    print("System should exit right now")
                    self.web.close()
                    sys.exit(0)
                print(self.date, self.month, self.year)

                self.completeDate1 = self.year + "-" + self.month + "-" + self.date
                self.path = "output/" + self.year + self.month + "/"
                if not os.path.exists(self.path):
                    os.mkdir(self.path)
                self.filename = "output/" + self.year + self.month + "/" + self.place + "-" + self.completeDate1 + ".xlsx"

                self.automate()

    def automate(self):
        try:
            selectDate = self.web.find_element_by_xpath('//*[@id="ctl00_MainContentHolder_txtPastDate"]')
        except:
            print("Error 404: Location not found. Please try again with correct inputs")
            self.web.close()
            sys.exit(0)
        selectDate.send_keys(self.completeDate1)

        self.web.find_element_by_id('ctl00_MainContentHolder_butShowPastWeather').click()

        #Create a Excel workbook and worksheet
        workbook = xlsxwriter.Workbook(self.filename)
        worksheet = workbook.add_worksheet(self.completeDate1)

        count = 1
        for i in range(1, 289, 12):
            time_xpath = '//*[@id="aspnetForm"]/div[4]/main/div[4]/div[1]/div[3]/div/div[1]/div/div[2]/div/div['+str(i)+']'
            time = self.web.find_element_by_xpath(time_xpath).text

            i += 2
            temp_xpath = '//*[@id="aspnetForm"]/div[4]/main/div[4]/div[1]/div[3]/div/div[1]/div/div[2]/div/div[' + str(i) + ']'
            temp = self.web.find_element_by_xpath(temp_xpath).text

            i += 1
            feeltemp_xpath = '//*[@id="aspnetForm"]/div[4]/main/div[4]/div[1]/div[3]/div/div[1]/div/div[2]/div/div[' + str(i) + ']'
            feeltemp = self.web.find_element_by_xpath(feeltemp_xpath).text

            i += 1
            wind_xpath = '//*[@id="aspnetForm"]/div[4]/main/div[4]/div[1]/div[3]/div/div[1]/div/div[2]/div/div[' + str(i) + ']'
            wind = self.web.find_element_by_xpath(wind_xpath).text

            i += 1
            gust_xpath = '//*[@id="aspnetForm"]/div[4]/main/div[4]/div[1]/div[3]/div/div[1]/div/div[2]/div/div[' + str(i) + ']'
            gust = self.web.find_element_by_xpath(gust_xpath).text

            i += 1
            rain_xpath = '//*[@id="aspnetForm"]/div[4]/main/div[4]/div[1]/div[3]/div/div[1]/div/div[2]/div/div[' + str(i) + ']'
            rain = self.web.find_element_by_xpath(rain_xpath).text

            i += 1
            humidity_xpath = '//*[@id="aspnetForm"]/div[4]/main/div[4]/div[1]/div[3]/div/div[1]/div/div[2]/div/div[' + str(i) + ']'
            humidity = self.web.find_element_by_xpath(humidity_xpath).text

            i += 1
            cloud_xpath = '//*[@id="aspnetForm"]/div[4]/main/div[4]/div[1]/div[3]/div/div[1]/div/div[2]/div/div[' + str(i) + ']'
            cloud = self.web.find_element_by_xpath(cloud_xpath).text

            i += 1
            pressure_xpath = '//*[@id="aspnetForm"]/div[4]/main/div[4]/div[1]/div[3]/div/div[1]/div/div[2]/div/div[' + str(i) + ']'
            pressure = self.web.find_element_by_xpath(pressure_xpath).text

            #write the data to Excel file
            time_count = 'A'+str(count)
            temp_count = 'B'+str(count)
            feeltemp_count = 'C'+str(count)
            wind_count = 'D'+str(count)
            gust_count = 'E'+str(count)
            rain_count = 'F'+str(count)
            humidity_count = 'G'+str(count)
            cloud_count = 'H'+str(count)
            pressure_count = 'I'+str(count)

            worksheet.write(time_count, time)
            worksheet.write(temp_count, temp)
            worksheet.write(feeltemp_count, feeltemp)
            worksheet.write(wind_count, wind)
            worksheet.write(gust_count, gust)
            worksheet.write(rain_count, rain)
            worksheet.write(humidity_count, humidity)
            worksheet.write(cloud_count, cloud)
            worksheet.write(pressure_count, pressure)
            count += 1

            #Stop to the loop
            if time == "21:00":
                break
        try:
            workbook.close()
        except:
            print("Error: Occurred while writing to the Excel File. Close the file " + self.filename + " and run the code again.")
            sys.exit(0)

if __name__ == '__main__':
    WebAutomation()