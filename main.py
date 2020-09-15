import csv
import os
from datetime import date
import requests
from bs4 import BeautifulSoup as BS
import xlrd
from validate_email import validate_email
import re

time = date.today()

# vars for index, need to know the number of iterations
index_email = 1
index_phone = 1
index_website = 1
index_address = 1


def name_editor():  # name editing function
    xls_file = xlrd.open_workbook('mls.xls', formatting_info=True)  # read the file
    sheet = xls_file.sheet_by_index(0)
    column = sheet.col_values(3, 1)  # get the data column which we need
    characters = ["(", ")", ".", ",", "&", "'"]  # all characters to delete
    ready_names = []
    for engname in column:
        engname.strip()
        for character in characters:  # characters deleting loop
            engname = engname.replace(character, '').replace(' ', '-')
            engname = engname.replace('--', '-')  # if name has two space
        ready_names.append(engname)  # add name to ready_names after editing
    return ready_names  # return the list with ready names in english


def get_mlr():  # get MLR_number from excel file
    file = xlrd.open_workbook('mls.xls', formatting_info=True)  # read the file
    sheet = file.sheet_by_index(0)
    mlr_list = sheet.col_values(0, 1)  # column with MLR_no in excel has index 0
    return mlr_list  # returns list


def full_links(lists):  # make full url
    url = 'https://www.dimloan.com/MoneyLender/'  # part of url
    links = []
    for name in lists:  # iteration to get the page name one at  one
        full_url = url + name
        links.append(full_url)
    return links  # returns a finished list with all links


def get_html(url):  # function for request
    response = requests.get(url)
    return response.text  # retuns html text


def get_page_email(html, mlr_no):  # get the email from page
    global index_email  # add the global var which already exists
    soup = BS(html, 'html.parser')
    try:
        # returns all contacts information(email,website, phone number) as list
        # example - ['example.com', '2222 1111', 'example@mail.com']
        contact = soup.find('div', {'id': 'contact'}).find_all('div', {'class': 'value'})
        email = contact[2].text  # get only email with index 2

        # if email is correct or exist returns true
        is_valid = validate_email(email)
        if is_valid:
            email_data = {
                'MLR_no': mlr_no,
                'email': email,
                'Time-Scrapped': time,
            }
            index_email += 1  # enumerator our index

            # dict with ready data
            return email_data
        else:
            return None
    except:
        return None  # if get emptiness or if this page don't exist , return None


def get_page_website(html, mlr_no):
    global index_website  # add the global var which already exists
    soup = BS(html, 'html.parser')
    try:

        contact = soup.find('div', {'id': 'contact'}).find_all('div', {'class': 'value'})
        website = contact[0].text  # website data has index 0
        # use regular expression to search website string
        is_valid = re.search(r'\w+\.+\w', website)
        # if page haven't  website then there is character 沒
        if is_valid is not None and '沒' not in website:
            website_data = {
                'MLR_no': mlr_no,
                'Website': website,
                'Time-Scrapped': time,
            }
            index_website += 1  # enumerator our index
            return website_data
        else:
            return None
    except:
        return None


def get_page_phone(html, mlr_no):
    global index_phone  # add the global var which already exists
    soup = BS(html, 'html.parser')
    try:
        contact = soup.find('div', {'id': 'contact'}).find_all('div', {'class': 'value'})
        phone_number = contact[1].text

        # check the phone number for validity.
        if phone_number.replace(' ', '').isnumeric():  # returns True, if it's number
            phone_data = {
                'MLR_no': mlr_no,
                'Phone': phone_number,
                'Time-Scrapped': time,
            }
            index_phone += 1  # enumerator our index
            return phone_data
        else:  # if it isn't number returns None
            return None
    except:  # if he don't found the page
        return None


def get_page_address(html, mlr_no):
    global index_address  # add the global var which already exists
    soup = BS(html, 'html.parser')
    all_addresses = []  # empty list for addresses

    try:
        # the page shows the number of company addresses
        # 財務公司註冊地址數目
        #
        count = soup.find('div', {'id': 'highValue'}).find('div', {'class': 'value'}).text
        if count == '1':  # if the company has only one address
            address = soup.find('div', {'id': 'addrBack'}).text
            address_data = {
                'MLR_no': mlr_no,
                'index_address': 1,
                'Address': address,
                'Time-Scrapped': time,
            }
            index_address += 1  # enumerator our index
            return address_data  # returns dict with  data


        elif count is None:  # if we didn't  find this information
            return None
        else:  # if the company has more than one address

            # find_all returns all addresses as list
            for address in soup.find('div', {'id': 'addrBack'}).find_all('div'):

                if address is None:
                    return None
                else:
                    address = address.get_text()
                    all_addresses.append(address)

            addresses_data = {
                'MLR_no': mlr_no,
                'index_address': index_address,
                'Address': all_addresses,
                'Time-Scrapped': time,
            }
            return addresses_data
    except:
        return None


# function to write csv file with address
def write_address(data):
    global index_address
    try:
        with open('MLR_Address.csv', 'a') as file:
            writer1 = csv.writer(file)  # if data have more addresses
            writer2 = csv.DictWriter(file, fieldnames=list(data.keys()))  # if returned one address
            if index_address == 2:  # for the head columns
                writer2.writeheader()  # write columns names
            if type(data['Address']) == list:

                # get list with addresses from dict wih key  'Address'
                for full_address in data['Address']:
                    address = full_address.split('.')
                    # uses split to split up this string and get the number of address
                    # address[0] is number of address and write to csv file
                    # uses func join() to get to get the rest of string
                    rest_of_address = '.'.join(address[1:len(address)])
                    writer1.writerow((data['MLR_no'],
                                      address[0],
                                      rest_of_address,
                                      data['Time-Scrapped']))

                    index_address += 1  # enumerator  index
                    # !!! it's important, in the loop index +1 and it uses as row
            else:
                writer2.writerow(data)

            print('to 2544', index_address)  # prints an index to show when parsing ends
    except:  # do nothing if data=None or other situation
        pass


# to write data to files  'MLR_Email.csv', 'MLR_Website.csv', 'MLR_Phone.csv'
def write_contacts(data, file_name):
    try:
        with open(file_name, 'a') as file:
            writer = csv.DictWriter(file, fieldnames=list(data.keys()))

            if file_name == 'MLR_Phone.csv' and index_phone == 2:  #
                # if creates new file  then write columns header
                writer.writeheader()
            elif file_name == 'MLR_Website.csv' and index_website == 2:
                writer.writeheader()
            elif file_name == 'MLR_Email.csv' and index_email == 2:
                writer.writeheader()
            writer.writerow(data)
    except:
        pass


def main():  # main func to run scripts

    # list of some files to write
    file_names = ['MLR_Email.csv',
                  'MLR_Website.csv',
                  'MLR_Phone.csv']

    all_links = full_links(name_editor())  # runs a script to get all the links
    # loop to receive data in order

    for url, mlr_number in zip(all_links, get_mlr()):
        html = get_html(url)  # request and returned html
        # calls all functions
        # call func to write MLR_phone
        write_contacts(get_page_phone(html, mlr_number),
                       file_names[2])  # get name of files by index
        # calls func to write  MLR_website
        write_contacts(get_page_website(html, mlr_number), file_names[1])
        # calls function to write MLR_email
        write_contacts(get_page_email(html, mlr_number), file_names[0])
        # calls function to write MLR_addresses
        write_address(get_page_address(html, mlr_number))  # calls the func


if __name__ == '__main__':
    main()
