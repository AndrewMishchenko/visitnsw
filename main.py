import os
import time

import xlwt
from selenium import webdriver

path = os.path.dirname(__file__) + '/geckodriver'


class Parser(object):
    def __init__(self):
        self.driver = webdriver.Firefox(executable_path=path)
        # print(path)

    def next_page(self):
        time.sleep(3)
        times = 0
        while times <= 60:
            try:
                self.driver.find_element_by_css_selector(
                    'button#atlas_ajax').click()
            except Exception:
                break
            times += 1
            print(times)
        return

    def close_popup(self):
        try:
            self.driver.find_element_by_css_selector(
                'span.close').click()
        except Exception:
            pass
        return

    def get_all_href(self):
        hrefs = self.driver.find_elements_by_css_selector('a.image')
        hrefs = [href.get_attribute('href') for href in hrefs]
        return hrefs

    def get_page_content(self):
        try:
            name = self.driver.find_element_by_css_selector(
                'h1.main-heading').text
        except Exception:
            name = ''

        try:
            site = self.driver.find_element_by_css_selector(
                '.side-box').find_element_by_css_selector(
                'a.button').get_attribute('href')
        except Exception:
            site = ''

        try:
            email = self.driver.find_element_by_css_selector(
                '.side-box').find_element_by_css_selector(
                'a.email').get_attribute('href').split(':')[-1]
        except Exception:
            email = ''

        try:
            tel = self.driver.find_element_by_css_selector(
                'span.tel').text
        except Exception:
            tel = ''

        try:
            facebook = self.driver.find_element_by_css_selector(
                'a.ga_URL_lead_facebook').get_attribute('href')
        except Exception:
            facebook = ''

        try:
            twitter = self.driver.find_element_by_css_selector(
                'a.ga_URL_lead_twitter').get_attribute('href')
        except Exception:
            twitter = ''

        return {
            'name': name,
            'site': site,
            'email': email,
            'tel': tel,
            'facebook': facebook,
            'twitter': twitter
        }


if __name__ == '__main__':
    parser = Parser()
    parser.driver.get('http://www.visitnsw.com/events/search')
    parser.close_popup()
    parser.next_page()
    hrefs = parser.get_all_href()

    index = 1
    doc = xlwt.Workbook('result.xls')
    sheet = doc.add_sheet('sheet1')

    for href in hrefs:
        try:
            parser.driver.get(href)
            page = parser.get_page_content()

            sheet_row = sheet.row(index)
            sheet_row.write(0, page['name'])
            sheet_row.write(1, page['site'])
            sheet_row.write(2, page['email'])
            sheet_row.write(3, page['tel'])
            sheet_row.write(4, page['facebook'])
            sheet_row.write(5, page['twitter'])
            doc.save('result.xls')
            print(index)
            index += 1
        except Exception:
            continue
