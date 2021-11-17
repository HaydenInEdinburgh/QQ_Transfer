# Packages
import logging
import logging.handlers
import smtplib
import time

from selenium.common.exceptions import NoSuchElementException, WebDriverException
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium import webdriver
import csv
import pickle
import locale
locale.setlocale(locale.LC_ALL, 'en_US.UTF8')

class Param:
    def __init__(self):
        self.driver_path = 'geckodriver.exe'
        self.qq_index = 'https://app.qqcatalyst.com/'
        self.login_email = 'albertm@eigenanalytics.com'
        self.login_password = 'r8YUdNQfeWa7nuYX'
        #self.qq_data_path = 'waterstone_customers_2.csv'
        self.ori_data_path = 'waterstone_customers_records.csv'
        self.processed_customer_number_path = '1115_2021/processed_cust_num_set.txt'
        self.processed_policy_number_path = '1115_2021/processed_policy_num_set.txt'
        self.processed_customer_records_path = '1115_2021/processed_customers_records_11152021.csv'
        self.processed_policy_records_path = '1115_2021/processed_policy_records_11152021.csv'

class SeleniumWeb:
    def __init__(self, driver):
        self.driver = driver

    '''
    CHECK EXIST FUNCTIONS
    '''

    def check_exists_by_xpath(self, xpath):
        try:
            self.driver.find_element_by_xpath(xpath)
        except NoSuchElementException:
            return False
        return True

    def check_exists_by_id(self, element_id):
        try:
            self.driver.find_element_by_id(element_id)
        except NoSuchElementException:
            return False
        return True

    def check_clickable_by_xpath(self, xpath):
        try:
            element = self.driver.find_element_by_xpath(xpath)
            element.click()
        except WebDriverException:
            return False
        return True

    def get_search_box_xpath(self):
        return '//input[@name="SearchWords"]'

    def get_search_result_items_xpath(self):
        return '//ul[@id="search-result-item-template"]'

    def get_cust_number_xpath(self):
        return '//form[@id="AccountInfo"]//div[contains(@class, "SectionItem")]/div[1]/div//div[@class="value"]'

    def get_button_edit_xpath(self, is_package):
        if is_package:
            return '//form[@id="PackageAdditionalInfo"]/div[contains(@class, "SectionButtons")]/a[contains(@class, "section_edit")]'
        return '//form[@name="PolicyAdditional"]/div[contains(@class, "SectionButtons")]/a[contains(@class, "section_edit")]'

    def get_button_save_xpath(self, is_package):
        if is_package:
            return '//form[@name="PackageAdditionalInfo"]/div[contains(@class, "SectionButtons")]/a[contains(@class, "section_save")]'
        return '//form[@name="PolicyAdditional"]/div[contains(@class, "SectionButtons")]/a[contains(@class, "section_save")]'

    def get_original_agent_xpath(self):
        return '//form[@name="PolicyAdditional"]//label[contains(text(), "Agent")]/following-sibling::div[@class="value"]'

    def get_original_csr_xpath(self):
        return '//form[@id="AccountInfo"]//div[contains(@class, "SectionItem")]//label[@for="account-info-csr"]/following-sibling::div[@class="value"]'

    def get_no_results_xpath(self):
        return '/html/body//div[@class="search-results"]/div[contains(@class, "no-search-results")]'

    def get_policies_from_list_xpath(self):
        return '//form[@id="PolicyList"]/table[@class="subgrid"]/tbody/tr[contains(@class, "policyStatusEquals")]/td[contains(@class, "lob")]/div[2]'

    def get_policy_premium_xpath(self):
        return '//section[@id="overview-panel"]//div[contains(@class, "policypremium")]'

    def get_policy_number_xpath(self):
        return '//section[@id="overview-panel"]//div[contains(@class, "policyno")]'

    def get_policy_status_xpath(self, is_package = False):
        if is_package:
            return '//form[@id="PackagePolicyInfo"]//div[contains(@class, "PolicyInfo-PolicyStatus")]/div[contains(@class, "statusText")]'
        return '//form[@name="PolicyInfo"]//div[contains(@class, "PolicyInfo-PolicyStatus")]/div[contains(@class, "statusText")]'

    def get_policy_agent_input_xpath(self):
        return '//select[@name="AgentID"]/option'

    def login(self, login_email, login_password):
        WebDriverWait(self.driver, 60).until(lambda d: d.find_element_by_xpath('//input[@type="password"]'))
        time.sleep(3)
        self.driver.find_element_by_xpath('//input[@name="txtEmail"]').send_keys(login_email)
        self.driver.find_element_by_xpath('//input[@type="password"]').send_keys(login_password)
        time.sleep(3)
        self.driver.find_element_by_xpath('//a[@id="lnkSubmit"]').click()
        time.sleep(10)
        if len(self.driver.find_elements_by_xpath('//a[@id="lnkCancel"]')) > 0:
            self.driver.find_element_by_xpath('//a[@id="lnkCancel"]').click()
            time.sleep(10)

    def scroll(self, export_button):
        desired_y = (export_button.size['height'] / 2) + export_button.location['y']
        window_h = self.driver.execute_script('return window.innerHeight')
        window_y = self.driver.execute_script('return window.pageYOffset')
        current_y = (window_h / 2) + window_y
        scroll_y_by = desired_y - current_y
        self.driver.execute_script("window.scrollBy(0, arguments[0]);", scroll_y_by)

    def search_customer_by_full_name(self, cust_full_name):
        search_box_xpath = self.get_search_box_xpath()
        if self.check_exists_by_xpath(search_box_xpath):
            element_search_box = self.driver.find_element_by_xpath(search_box_xpath)
            element_search_box.clear()
            time.sleep(2)
            element_search_box.send_keys(cust_full_name)
            time.sleep(5)
        else:
            print('No Search Box Found')

    def access_cust_details(self):
        search_results_xpath = self.get_search_result_items_xpath()
        no_result_xpath = self.get_no_results_xpath()
        if_no_results = False
        if self.check_exists_by_xpath(no_result_xpath):
            no_result_block = self.driver.find_element_by_xpath(no_result_xpath)
            attributes = no_result_block.get_attribute('style')
            if_no_results = attributes == 'display: block;'

        if self.check_exists_by_xpath(search_results_xpath) and not if_no_results:
            element_search_result = self.driver.find_elements_by_xpath(search_results_xpath)[0]
            try:
                element_search_result.click()
            except WebDriverException:
                print('Not Clickable')
                return False
            time.sleep(5)
            return True
        else:
            print('No Search Items Found')
            return False

    def access_policy_tab(self):
        policy_tab_id = 'tab-Policies'
        button_policy_tab = self.driver.find_element_by_id(policy_tab_id)
        button_policy_tab.click()
        return self.driver.current_url

    def access_policies(self, policy_tab_url):
        policy_items_xpath = self.get_policies_from_list_xpath()
        policy_elements = self.driver.find_elements_by_xpath(policy_items_xpath)
        if policy_elements:
            return policy_elements
        else:
            print('No policy with this customer')
            return

    def access_single_policy_details(self, policy_element):
        self.scroll(policy_element)
        policy_element.click()

        time.sleep(3)
        is_package = self.check_exists_by_id('PackagePolicyInfo')

        policy_number = self.driver.find_element_by_xpath(self.get_policy_number_xpath()).text
        policy_premium = self.driver.find_element_by_xpath(self.get_policy_premium_xpath()).text
        policy_status = self.driver.find_element_by_xpath(self.get_policy_status_xpath(is_package)).text
        policy_ori_agent = self.get_original_agent()

        return policy_number, policy_premium, policy_status, policy_ori_agent, is_package

    def check_customer_numer(self, cust_number_in_dataset):
        cust_number_xpath = self.get_cust_number_xpath()
        cust_num_on_page = self.driver.find_element_by_xpath(cust_number_xpath).text
        return cust_num_on_page == cust_number_in_dataset

    def get_original_agent(self):
        original_agent_xpath = self.get_original_agent_xpath()
        if self.check_exists_by_xpath(original_agent_xpath):
            ori_agent = self.driver.find_element_by_xpath(original_agent_xpath).text
            print('Original Agent: ' + ori_agent)
            return ori_agent
        print('No Original Agent Found')
        return ''

    def get_original_csr(self):
        original_csr_xpath = self.get_original_csr_xpath()
        if self.check_exists_by_xpath(original_csr_xpath):
            ori_csr = self.driver.find_element_by_xpath(original_csr_xpath).text
            return ori_csr
        print('No Original CSR Found')
        return ''

    def start_to_edit(self, is_package = False):
        button_edit_xpath = self.get_button_edit_xpath(is_package)
        if self.check_exists_by_xpath(button_edit_xpath):
            button_edit = self.driver.find_element_by_xpath(button_edit_xpath)
            self.scroll(button_edit)
            time.sleep(2)
            if self.check_clickable_by_xpath(button_edit_xpath):
                #button_edit.click()
                time.sleep(2)
                return True
            else:
                print('Edit Button Not Clickable')
        else:
            print('No Edit Button Found')

        return False

    def mark_agent(self, value="12816"):
        agent_id_element = self.driver.find_element_by_name('AgentID')
        self.scroll(agent_id_element)
        dropdown_agent = Select(agent_id_element)
        dropdown_agent.select_by_value(value)
        time.sleep(5)

    def mark_csr(self, value="12816"):
        dropdown_csr = Select(self.driver.find_element_by_id('account-info-csr'))
        dropdown_csr.select_by_value(value)
        time.sleep(5)

    def save_info(self, is_package = False):
        button_save_xpath = self.get_button_save_xpath(is_package)
        if self.check_exists_by_xpath(button_save_xpath):
            button_save = self.driver.find_element_by_xpath(button_save_xpath)
            self.scroll(button_save)
            if not self.check_clickable_by_xpath(button_save_xpath):
                print('Save Button Not Clickable')
        time.sleep(5)


def read_csv_data(csv_path):
    rows = []
    with open(csv_path, 'r') as file:
        csvreader = csv.DictReader(file)
        for row in csvreader:
            rows.append(row)

    return rows


def main():
    param = Param()
    '''
    FireFox Driver
    '''
    driver = webdriver.Firefox(executable_path=param.driver_path)
    print('Invoke Driver Successfully')

    driver.get(param.qq_index)
    driver.maximize_window()

    '''
    Login
    '''
    selenium_web = SeleniumWeb(driver)
    selenium_web.login(param.login_email, param.login_password)
    print('Login Successfully')
    time.sleep(3)

    '''
    Read QQ Data
    '''
    customer_rows = read_csv_data(param.ori_data_path)
    print('Get {number} pieces of data'.format(number=str(len(customer_rows))))

    '''
    Open Record Files
    '''
    pp_policy_num_f = open(param.processed_policy_number_path, 'ab')
    pp_policy_records_f = open(param.processed_policy_records_path, 'a')
    '''
    Mark Account Info As Transfer Account
    '''
    # Customer Number
    processed_cust_num_file = open(param.processed_customer_number_path, 'rb')
    cust_num_list = []
    while True:
        try:
            cust_num_list.append(pickle.load(processed_cust_num_file))
        except EOFError:
            break
    processed_cust_num_file.close()
    processed_cust_num_set = set(cust_num_list[1:])
    print('Processed Customer: {number} '.format(number=str(len(processed_cust_num_set))))
    # Policy Number
    processed_policy_num_file = open(param.processed_customer_number_path, 'rb')
    policy_num_list = []
    # while True:
    #     try:
    #         policy_num_list.append(pickle.load(processed_policy_num_file))
    #     except EOFError:
    #         break
    # processed_policy_num_file.close()
    processed_policy_num_set = set(policy_num_list[1:])
    print('Processed Policy: {number} '.format(number=str(len(processed_policy_num_set))))
    # with open('processed_cust_num_set.txt', 'rb') as f:
    #     processed_cust_num_set = set(pickle.load(f))
    # print(processed_cust_num_set)
    for index, cust in enumerate(customer_rows):
        cust_first_name, cust_last_name = cust['Customer First Name'], cust['Customer Last Name']
        cust_num = cust['Customer Number']
        cust_full_name = cust_first_name + ' ' + cust_last_name

        print(f'--- No.{index} {cust_full_name}  {cust_num}')
        if cust_num in processed_cust_num_set:
            print('exist!')
            continue

        selenium_web.search_customer_by_full_name(cust_num)
        if_search_results_exist = selenium_web.access_cust_details()
        if if_search_results_exist:
            customer_rows[index]['If Exist'] = 1
        else:
            customer_rows[index]['If Exist'] = 0
            processed_cust_num_set.add(cust_num)
            driver.get(param.qq_index)
            time.sleep(5)
            with open(processed_cust_num_file, 'ab') as f:
                pickle.dump(cust_num, f)
            continue

        if selenium_web.check_exists_by_id('AccountInfo'):
            element_account_info = selenium_web.driver.find_element_by_id('AccountInfo')
            selenium_web.scroll(element_account_info)
            if_customer_number_match = selenium_web.check_customer_numer(cust_num)
            if not if_customer_number_match:
                break
            # Go to the policy tab
            policy_tab_url = selenium_web.access_policy_tab()
            time.sleep(5)
            policy_number = len(selenium_web.access_policies(policy_tab_url))
            if policy_number:
                for index in range(policy_number):
                    policy_items = selenium_web.driver.find_elements_by_xpath(selenium_web.get_policies_from_list_xpath())
                    policy = policy_items[index]
                    selenium_web.scroll(policy)
                    plc_num_from_preview = policy.text
                    print('  {index}. {plc_num_from_preview}'.format(index=index, plc_num_from_preview=plc_num_from_preview))
                    # if plc_num_from_preview in processed_policy_num_set:
                    #     continue

                    plc_num, premium, status, ori_agent, is_package= selenium_web.access_single_policy_details(policy)
                    if not is_package:
                        if_clickable = selenium_web.start_to_edit()
                        if if_clickable :
                            selenium_web.mark_agent()
                            selenium_web.save_info()
                    else:
                        if_clickable = selenium_web.start_to_edit(is_package)
                        if if_clickable:
                            selenium_web.mark_agent()
                            selenium_web.save_info(is_package)

                    # record the policy info
                    policy_record = {'Customer First Name': cust_first_name, 'Customer Last Name': cust_last_name, \
                                     'Customer Number': cust_num, 'Policy Num': plc_num, \
                                     'Policy Premium': premium, 'Policy Status': status, \
                                     'Agent On Policy': ori_agent}

                    processed_policy_num_set.add(plc_num)

                    with open(param.processed_policy_number_path, 'ab') as pp_policy_num_f:
                        pickle.dump(plc_num, pp_policy_num_f)

                    with open(param.processed_policy_records_path, 'a') as pp_policy_records_f:
                        policy_record_writer = csv.DictWriter(pp_policy_records_f, fieldnames = policy_record.keys())
                        policy_record_writer.writerow(policy_record)

                    selenium_web.driver.get(policy_tab_url)
                    time.sleep(5)



            # # Get the Original Agent & CSR
            # original_agent = selenium_web.get_original_agent()
            # original_csr = selenium_web.get_original_csr()
            # customer_rows[index]['Original Agent'] = original_agent
            # customer_rows[index]['Original CSR'] = original_csr
            # # Edit
            # selenium_web.start_to_edit()
            # selenium_web.mark_agent()
            # #selenium_web.mark_csr()
            # selenium_web.save_info()

        processed_cust_num_set.add(cust_num)

        with open(param.processed_customer_number_path, 'ab') as f:
            pickle.dump(cust_num, f)

        with open(param.processed_customer_records_path, 'a', newline='') as csvfile:
            fieldnames = list(cust.keys())
            writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
            writer.writeheader()
            writer.writerow(customer_rows[index])

        driver.get(param.qq_index)
        time.sleep(5)
        print(index)
#        if index == 3:
#            raise Exception("Hit 3")



def mark():

    try:
        main()
    except Exception as e:
        print(e)

        send_email_to('pye3@wpi.edu')

def send_email_to(toEmail):
    gmail_user = 'berniey@eigenanalytics.com'
    gmail_password = 'Sb^q3dWs6ZMm6ME'
    sent_from = gmail_user
    to = toEmail
    subject = 'OMG Super Important Message'
    body = 'Bug Appeared'

    email_text = """\
           From: %s
           Subject: %s

           %s
           """ % (sent_from, subject, body)
    server = smtplib.SMTP_SSL('smtp.gmail.com', 465)
    server.ehlo()
    server.login(gmail_user, gmail_password)
    server.sendmail(sent_from, to, email_text)
    server.close()

    print('Email sent!')


if __name__ == '__main__':
    mark()