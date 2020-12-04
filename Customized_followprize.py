from selenium.common.exceptions import WebDriverException, NoSuchElementException
from selenium import webdriver
import unittest
from time import sleep
import logging
from datetime import datetime
import ddt
import pandas as pd
import os
import numpy as np
from selenium.webdriver import ActionChains

ROOT_PATH = os.getcwd()
FILE_PATH = os.path.join(os.getcwd(), "Follow_prize_TW.xlsx")

class ShopFollowPrize():

    def __init__(self):
        # create logger
        logger_name_error = "error"
        logger_name_success = "success"
        self.logger_error = logging.getLogger(logger_name_error)
        self.logger_error.setLevel(logging.ERROR)
        self.logger_success = logging.getLogger(logger_name_success)
        self.logger_success.setLevel(logging.CRITICAL)

        # create file handler
        readable_time = datetime.now().strftime('%Y-%m-%d %H')
        log_path_err = f"./Follow Prize_errors occurring at around {readable_time} o'clock.txt"
        fh_err = logging.FileHandler(log_path_err)
        fh_err.setLevel(logging.ERROR)
        log_path_suc = f"./Follow Prize_success shops at around {readable_time} o'clock.txt"
        fh_suc = logging.FileHandler(log_path_suc)
        fh_suc.setLevel(logging.CRITICAL)

        # create formatter
        fmt_err = "%(asctime)s %(message)s"
        fmt_suc = "%(asctime)s %(message)s"
        datefmt = "%a %d %b %Y %H:%M:%S"
        formatter_err = logging.Formatter(fmt_err, datefmt)
        formatter_suc = logging.Formatter(fmt_suc, datefmt)
        # add handler and formatter to logger
        fh_err.setFormatter(formatter_err)
        fh_suc.setFormatter(formatter_suc)
        self.logger_error.addHandler(fh_err)
        self.logger_success.addHandler(fh_suc)

        #create tracker
        self.failing_tracker = pd.DataFrame()
        self.successful_tracker = pd.DataFrame()

        try:
            # driver_path = C:/Users/ellen.huang/Downloads/Shop Management Project/Automation/decoration/chromedriver.exe
            options = webdriver.ChromeOptions()
            options.add_argument('lang=zh_CN.UTF-8')
            self.driver = webdriver.Chrome(
                executable_path="chromedriver.exe", chrome_options=options)
            self.logger_error.debug("Successfully loaded the driver")
        except WebDriverException as e:
            print("\n-----------------------------WebDriverException-------------------------------------------------\n"
                  , e)
            self.logger_error.error(e, msg="无法连接到Chrome driver")
            print("\n--------------------------------------------------------------------------------------------\n")
        self.driver.maximize_window()
        # 切换到无头模式远程浏览器
        # seleniu_grid_url = "http://192.168.33.128:4444/wd/hub"
        # capabilities = webdriver.DesiredCapabilities.CHROME.copy()
        # driver_remote = webdriver.Remote(command_executor=seleniu_grid_url, desired_capabilities=capabilities)
        # driver_remote.maximize_window()
        self.driver.implicitly_wait(6)


    def handler(self):
        # logs = CustomLogger('Failed Shop Record')
        # filename = 'logs.csv'
        # logs = []
        xls = pd.ExcelFile(FILE_PATH)
        df_shop_info = pd.read_excel(xls, 'raw')
        xls1 = pd.ExcelFile('BR_SIP_Shop_passwords.xlsx')
        df_password = pd.read_excel(xls1, 'Sheet1')
        shopid_list = df_shop_info['affi_shopid'].drop_duplicates().tolist()
        iterName = iter(shopid_list)
        # package = read_data()
        # for shop in package:
        #     print(shop)
        while True:
            try:
                element = next(iterName)
                username = df_shop_info.loc[df_shop_info['affi_shopid'] == element]['username'].drop_duplicates().to_string(
                    index=None).strip()
                follow_prize_name = df_shop_info.loc[df_shop_info['affi_shopid'] == element]['campaign_name'].drop_duplicates().to_list()[0]
                Discount_percentage = df_shop_info.loc[df_shop_info['affi_shopid'] == element]['Discount_percentage'].drop_duplicates().to_list()[0]
                loc = username.split('.')[-1]
                Fixed_amount = str(df_shop_info.loc[df_shop_info['affi_shopid'] == element]['Fixed_amount'].drop_duplicates().to_list()[0])
                Min_spend = str(df_shop_info.loc[df_shop_info['affi_shopid'] == element]['Min_Spend'].drop_duplicates().to_list()[0])
                Final_cap = str(df_shop_info.loc[df_shop_info['affi_shopid'] == element]['Final cap'].drop_duplicates().to_list()[0])
                Month_limit = str(df_shop_info.loc[df_shop_info['affi_shopid'] == element]['Month limit'].drop_duplicates().to_list()[0])
                start_date = df_shop_info.loc[df_shop_info['affi_shopid'] == element]['start_time'].drop_duplicates().to_list()[0]
                end_date = df_shop_info.loc[df_shop_info['affi_shopid'] == element]['end_time'].drop_duplicates().to_list()[0]
                start_day = str(start_date).split(' ')[0].split('-')[2]
                end_day = str(end_date).split(' ')[0].split('-')[2]
                start_month = str(start_date).split(' ')[0].split('-')[1]
                end_month = str(end_date).split(' ')[0].split('-')[1]

                # get password
                if username in df_password['username'].tolist():
                    password = df_password.loc[df_password.username == username]['password'].tolist()[0]
                else:
                    password = 'Scb88@Sip99'

                # 获取webdriver
                browser = self.driver
                logger_error = self.logger_error
                logger_success = self.logger_success

                # 登录流程开始
                browser.delete_all_cookies()
                if loc =='tw':
                    browser.get("https://seller.xiapi.shopee.cn/account/signin")
                else:
                    browser.get("https://seller.{}.shopee.cn/account/signin".format(loc))
                sleep(4)
                try:
                    browser.switch_to.alert.accept()
                except Exception as e:
                    print(e)
                    pass

                # 用户名
                try:
                    browser.find_elements_by_xpath(
                        '//label[@for="username"]/following-sibling::div//input')[0].send_keys(username)
                    sleep(2)
                    # 密码
                    browser.find_elements_by_xpath(
                        '//label[@for="password"]/following-sibling::div//input')[0].send_keys(password)
                    sleep(2)
                    # 登录按钮
                    browser.find_elements_by_xpath(
                        '//div[@class="shopee-form-item"]//button')[0].click()
                    sleep(3)
                    if loc == 'tw':
                        assert(browser.current_url == f'https://seller.xiapi.shopee.cn/')
                    else:
                        assert(browser.current_url == f'https://seller.{loc}.shopee.cn/')
                    sleep(3)
                except Exception as e:
                    logger_error.error(f'{username}')
                    # logerror = {'username' : f'{username}', 'message' : f'{e}'}
                    # logerror.append(logs)
                    browser.refresh()
                    sleep(3)
                    continue

                # 关闭卖家条款
                try:
                    browser.find_elements_by_xpath('//div[@class="shopee-modal__body"]//div[@class="checkbbox"]//span')[0].click()
                    sleep(1)
                    browser.find_elements_by_xpath('//div[@class="shopee-modal__body"]//div[@class="footer"]//button')[0].click()
                    sleep(1)
                except Exception as e:
                    print("No seller agreement")
                    pass

                # 需要先把语言改成母语
                try:
                    if loc == 'tw':
                        browser.get('https://seller.xiapi.shopee.cn/portal/settings/basic/shop')
                    else:
                        browser.get(f'https://seller.{loc}.shopee.cn/portal/settings/basic/shop')
                        sleep(2)
                    try:
                        browser.find_elements_by_xpath('//div[@class="onboarding-tips-buttons"]//button')[1].click()
                        sleep(2)
                    except Exception as e:
                        pass
                    browser.find_elements_by_xpath('//label//span[@class="shopee-radio__indicator"]')[0].click()
                    sleep(2)
                except Exception as e:
                    print(e)
                    logger_error.error(f'{username}redirect')
                    msg = {'shop_id': [username], 'msg': [f'shop {username} has failed to redirect'], 'detail': [e]}
                    to_update = pd.DataFrame(msg)
                    self.failing_tracker = self.failing_tracker.append(to_update)


                # 新建FollowPrize流程开始
                try:
                    if loc == 'tw':
                        browser.get("https://seller.xiapi.shopee.cn/portal/marketing/follow-prize/create")
                    else:
                        browser.get("https://seller.{}.shopee.cn/portal/marketing/follow-prize/create".format(loc))
                    sleep(4)
                    if loc == 'tw':
                        assert (browser.current_url == f'https://seller.xiapi.shopee.cn/portal/marketing/follow-prize/create')
                    else:
                        assert (browser.current_url == "https://seller.{}.shopee.cn/portal/marketing/follow-prize/create".format(loc))
                    # self.assertEqual("https://seller.{}.shopee.cn/portal/marketing/follow-prize/create".format(loc),
                    #                  browser.current_url)
                except Exception as e:
                    print(e)
                    logger_error.error(f'{username}redirect')
                    msg = {'shop_id': [username], 'msg': [f'shop {username} has failed to redirect'], 'detail': [e]}
                    to_update = pd.DataFrame(msg)
                    self.failing_tracker = self.failing_tracker.append(to_update)
                    continue

                # 输入Follow Prize Name
                try:
                    browser.find_elements_by_xpath(f'//div[@class="edit_info"]//input[@class ="shopee-input__input"]')[0].send_keys(
                        f'{follow_prize_name}')
                    # browser.find_elements_by_xpath('/html/body/div[1]/div[2]/div/div/div/div[1]/div[2]/form/div[1]/div[1]/div/div[1]/div/input')[0].send_keys(f'{follow_prize_name[n]}')
                    sleep(2)
                    if start_month == end_month:
                        # 点击结束日期
                        browser.find_elements_by_css_selector(
                            '.shopee-date-picker__input')[1].click()
                        sleep(2)
                        browser.find_elements_by_xpath(
                            f'//div[@class="shopee-popper shopee-date-picker__picker"]//div[contains(@class,"shopee-date-table__cell normal")]//span[text()="{end_day}"]')[
                            1].click()
                        sleep(2)
                        browser.find_elements_by_xpath('/html/body/div[@class="shopee-popper shopee-date-picker__picker"]//button')[0].click()
                        sleep(2)
                    # 如果翻页
                    else:
                        # 点击结束日期
                        browser.find_elements_by_css_selector(
                            '.shopee-date-picker__input')[1].click()
                        sleep(2)
                        # 需要切换到下一个月再选择日期
                        browser.find_elements_by_xpath(
                            '//i[@class="shopee-icon shopee-picker-header__icon shopee-picker-header__next"]')[
                            2].click()
                        sleep(2)
                        browser.find_elements_by_xpath(
                            f'//div[@class="shopee-popper shopee-date-picker__picker"]//div[contains(@class,"shopee-date-table__cell normal")]//span[text()="{end_day}"]')[
                            1].click()
                        sleep(2)
                        browser.find_elements_by_xpath(
                            '/html/body/div[@class="shopee-popper shopee-date-picker__picker"]//button')[0].click()
                        sleep(2)
                except Exception as e:
                    print(e)
                    logger_error.error(f'{username} error_select_date')
                    msg = {'shop_id': [username], 'msg': [f'shop {username} has failed to update date'], 'detail': [e]}
                    to_update = pd.DataFrame(msg)
                    self.failing_tracker = self.failing_tracker.append(to_update)
                    continue

                try:
                    if Discount_percentage:
                        #点开drop down list
                        browser.find_elements_by_xpath('//div[@class="shopee-selector shopee-selector--normal"]')[2].click()
                        sleep(1)
                        browser.find_elements_by_xpath('//div[@class="shopee-scrollbar"]//div[@class="shopee-option"]')[0].click()
                        sleep(1)
                        #输入discount amount
                        # browser.find_elements_by_xpath(
                        #     '//div[@class="shopee-form-item__control"]//input')[
                        #     6].send_keys(f'{discount_amount[n]}')
                        # sleep(0.5)
                        # 输入Discount Amount
                        browser.find_elements_by_xpath(
                            '//div[@class="shopee-input discount-input"]//input')[
                            0].send_keys(f'{Discount_percentage}')
                        sleep(1)
                        #输入maximum Discount Price
                        browser.find_elements_by_xpath('//div[@class="shopee-input currency-input"]//input')[1].send_keys(f'{Final_cap}')
                        sleep(1)
                        #输入Minimum Basket Price
                        browser.find_elements_by_xpath('//div[@class="shopee-input currency-input"]//input')[2].send_keys(f'{Min_spend}')
                        sleep(1)
                        #输入Follow Prize Quantity
                        browser.find_elements_by_xpath('//div[@class="shopee-form-item"]//input')[-1].send_keys(f'{Month_limit}')
                        sleep(2)

                    else:
                        # 输入Discount Amount
                        browser.find_elements_by_xpath(
                            '//div[@class="shopee-input currency-input"]//input')[
                            0].send_keys(f'{Fixed_amount}')
                        sleep(1)
                        # 输入Minimum Basket Price
                        browser.find_elements_by_xpath(
                            '//div[@class="shopee-input currency-input"]//input')[
                            1].send_keys(f'{Min_spend}')
                        sleep(1)
                        # 输入Follow Prize Quantity
                        browser.find_elements_by_xpath(
                            '//div[@class="shopee-form-item"]//input')[
                            -1].send_keys(f'{Month_limit}')
                        sleep(2)
                    # 滑到元素并点击
                    current_element = browser.find_elements_by_xpath('//div[@class="shopee-fix-bottom-card bottom-card"]//button')[1]
                    ActionChains(browser).move_to_element(current_element).perform()
                    sleep(2)
                    current_element.click()
                    sleep(2)

                    msg = {'shop_id': [username], 'msg': [f'shop {username} has created follow prize successfully']}
                    to_update = pd.DataFrame(msg)
                    self.successful_tracker = self.successful_tracker.append(to_update)
                    # logerror = {'username': f'{username}', 'message': 'failed to select date'}
                    # logerror.append(logs)
                except Exception as e:
                    msg = {'shop_id': [username], 'msg': [f'shop {username} has failed to create follow prize'], 'detail': [e]}
                    to_update = pd.DataFrame(msg)
                    self.failing_tracker = self.failing_tracker.append(to_update)
                    continue

            except StopIteration:
                t = datetime.now().strftime('%Y-%m-%d_%H%M')
                self.failing_tracker.to_csv(ROOT_PATH + '/' + t + '_failed_followprize_shops_log.csv', encoding='utf-8-sig')
                self.successful_tracker.to_csv(ROOT_PATH + '/' + t + '_success_followprize_shops_log.csv', encoding='utf-8-sig')
                self.driver.quit()
                self.driver.close()

if __name__ == '__main__':
    agent = ShopFollowPrize()
    agent.handler()
    verbosity = 2



