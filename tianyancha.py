from my.MySeleniumCralwer import *
import timeit
import xlrd
import csv

Group = 350  # excel的组号
InputPath = r"./" + str(Group) + "-2021中国环保行业企业名录.xlsx"
Main_Staff_Path = r"./" + str(Group) + "-2021中国环保行业企业名录主要人员列表.csv"
Failure_List_Path = r"./" + str(Group) + "-2021中国环保行业企业名录抓取失败列表.csv"
Shareholder_Path = r"./" + str(Group) + "-2021中国环保行业企业名录股东信息列表.csv"
Change_Log_Path = r"./" + str(Group) + "-2021中国环保行业企业名录变更记录列表.csv"
Supplier_Path = r"./" + str(Group) + "-2021中国环保行业企业名录供应商列表.csv"
Customer_Path = r"./" + str(Group) + "-2021中国环保行业企业名录客户列表.csv"

# 环保企业名录路径
workbook = xlrd.open_workbook(InputPath, 'rb')
sheet1_object = workbook.sheet_by_index(0)
# 打开csv文件并设置为续写模式
Main_Staff = open(Main_Staff_Path, 'a', encoding='utf-8-sig', newline='')
csv_writer1 = csv.writer(Main_Staff)
# csv_writer1.writerow(['企业名称', '统一社会信用代码', 'url', '姓名', '职位', '持股比例', '最终受益股份'])

Failure_List = open(Failure_List_Path, 'a', encoding='utf-8-sig', newline='')
csv_writer2 = csv.writer(Failure_List)
# csv_writer2.writerow(['企业名称', '统一社会信用代码', 'url'])

Shareholder = open(Shareholder_Path, 'a', encoding='utf-8-sig', newline='')
csv_writer3 = csv.writer(Shareholder)
# csv_writer3.writerow(['企业名称', '统一社会信用代码', 'url', '股东（发起人）', '持股比例', '最终收益股份', '认缴出资额', '认缴出资日期'])

Change_Log = open(Change_Log_Path, 'a', encoding='utf-8-sig', newline='')
csv_writer4 = csv.writer(Change_Log)
# csv_writer4.writerow(['企业名称', '统一社会信用代码', 'url', '变更日期', '变更项目', '变更前', '变更后'])

Supplier = open(Supplier_Path, 'a', encoding='utf-8-sig', newline='')
csv_writer5 = csv.writer(Supplier)
# csv_writer5.writerow(['企业名称', '统一社会信用代码', 'url', '供应商', '采购占比', '采购金额', '报告期' ,'数据来源', '关联关系'])

Customer = open(Customer_Path, 'a', encoding='utf-8-sig', newline='')
csv_writer6 = csv.writer(Customer)
# csv_writer6.writerow(['企业名称', '统一社会信用代码', 'url', '客户', '销售占比', '销售金额', '报告期' ,'数据来源', '关联关系'])


# 新建爬虫
tianyancha = MySeleniumCralwer('tianyancha', 'test', 'https://www.tianyancha.com/?jsid=SEM-BAIDU-PZ-SY-2021112-BIAOTI')
tianyancha.driver.maximize_window()

# 移除广告
# tianyancha.Click('//div[@class="modal-content"]/div[1]')
# 登录界面
tianyancha.Click('J_NavTypeLink', type='ID')
tianyancha.Click('//div[@class="toggle_box -qrcode"]')
tianyancha.Click('//div[@class="sign-in"]/div[1]/div[2]')
# 登录输入
tianyancha.Login('mobile', 'XXXXXXXXX', 'password', 'XXXXXXXXXXXX', type='ID')
sleep(1)
tianyancha.Click('//div[@class="sign-in"]/div[2]/div[2]')

# 等待人工滑动验证码
print('请在10s内手动通过验证......')
sleep(10)

# 进入任意一个搜索页面（不能是网站主页）
tianyancha.Get('https://www.tianyancha.com/search?key=913408816808445565')

# 逐行读取企业名录
StartNum = 0
for i in range(StartNum, 300):
    # 判断当前页面是否为机器人验证
    if tianyancha.driver.current_url.find('antirobot') != -1:
        tianyancha.Alarm()
        input()

    # 开始计时
    start = timeit.default_timer()
    # 获取企业名称及统一社会信用代码
    all_row_values = sheet1_object.row_values(rowx=i + 1)
    company = all_row_values[1]
    nums = all_row_values[15]
    # 判断统一社会信用代码是否为空
    if len(nums) == 0:
        print(tianyancha.driver.current_url, "该企业没有统一社会信用代码")
        csv_writer2.writerow([company, nums, tianyancha.driver.current_url, "该企业没有统一社会信用代码"])
        continue
    print("正在爬取第" + str(i) + "个企业")
    # print("企业名称：", company)
    # print("统一社会信用代码:", nums)

    MaxWaitTime = 10

    try:
        # 清除内容并通过统一社会信用代码搜索
        tianyancha.FindElement("//*[@id='header-company-search']").clear()
        tianyancha.Input("//*[@id='header-company-search']", content=nums)
        # 定位搜索按钮并点击
        if i == StartNum:
            sleep(1)
        tianyancha.Click("//body/div[1]/div[1]/div[2]/div[1]/div[2]/div[1]/div[1]")
        # 通过企业名称定位标签并点击
        tianyancha.Click("//a[contains(text(),'" + company + "')]")
        # 切换到最近打开的标签页，并打印当前url
        tianyancha.driver.switch_to.window(tianyancha.driver.window_handles[-1])
        # print('当前浏览地址为：{0}'.format(tianyancha.driver.current_url))

        # 获取主要人员信息
        try:
            tianyancha.FindElement("//span[contains(text(),'主要人员')]")
        except Exception as e:
            print(tianyancha.driver.current_url, "该企业页面没有主要人员信息")
            csv_writer2.writerow([company, nums, tianyancha.driver.current_url, "该企业页面没有主要人员信息"])
        else:
            try:
                personNum = 0
                persons = tianyancha.FindElements("#_container_staffCount > div > table > tbody > tr", type='CSS')
                for person in persons:
                    personNum = personNum + 1
                for j in range(personNum):
                    # 姓名
                    name = tianyancha.FindElement('#_container_staffCount > div > table > tbody > tr:nth-child(' + str(
                        j + 1) + ') > td:nth-child(2) > table > tbody > tr > td:nth-child(2) > a.link-click', type='CSS')
                    # 职位
                    title = tianyancha.FindElement('#_container_staffCount > div > table > tbody > tr:nth-child(' + str(
                        j + 1) + ') > td:nth-child(3) > span', type='CSS')
                    # 持股比例
                    percentage = tianyancha.FindElement('#_container_staffCount > div > table > tbody > tr:nth-child(' + str(
                        j + 1) + ') > td:nth-child(4)', type='CSS')
                    # 最终收益股份
                    final_percentage = tianyancha.FindElement('#_container_staffCount > div > table > tbody > tr:nth-child(' + str(
                        j + 1) + ') > td:nth-child(5)', type='CSS')
                    name = name.text
                    title = title.text
                    percentage = percentage.text
                    final_percentage = final_percentage.text.strip('"股权链\n"')
                    print(company, nums, tianyancha.driver.current_url, name, title, percentage, final_percentage)
                    csv_writer1.writerow([company, nums, tianyancha.driver.current_url, name, title, percentage, final_percentage])    
            except Exception as e:
                print(tianyancha.driver.current_url, "获取主要人员信息失败")
                csv_writer2.writerow([company, nums, tianyancha.driver.current_url, "获取主要人员信息失败"])

        # 获取股东信息
        try:
            tianyancha.FindElement("//span[contains(text(),'股东信息')]")
        except Exception as e:
            print(tianyancha.driver.current_url, "该企业页面没有股东信息")
            csv_writer2.writerow([company, nums, tianyancha.driver.current_url, "该企业页面没有股东信息"])
        else:
            try:
                personNum = 0
                persons = tianyancha.FindElements("#_container_holderCount > div > table > tbody > tr", type='CSS')
                for person in persons:
                    personNum = personNum + 1
                for j in range(personNum):
                    # 股东（发起人）
                    try:
                        name = tianyancha.FindElement('#_container_holderCount > div > table > tbody > tr:nth-child(' + str(
                            j + 1) + ') > td:nth-child(2) > table > tbody > tr > td:nth-child(2) > a.link-click', type='CSS')
                    except Exception as e:
                        name = tianyancha.FindElement('#_container_holderCount > div > table > tbody > tr:nth-child(' + str(
                            j + 1) + ') > td:nth-child(2) > table > tbody > tr > td:nth-child(2) > span', type='CSS')
                    # 持股比例
                    percentage1 = tianyancha.FindElement('#_container_holderCount > div > table > tbody > tr:nth-child(' + str(
                        j + 1) + ') > td:nth-child(3) > div > div > span', type='CSS')
                    # 最终收益股份
                    percentage2 = tianyancha.FindElement('#_container_holderCount > div > table > tbody > tr:nth-child(' + str(
                        j + 1) + ') > td:nth-child(4) > span:nth-child(1)', type='CSS')
                    # 认缴出资额
                    money = tianyancha.FindElement('#_container_holderCount > div > table > tbody > tr:nth-child(' + str(
                        j + 1) + ') > td:nth-child(5) > div > span', type='CSS')
                    # 认缴出资日期
                    date = tianyancha.FindElement('#_container_holderCount > div > table > tbody > tr:nth-child(' + str(
                        j + 1) + ') > td:nth-child(6) > div > span', type='CSS')
                    name = name.text
                    percentage1 = percentage1.text
                    percentage2 = percentage2.text.strip('"股权链\n"')
                    money = money.text
                    date = date.text
                    print(company, nums, tianyancha.driver.current_url, name, percentage1, percentage2, money, date)
                    csv_writer3.writerow([company, nums, tianyancha.driver.current_url, name, percentage1, percentage2, money, date])
            except Exception as e:
                print(tianyancha.driver.current_url, "获取股东信息失败")
                csv_writer2.writerow([company, nums, tianyancha.driver.current_url, "获取主要人员信息失败"])

        # 获取变更记录
        try:
            tianyancha.FindElement("//span[contains(text(),'变更记录')]")
        except Exception as e:
            print(tianyancha.driver.current_url, "该企业页面没有变更记录")
            csv_writer2.writerow([company, nums, tianyancha.driver.current_url, "该企业页面没有变更记录"])
        else:
            try:
                total = tianyancha.FindElement("#nav-main-changeCount > span.data-count", type='CSS')
                total = int(total.text)
                pages = int(total / 10)
                if total % 10 != 0:
                    pages = pages + 1
                page = 1
                while page <= pages:
                    tmp = tianyancha.FindElement('#_container_changeinfo > div > table > tbody > tr:nth-child(1) > td:nth-child(1)'
                        , type='CSS')
                    tmp = int(tmp.text)
                    if tmp != (page - 1) * 10 + 1:
                        sleep(0.5)
                        continue
                    
                    personNum = 0
                    persons = tianyancha.FindElements("#_container_changeinfo > div > table > tbody > tr", type='CSS')
                    for person in persons:
                        personNum = personNum + 1
                    for j in range(personNum):
                        # 变更日期
                        date = tianyancha.FindElement('#_container_changeinfo > div > table > tbody > tr:nth-child(' + str(
                            j + 1) + ') > td:nth-child(2)', type='CSS')
                        # 变更项目
                        item = tianyancha.FindElement('#_container_changeinfo > div > table > tbody > tr:nth-child(' + str(
                            j + 1) + ') > td.left-col', type='CSS')
                        # 变更前
                        before = tianyancha.FindElement('#_container_changeinfo > div > table > tbody > tr:nth-child(' + str(
                            j + 1) + ') > td:nth-child(4)', type='CSS')
                        # 变更后
                        after = tianyancha.FindElement('#_container_changeinfo > div > table > tbody > tr:nth-child(' + str(
                            j + 1) + ') > td:nth-child(5)', type='CSS')
                        date = date.text
                        item = item.text
                        before = before.text
                        after = after.text
                        print(company, nums, tianyancha.driver.current_url, date, item, before, after)
                        csv_writer4.writerow([company, nums, tianyancha.driver.current_url, date, item, before, after])
                    # 下一页
                    if page != pages:
                        if page == 1:
                            tianyancha.Click("#_container_changeinfo > div > div > ul > li:nth-child(" + str(
                                pages + 1) + ") > a", type='CSS')
                        else:
                            tianyancha.Click("#_container_changeinfo > div > div > ul > li:nth-child(" + str(
                                pages + 2) + ") > a", type='CSS')
                    page += 1
            except Exception as e:
                print(tianyancha.driver.current_url, "获取变更记录失败")
                csv_writer2.writerow([company, nums, tianyancha.driver.current_url, "获取变更记录失败"])

        # 获取供应商记录
        try:
            tianyancha.FindElement("//span[contains(text(),'供应商')]")
        except Exception as e:
            print(tianyancha.driver.current_url, "该企业页面没有供应商信息")
            csv_writer2.writerow([company, nums, tianyancha.driver.current_url, "该企业页面没有供应商信息"])
        else:
            try:
                total = tianyancha.FindElement("#nav-main-suppliesV2Count > span.data-count", type='CSS')
                total = int(total.text)
                pages = int(total / 10)
                if total % 10 != 0:
                    pages = pages + 1
                page = 1
                while page <= pages:
                    tmp = tianyancha.FindElement('#_container_supplies > div > table > tbody > tr:nth-child(1) > td:nth-child(1)'
                        , type='CSS')
                    tmp = int(tmp.text)
                    if tmp != (page - 1) * 10 + 1:
                        sleep(0.5)
                        continue
                    
                    personNum = 0
                    persons = tianyancha.FindElements("#_container_supplies > div > table > tbody > tr", type='CSS')
                    for person in persons:
                        personNum = personNum + 1
                    for j in range(personNum):
                        # 供应商
                        try:
                            supplier = tianyancha.FindElement('#_container_supplies > div > table > tbody > tr:nth-child(' + str(
                                j + 1) + ') > td:nth-child(2) > table > tbody > tr > td:nth-child(2) > a.link-click', type='CSS')
                        except Exception as e:
                            supplier = tianyancha.FindElement('#_container_supplies > div > table > tbody > tr:nth-child(' + str(
                                j + 1) + ') > td:nth-child(2) > table > tbody > tr > td:nth-child(2) > div > span', type='CSS')
                        # 采购占比
                        percentage = tianyancha.FindElement('#_container_supplies > div > table > tbody > tr:nth-child(' + str(
                                j + 1) + ') > td:nth-child(3)', type='CSS')
                        # 采购金额
                        price = tianyancha.FindElement('#_container_supplies > div > table > tbody > tr:nth-child(' + str(
                                j + 1) + ') > td:nth-child(4)', type='CSS')
                        # 报告期
                        date = tianyancha.FindElement('#_container_supplies > div > table > tbody > tr:nth-child(' + str(
                                j + 1) + ') > td:nth-child(5)', type='CSS')
                        # 数据来源
                        source = tianyancha.FindElement('#_container_supplies > div > table > tbody > tr:nth-child(' + str(
                                j + 1) + ') > td:nth-child(6)', type='CSS')
                        # 关联关系
                        relation = tianyancha.FindElement('#_container_supplies > div > table > tbody > tr:nth-child(' + str(
                                j + 1) + ') > td:nth-child(7)', type='CSS')
                        supplier = supplier.text
                        percentage = percentage.text
                        price = price.text
                        date = date.text
                        source = source.text
                        relation = relation.text
                        print(company, nums, tianyancha.driver.current_url, supplier, percentage, price, date, source, relation)
                        csv_writer5.writerow([company, nums, tianyancha.driver.current_url, supplier, percentage, price, date, source, relation])
                    # 下一页
                    if page != pages:
                        if page == 1:
                            element = tianyancha.FindElement("#_container_supplies > div > div > ul > li:nth-child(" + str(
                                pages + 1) + ") > a", type="CSS")
                        else:
                            element = tianyancha.FindElement("#_container_supplies > div > div > ul > li:nth-child(" + str(
                                pages + 2) + ") > a", type="CSS")
                        element.click()
                    page += 1
            except Exception as e:
                print(tianyancha.driver.current_url, "获取供应商信息失败")
                csv_writer2.writerow([company, nums, tianyancha.driver.current_url, "获取供应商信息失败"])

        # 获取客户记录
        try:
            tianyancha.FindElement("#nav-main-clientsV2Count > span.data-title", type='CSS')
        except Exception as e:
            print(tianyancha.driver.current_url, "该企业页面没有客户信息")
            csv_writer2.writerow([company, nums, tianyancha.driver.current_url, "该企业页面没有客户信息"])
        else:
            try:
                total = tianyancha.FindElement("#nav-main-clientsV2Count > span.data-count", type='CSS')
                total = int(total.text)
                pages = int(total / 10)
                if total % 10 != 0:
                    pages = pages + 1
                page = 1
                while page <= pages:
                    tmp = tianyancha.FindElement('#_container_clients > div > table > tbody > tr:nth-child(1) > td:nth-child(1)'
                        , type='CSS')
                    tmp = int(tmp.text)
                    if tmp != (page - 1) * 10 + 1:
                        sleep(0.5)
                        continue
                    
                    personNum = 0
                    persons = tianyancha.FindElements("#_container_clients > div > table > tbody > tr", type='CSS')
                    for person in persons:
                        personNum = personNum + 1
                    for j in range(personNum):
                        # 客户
                        try:
                            client = tianyancha.FindElement('#_container_clients > div > table > tbody > tr:nth-child(' + str(
                                j + 1) + ') > td:nth-child(2) > table > tbody > tr > td:nth-child(2) > a.link-click', type='CSS')
                        except Exception as e:
                            client = tianyancha.FindElement('#_container_clients > div > table > tbody > tr:nth-child(' + str(
                                j + 1) + ') > td:nth-child(2) > div > table > tbody > tr > td:nth-child(2) > div > span', type='CSS')
                        # 销售占比
                        percentage = tianyancha.FindElement('#_container_clients > div > table > tbody > tr:nth-child(' + str(
                            j + 1) + ') > td:nth-child(3)', type='CSS')
                        # 销售金额
                        price = tianyancha.FindElement('#_container_clients > div > table > tbody > tr:nth-child(' + str(
                            j + 1) + ') > td:nth-child(4)', type='CSS')
                        # 报告期
                        date = tianyancha.FindElement('#_container_clients > div > table > tbody > tr:nth-child(' + str(
                                j + 1) + ') > td:nth-child(5)', type='CSS')
                        # 数据来源
                        source = tianyancha.FindElement('#_container_clients > div > table > tbody > tr:nth-child(' + str(
                                j + 1) + ') > td:nth-child(6)', type='CSS')
                        # 关联关系
                        relation = tianyancha.FindElement('#_container_clients > div > table > tbody > tr:nth-child(' + str(
                                j + 1) + ') > td:nth-child(7)', type='CSS')
                        client = client.text
                        percentage = percentage.text
                        price = price.text
                        date = date.text
                        source = source.text
                        relation = relation.text
                        print(company, nums, tianyancha.driver.current_url, client, percentage, price, date, source,relation)
                        csv_writer6.writerow([company, nums, tianyancha.driver.current_url, client, percentage, price, date, source,relation])
                    # 下一页
                    if page != pages:
                        if page == 1:
                            element = tianyancha.FindElement("#_container_clients > div > div > ul > li:nth-child(" + str(
                                pages + 1) + ") > a", type="CSS")
                        else:
                            element = tianyancha.FindElement("#_container_clients > div > div > ul > li:nth-child(" + str(
                                pages + 2) + ") > a", type="CSS")
                        element.click()
                    page += 1
            except Exception as e:
                print(tianyancha.driver.current_url, "获取客户信息失败")
                csv_writer2.writerow([company, nums, tianyancha.driver.current_url, "获取客户信息失败"])

    except Exception as e:
        print(tianyancha.driver.current_url, "机器人验证或企业名称不匹配，抓取失败")
        csv_writer2.writerow([company, nums, tianyancha.driver.current_url, "机器人验证或企业名称不匹配，抓取失败"])
        # time.sleep(random.uniform(MinSleepTime, MaxSleepTime))
    if tianyancha.driver.current_url.find('company') != -1:
        tianyancha.driver.close()
        tianyancha.driver.switch_to.window(tianyancha.driver.window_handles[0])
    # print('当前浏览地址为：{0}'.format(tianyancha.driver.current_url))
    Main_Staff.flush() # 解决缓冲区没能及时刷新导致写入失败的问题
    Failure_List.flush()
    Shareholder.flush()
    Change_Log.flush()
    Supplier.flush()
    Customer.flush()

    sleep(1)
    end = timeit.default_timer()
    print('Running time: %s Seconds' % (end - start))

# sleep(100)