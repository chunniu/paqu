import time
import xlwt
import xlrd
import datetime
from selenium import webdriver

from selenium.webdriver.common.proxy import Proxy
from selenium.webdriver.common.proxy import ProxyType
from selenium.webdriver.common.by import By

# 设置代理服务器
# options = webdriver.ChromeOptions()
# options.proxy = Proxy({'proxyType': ProxyType.MANUAL, 'httpProxy': 'http://127.0.0.1'})
start_time = time.time()
# 查询参数
rbx = xlrd.open_workbook('./keys.xls')
sr = rbx.sheet_by_index(0)
queries = []

for i in range(sr.nrows):
    queries.append(sr.cell_value(i, 0))

for query in queries:

    print(f'查询词语:{query}')

    try_times_end = 10  # 重试次数
    try_times = 1  # 重试次数开始
    page = 10  # 翻页数量
    time_sleep = 3  # 休眠时间-秒

    # 下滑滑动次数，page=10 代表 翻页 10.
    print(f'查询数据:{query}')
    print(f'翻页次数:{page}')
    # 如果重试次数 > 10次，则跳出循环
    print(f'重试次数:{try_times}')
    print(f'休眠时间:{time_sleep}')

    driver = webdriver.Chrome()
    driver.get(f"https://www.youtube.com/results?search_query={query}")

    # 获取 youtube app应用
    ytd_app = driver.find_element(By.TAG_NAME, "ytd-app")
    driver.implicitly_wait(0.5)

    last_ytd_app_height = ytd_app.size.get('height')

    # 滚动到页面底部 翻页次数
    while True:
        print(f'执行翻页:{page}')
        time.sleep(time_sleep)
        driver.execute_script("window.scrollTo(0,document.getElementsByTagName('ytd-app')[0].scrollHeight)")

        new_ytd_app_height = ytd_app.size.get('height')
        if new_ytd_app_height > last_ytd_app_height:
            page -= 1
            # 成功一次，重新计数
            try_times = 1
            last_ytd_app_height = new_ytd_app_height

        else:
            try_times += 1
            print(f'重试次数:{try_times}')
            if try_times > try_times_end:
                break

        if page < 0:
            break

    # 获取页面数据链接
    print('开始获取页面的a标签')
    channel_names = driver.find_elements(By.TAG_NAME, "ytd-channel-name")
    links = set()
    for channel in channel_names:
        try:
            l = channel.find_element(By.TAG_NAME, "a").get_attribute('href')
            links.add(l)
        except Exception as e:
            print(f'抓取数据异常:{channel},{e}')

    driver.implicitly_wait(10.5)
    # 获取页面
    print(f'获取的链接数量{len(links)}')
    for link in links:
        print(f'开始抓取:{link}')

    # 链接依次打开
    # 点击展开
    # 获取弹窗内容


    links_count = len(links)
    indx = 1
    content_list = []

    for link in links:
        print(f'抓取:{link},第{indx}条。总共{links_count}条。还剩下:{links_count - indx}条')
        # 链接优化一下，如果有 ,Messsage 则去掉
        link2 = link.replace(',Message', '')
        print(f'优化后的链接:{link2}')
        driver.get(link2)
        time.sleep(time_sleep)
        find_text = ''
        state = '失败'
        except_text = ''
        # 获取 page-header-view-model-wiz__page-header-content-metadata
        try:
            model = driver.find_element(By.TAG_NAME, "truncated-text-content")
            model.click()
            time.sleep(time_sleep)
            model_show = driver.find_element(By.TAG_NAME, 'ytd-engagement-panel-section-list-renderer')

            state = '成功'
            find_text = model_show.text
        except Exception as e:
            print(f'抓取数据异常:{link2},{e}')
            state = '异常'
            except_text = f'{e}'
        finally:
            content_list.append({
                'index':indx,
                'state':state,
                'link':link2,
                'find_text':find_text,
                'except_text':except_text,
            })
        indx += 1
        time.sleep(time_sleep)

    # 存入 excel
    print('开始写入数据')
    workbook = xlwt.Workbook()
    sheet = workbook.add_sheet('sheet1')
    sheet.write(0, 0, '序号')
    sheet.write(0, 1, '状态')
    sheet.write(0, 2, '链接')
    sheet.write(0, 3, '内容')
    sheet.write(0, 4, '异常')
    sheet.write(0, 5, query)
    for i in range(len(content_list)):
        sheet.write(i + 1, 0, content_list[i]['index'])
        sheet.write(i + 1, 1, content_list[i]['state'])
        sheet.write(i + 1, 2, content_list[i]['link'])
        sheet.write(i + 1, 3, content_list[i]['find_text'])
        sheet.write(i + 1, 4, content_list[i]['except_text'])

    workbook.save(f'{datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")}.xls')
    print('写入数据完成')

    time.sleep(1)

driver.quit()

print(f'程序执行时间:{time.time() - start_time}')

input("Press Enter to exit")

if __name__ == '__main__':
    pass

# See PyCharm help at https://www.jetbrains.com/help/pycharm/
