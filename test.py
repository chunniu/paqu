from selenium import webdriver
from selenium.webdriver.chrome.service import Service

# 创建 Service 对象
ser = Service()
ser.executable_path = f'C:\project\chrome-headless-shell-win64\chrome-headless-shell.exe'	# 指定 ChromeDriver 的路径

# 初始化 WebDriver，使用之前创建 Service 对象
driver = webdriver.Chrome(service=ser)

# 打开网页
driver.get('http://www.baidu.com')

# 关闭浏览器
driver.quit()