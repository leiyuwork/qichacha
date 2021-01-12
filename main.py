# -*- coding: utf-8 -*-
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
import time
from selenium.webdriver import ActionChains
import pyautogui
import pyperclip


# 伪装成浏览器，防止被识破
option = Options()
option.add_experimental_option('excludeSwitches', ['enable-automation'])
option.add_experimental_option('useAutomationExtension', False)
option.add_argument("--no-sandbox")
option.add_argument("--disable-dev-usage")
option.add_argument(
    '--user-agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/65.0.3325.146 Safari/537.36"')
driver = webdriver.Chrome(options=option)

driver.execute_cdp_cmd("Page.addScriptToEvaluateOnNewDocument", {
  "source": """
    Object.defineProperty(navigator, 'webdriver', {
      get: () => undefined
    })
  """
})


# 打开登录页面
driver.get('https://www.qichacha.com/user_login')
driver.maximize_window()

# 单击用户名密码登录的标签
tag = driver.find_element_by_xpath('//*[@id="normalLogin"]')
tag.click()

# 将用户名、密码注入
driver.find_element_by_id('nameNormal').send_keys('17061222592')
driver.find_element_by_id('pwdNormal').send_keys('aa411229')
time.sleep(5)  # 休眠，人工完成验证步骤，等待程序单击“登录”

#获取滑块
button = driver.find_element_by_id('nc_1_n1z')

# 滑动滑块
ActionChains(driver).click_and_hold(button).perform()
ActionChains(driver).move_by_offset(xoffset=308, yoffset=0).perform()
ActionChains(driver).release().perform()

time.sleep(3)

# 单击登录按钮
btn = driver.find_element_by_xpath('//*[@id="user_login_normal"]/button')
btn.click()
time.sleep(3)

inc_list = ['娃哈哈集团','阿海珐输配电集团','阿克苏诺贝尔中国有限公司','阿斯克（中国）软件技术有限公司','阿斯特尼电气（北京）有限公司','阿斯利康','阿斯利康（无锡）制药有限公司','阿森纳投资顾问（上海）有限公司','阿特拉斯·科普柯中国香港公司','阿里影业','阿里影业股份有限公司','阿里巴巴','阿里巴巴（中国）网络技术有限公司','阿里巴巴影业集团','阿里巴巴集团','阿里巴巴美国公司','阿呀呀投资管理有限公司','阿拉山口市雅本创业投资有限公司','阿拉善SEE公益环保两届','阿迪达斯（苏州）有限公司','阿尔西制冷工程技术（北京）有限公司','阿尔创（广州）信息技术有限公司','阿尔卡特朗讯（中国）公司','阿尔卡特朗讯等公司','阿谢资本管理有限公司','旭化成管理（上海）有限公司','旭日玩具厂','旭电科技','安永（珠海）国际保健品有限公司','安永（中国）企业咨询有限公司','安永会计师事务所','安永会计师事务所（特殊普通合伙）','安永信财经顾问公司','安永大华会计师事务所','安永大华会计师事务所有限公司（原大华会计师事务所）','安永华明会计师事务所','安永华明会计师事务所（特殊普通合伙）','安永华明会计师事务所有限责任公司','安益（杭州）财富管理有限公司','安科恒益公司','安科瑞电气股份有限公司（原“上海安科瑞电气股份有限公司”）','安岳县乡镇企业管理局','安徽安科生物工程（集团）股份有限公司','安徽安科生物高技术公司','安徽安科余良卿药业有限公司','安徽安元投资基金管理有限公司','安徽安泰达律师事务所','安徽安利科技投资集团股份有限公司（以下简称“安利投资”）','安徽安利合成革股份有限公司','安徽安利合成革有限公司','安徽安利合成革有限公司（以下简称“安利有限”）','安徽安利材料科技股份有限公司','安徽安利材料科技股份有限公司（以下简称“安利股份”）','安徽安凯客车股份有限公司','安徽安德利百货股份有限公司','安徽安纳达钛业股份有限公司','安徽医科大学','安徽云松投资管理有限公司','安徽黄山经济技术发展公司','安徽黄山胶囊股份有限公司','安徽科苑集团股份有限公司','安徽科佳电脑有限公司','安徽科大国创云网科技有限公司','安徽科大智能物流系统有限公司','安徽科大智能电网技术有限公司','安徽科大讯飞信息科技股份有限公司','安徽科大鲁能科技有限公司','安徽会计师事务所','安徽海思达机器人有限公司','安徽海丰精细化工股份公司','安徽海宁投资实业有限公司','安徽外经建设集团公司','安徽徽风报刊营销网络公司','安徽玉山食品公司','安徽欣意电缆有限公司','安徽九华山旅游发展股份公司','安徽桑尼生物技术研究所','安徽桑乐金股份有限公司','安徽桑铌科技股份有限公司','安徽慧通互联科技有限公司','安徽迎驾贡酒股份有限公司','安徽健友律师事务所','安徽建筑工业学院','安徽建筑大学','安徽古井集团有限责任公司','安徽古井雪地啤酒有限责任公司','安徽古井贡酒股份有限公司','安徽古南岳农业发展有限公司','安徽光机所受控站实习','安徽宏宇五洲医疗器械股份有限公司','安徽宏晶新材料股份有限公司','安徽工学院','安徽工商管理学院','安徽工程大学','安徽工业大学','安徽恒源煤电股份有限公司','安徽江南化工股份有限公司','安徽江淮汽车股份有限公司','安徽合肥振安医学检验用品公司','安徽合肥制药厂','安徽国安创业投资有限公司','安徽国元安通有限责任公司','安徽国元实业投资公司','安徽国信实业有限公司','安徽国通高新管业股份有限公司','安徽国祯集团股份有限公司','安徽国祯环保节能科技股份有限公司','安徽国际信托投资有限公司','安徽国际经济技术合作公司','安徽国风集团公司','安徽国风集团有限公司','安徽国风塑业股份有限公司','安徽酷智投资管理有限公司','安徽三建集团','安徽三联交通应用技术有限公司','安徽山河药用辅料股份有限公司','安徽四创电子股份有限公司（600990）','安徽子公司','安徽首泰东方资产管理有限公司','安徽商业高等专科学校','安徽省安泰科技股份有限公司','安徽省安通发展有限公司','安徽省安庆市华茂股份有限公司','安徽省安庆毛纺厂','安徽省委党校经济管理教研室','安徽省委办公厅','安徽省医药行业协会','安徽省医药集团公司','安徽省医药商业协会','安徽省黄山链条厂','安徽省黄山风景区管理委员会','安徽省科协','安徽省会计学会','安徽省界首市第二中学','安徽省核学会','安徽省企业经营管理研究会','安徽省机械化粮库','安徽省机械科学研究所','安徽省渠道网络股份有限公司','安徽省金融会计学会','安徽省金融学会','安徽省光子器件与材料重点实验室','安徽省高新技术产业投资公司','安徽省国家税务局培训中心','安徽省国际信托投资公司','安徽省司尔特肥业股份有限公司','安徽省社会科学联合会','安徽省省直机关青联常委','安徽省信息化专家咨询委员会','安徽省信息产业投资控股有限公司','安徽省信托投资公司','安徽省新能源协会','安徽省人民检察院','安徽省政府','安徽省生物化学与分子生物学学会','安徽省生物研究所','安徽省生物工程学会','安徽省青年法律工作者协会','安徽省池州师范专科学校（现池州学院）','安徽省茶叶学会','安徽省中安健康养老服务产业产业投资合伙企业（有限合伙）','安徽省中医药学会','安徽省中药研究与开发重点实验室','安徽省注册会计师协会','安徽省直工委工商分委会','安徽省投资集团控股有限公司','安徽省内部审计协会、','安徽省文胜生物工程股份有限公司','安徽省律师协会','安徽省无为县供销社','安徽省淮南市人民政府','安徽省淮南发电总厂','安徽省淮北市公安局','安徽省淮北市税务局','安徽省蚌埠市公安局','安徽省东至县五丰中学']
# inc_len = len(inc_list)
for company_name in inc_list:

    driver.get('https://www.qcc.com/web/search?key=' + company_name)


    # 单击第一个结果
    try:
        srh_rst = driver.find_element_by_xpath('/html/body/div[1]/div[2]/div[2]/div[4]/div/div[2]/div/table/tr[1]/td[3]/div/a')
        srh_rst.click()
        time.sleep(10)

        pyautogui.hotkey('ctrl','s')
        pyperclip.copy(company_name)
        time.sleep(2)
        pyautogui.hotkey('ctrl', 'v')
        pyautogui.press('tab')
        pyautogui.press("down")
        pyautogui.press("up")
        pyautogui.press("up")
        pyautogui.press('enter')
        pyautogui.press('enter')
        time.sleep(2)
        pyautogui.hotkey('ctrl', 'w')
        time.sleep(3)
    except:
        with open("C:\\Users\Ray94\Downloads\\" + company_name + "null.html", "w") as file:
            file.write(company_name + " is null")
