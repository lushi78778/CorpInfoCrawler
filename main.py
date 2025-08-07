import time
import ddddocr
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException

# ==============================================================================
# --- 1. 配置 ---
# 在此区域集中管理所有需要用户手动修改的变量，方便维护

# 您的平台登录用户名
YOUR_USERNAME = "*"
# 您的平台登录密码
YOUR_PASSWORD = "*"

# 您希望循环爬取的企业条目数量
ITEMS_TO_SCRAPE = 10
# 最终输出的 Excel 文件名
OUTPUT_FILENAME = "scraped_data.xlsx"
# ==============================================================================


# --- 2. 网站和元素定位 (XPaths) ---
# 将所有元素的定位符（XPath）集中存放在这里，如果未来网站改版，只需修改此区域

# 登录页面的URL
login_url = "https://shad.samr.gov.cn/#/login"
# 登录后需要访问的企业列表主页URL
enterprise_url = "https://shad.samr.gov.cn/#/enterprise"

# 用户名输入框的XPath
username_xpath = '//*[@id="app"]/div/div[1]/div/div/div[2]/div[2]/div[2]/form/div[1]/div/div/div/div[1]/input'
# 密码输入框的XPath
password_xpath = '//*[@id="app"]/div/div[1]/div/div/div[2]/div[2]/div[2]/form/div[2]/div/div/div/div[1]/input'
# 验证码输入框的XPath
captcha_input_xpath = '//*[@id="app"]/div/div[1]/div/div/div[2]/div[2]/div[2]/form/div[3]/div[1]/div/div/div/div/input'
# 验证码图片的XPath
captcha_image_xpath = '//*[@id="v_container"]/img'
# 登录按钮的XPath
login_button_xpath = '//*[@id="app"]/div/div[1]/div/div/div[2]/div[2]/div[2]/form/div[4]/button'

# 【后续操作部分】
# 社区选择的下拉框XPath
community_dropdown_xpath = '//*[@id="app"]/div/section/div/div/div/div[2]/div[1]/div[1]/div[1]/div[2]/div/input'
# 下拉框中特定社区选项的XPath
community_option_xpath = '//*[@id="app"]/div/section/div/div/div/div[2]/div[1]/div[1]/div[1]/div[2]/div[2]/div[2]/div[2]'
# “查询”按钮的XPath
query_button_xpath = '//*[@id="app"]/div/section/div/div/div/div[2]/div[1]/div[2]/div[1]'
# “每页显示条数”下拉框的XPath
pagesize_dropdown_xpath = '//*[@id="app"]/div/section/div/div/div/div[3]/div[3]/div/span[2]/div/div/input'
# “100条/页”选项的XPath (这是一个更健壮的定位方式，通过文本内容查找)
pagesize_100_option_xpath = "//div[contains(@class, 'el-select-dropdown') and not(contains(@style,'display: none'))]//li//span[text()='100条/页']"


# --- 3. 辅助函数：爬取详情页所有资料 ---
def scrape_all_detail_data(driver):
    """
    这个函数负责在详情页弹窗出现后，爬取所有可见的字段数据。
    它被设计为在主循环中调用，以保持主循环的逻辑清晰。
    """
    # 创建一个空字典，用于存储当前企业的所有信息
    data = {}

    def get_value(xpath):
        """
        一个内部辅助函数，用于安全地获取元素值。
        它会尝试查找指定的XPath，如果找到就返回值，如果找不到就返回'未找到'，避免程序因某个字段缺失而崩溃。
        """
        try:
            # 尝试通过XPath定位元素
            element = driver.find_element(By.XPATH, xpath)
            # 判断元素是否是输入框(<input>)
            if element.tag_name == 'input':
                # 如果是输入框，我们获取的是它里面预填的值(value)
                return element.get_attribute('value')
            else:
                # 如果是其他元素(如div, span, font)，我们获取它显示的文本(text)
                return element.text.strip()
        except NoSuchElementException:
            # 如果根据XPath找不到这个元素，就返回'未找到'
            return "未找到"

    print("   [爬取] 正在提取「基本信息」...")
    # --- 基本信息区 ---
    # 逐一提取每个字段的值
    data['统一社会信用代码'] = get_value("//label[text()='统一社会信用代码']/following-sibling::div//input")
    data['登记状态'] = get_value("//label[text()='登记状态']/following-sibling::div//input")
    data['主体名称'] = get_value("//label[text()='主体名称']/following-sibling::div//input")
    data['主体状态'] = get_value("//label[text()='主体状态']/following-sibling::div//input")
    data['法定代表人'] = get_value("//label[text()='法定代表人']/following-sibling::div//input")
    data['主体类别'] = get_value("//label[text()='主体类别']/following-sibling::div//input")
    # “主体级别”是一个单选按钮组，我们定位那个被选中的(class='is-active')按钮，并获取其文本
    data['主体级别'] = get_value(
        "//label[text()='主体级别']/following-sibling::div//label[contains(@class, 'is-active')]/span")
    data['主体电话'] = get_value("//label[text()='主体电话']/following-sibling::div//input")
    data['主体网址'] = get_value("//label[text()='主体网址']/following-sibling::div//input")
    data['从业人数'] = get_value("//label[text()='从业人数']/following-sibling::div//input")
    data['营业收入'] = get_value("//label[text()='营业收入']/following-sibling::div//input")
    data['负责人'] = get_value("//label[text()='负责人']/following-sibling::div//input")
    # “职务”字段在页面中有多个，我们通过父级元素来限定范围，确保取到的是“负责人”那一行的职务
    data['负责人职务'] = get_value(
        "//div[.//label[text()='负责人']]//label[text()='职务']/following-sibling::div//input")
    data['负责人身份证号'] = get_value(
        "//div[.//label[text()='负责人']]//label[text()='身份证号']/following-sibling::div//input")
    data['负责人手机号'] = get_value(
        "//div[.//label[text()='负责人']]//label[text()='手机号']/following-sibling::div//input")

    # 特殊处理：拼接的「住所」字段
    try:
        # “住所”由多个部分组成，首先找到前面省、市、区、街道的文本部分
        address_parts = driver.find_elements(By.XPATH,
                                             "//label[text()='住所']/following-sibling::div/div[contains(@class,'jbselect')]//font")
        # 将这些部分用空格连接起来
        address_main = " ".join([part.text for part in address_parts])
        # 再找到后面的详细地址输入框
        address_detail = driver.find_element(By.XPATH,
                                             "//label[text()='住所']/following-sibling::div/div[contains(@class,'el-input')]/input").get_attribute(
            'value')
        # 将所有部分拼接成一个完整的地址
        data['住所'] = f"{address_main} {address_detail}".strip()
    except Exception:
        # 如果拼接过程中任何一步出错，则记录为'未找到'
        data['住所'] = "未找到"

    print("   [爬取] 正在提取「安全责任人」...")
    # --- 安全责任人区 ---
    # 这里的字段名（如职务、手机号）有重复，因此定位时必须先找到唯一的父级标签（如“食品安全总监”）来限定范围
    data['食品安全总监姓名'] = get_value("//label[text()='食品安全总监']/following-sibling::div//input")
    data['食品安全总监职务'] = get_value(
        "//div[.//label[text()='食品安全总监']]//label[text()='职务']/following-sibling::div//input")
    data['食品安全总监身份证号'] = get_value(
        "//div[.//label[text()='食品安全总监']]//label[text()='身份证号']/following-sibling::div//input")
    data['食品安全总监手机号'] = get_value(
        "//div[.//label[text()='食品安全总监']]//label[text()='手机号']/following-sibling::div//input")
    data['食品安全员姓名'] = get_value("//label[text()='食品安全员']/following-sibling::div//input")
    data['食品安全员职务'] = get_value(
        "//div[.//label[text()='食品安全员']]//label[text()='职务']/following-sibling::div//input")
    data['食品安全员身份证号'] = get_value(
        "//div[.//label[text()='食品安全员']]//label[text()='身份证号']/following-sibling::div//input")
    data['食品安全员手机号'] = get_value(
        "//div[.//label[text()='食品安全员']]//label[text()='手机号']/following-sibling::div//input")

    print("   [爬取] 正在提取「登记/备案/许可证信息」...")
    # --- 登记/备案/许可证信息区 ---
    # 为了精确定位，先找到这个区域的父容器
    license_section_xpath = "//div[contains(@class, 'xkzxx')]"
    # 然后所有此区域的查询都在这个父容器内部进行
    data['备案编号'] = get_value(f"{license_section_xpath}//label[text()='备案编号']/following-sibling::div//input")
    data['许可证状态'] = get_value(
        f"{license_section_xpath}//label[text()='登记/备案/许可证状态']/following-sibling::div//input")
    data['许可证-统一社会信用代码/身份证号'] = get_value(
        f"{license_section_xpath}//label[contains(text(), '统一社会信用代码')]/following-sibling::div//input")
    data['许可证-法定代表人'] = get_value(
        f"{license_section_xpath}//label[text()='法定代表人(负责人)']/following-sibling::div//input")
    data['食品经营者名称'] = get_value(
        f"{license_section_xpath}//label[text()='食品经营者名称']/following-sibling::div//input")
    data['许可证-法定代表人联系方式'] = get_value(
        f"{license_section_xpath}//label[contains(text(), '联系方式')]/following-sibling::div//input")
    data['经营场所住所'] = get_value(
        f"{license_section_xpath}//label[text()='经营场所住所']/following-sibling::div//input")

    # 所有字段提取完毕，返回包含这些数据的字典
    return data


# --- 4. 主程序 ---
# 初始化Chrome浏览器驱动
driver = webdriver.Chrome()
# 设置一个全局的智能等待，所有wait.until()最多等待20秒
wait = WebDriverWait(driver, 20)
# 最大化浏览器窗口，确保所有元素都在可视范围内
driver.maximize_window()
# 创建一个空列表，用于之后存储从每个企业详情页爬取到的数据字典
all_scraped_data = []

# 使用 try...except...finally 结构，确保程序无论成功还是失败，最后都能安全地关闭浏览器和保存数据
try:
    # === 步骤 1-9 : 登入并完成页面设定的完整流程 ===
    print("➡️ [流程开始] 正在执行登入和页面设定...")
    # 1. 打开登录页
    driver.get(login_url)

    # 2. 填写帐密
    wait.until(EC.presence_of_element_located((By.XPATH, username_xpath))).send_keys(YOUR_USERNAME)
    driver.find_element(By.XPATH, password_xpath).send_keys(YOUR_PASSWORD)
    time.sleep(2)

    # 3. OCR识别并填写验证码
    # 等待验证码图片可见
    captcha_element = wait.until(EC.visibility_of_element_located((By.XPATH, captcha_image_xpath)))
    # 初始化ddddocr
    ocr = ddddocr.DdddOcr()
    # 对验证码图片元素进行截图，并传入OCR进行识别
    recognized_text = ocr.classification(captcha_element.screenshot_as_png)
    print(f"   [INFO] OCR 识别结果: {recognized_text}")
    # 将识别结果填入输入框
    driver.find_element(By.XPATH, captcha_input_xpath).send_keys(recognized_text)

    # 4. 点击登录并主动跳转
    # 等待登录按钮变为可点击状态
    login_button = wait.until(EC.element_to_be_clickable((By.XPATH, login_button_xpath)))
    # 使用JavaScript点击，这种方式通常比物理点击更稳定，能穿透一些覆盖物
    driver.execute_script("arguments[0].click();", login_button)
    time.sleep(5)  # 在此处暂停5秒，等待服务器处理登录请求并设置cookie
    # 主动访问目标页面，这对于单页面应用(SPA)是关键一步
    driver.get(enterprise_url)
    time.sleep(3)

    # 5. 验证登录成功并进行页面筛选设置
    # 等待社区下拉框可点击
    community_dropdown = wait.until(EC.element_to_be_clickable((By.XPATH, community_dropdown_xpath)))
    # 点击下拉框以展开选项
    community_dropdown.click()
    # 等待目标社区选项可点击，然后点击它
    wait.until(EC.element_to_be_clickable((By.XPATH, community_option_xpath))).click()
    time.sleep(1)
    # 等待“查询”按钮可点击，然后点击它
    wait.until(EC.element_to_be_clickable((By.XPATH, query_button_xpath))).click()
    time.sleep(1)

    # 6. 设定每页显示100条
    # 等待分页大小的下拉框可点击
    pagesize_dropdown_element = wait.until(EC.element_to_be_clickable((By.XPATH, pagesize_dropdown_xpath)))
    # 使用JavaScript将此元素滚动到可视区域，防止因被遮挡而无法点击
    driver.execute_script("arguments[0].scrollIntoView(true);", pagesize_dropdown_element)
    # 点击下拉框
    pagesize_dropdown_element.click()
    # 等待“100条/页”选项出现并点击
    wait.until(EC.element_to_be_clickable((By.XPATH, pagesize_100_option_xpath))).click()

    print("✅ [流程完成] 登入和页面设定成功。")
    # 智能等待，直到列表的第一个项目出现，确保列表已按100条/页的标准刷新完成
    wait.until(EC.visibility_of_element_located((By.XPATH, "//div[@data-v-2fe107be and @class='itemList']")))

    # === 步骤 10: 循环爬取主体列表 ===
    print(f"\n➡️ [爬取开始] 准备循环爬取，目标处理 {ITEMS_TO_SCRAPE} 笔资料...")

    # 根据您在配置区设定的数量，开始循环
    for i in range(ITEMS_TO_SCRAPE):
        print("-" * 60)
        # 预设一个公司名称变量，以便在出错时也能打印
        company_name = "未知"
        try:
            # 关键：每次循环都重新查找当前页面的所有项目列表
            # 这是为了避免因页面跳转和刷新导致的 Stale Element Reference Exception (过时元素错误)
            item_list = wait.until(
                EC.presence_of_all_elements_located((By.XPATH, "//div[@data-v-2fe107be and @class='itemList']")))

            # 检查本页的实际项目数量是否少于我们想处理的索引号
            if i >= len(item_list):
                print(f"   [INFO] 页面上的项目 (共 {len(item_list)} 笔) 少于目标爬取索引 {i + 1}，爬取提前结束。")
                # 如果是，说明本页已处理完，跳出循环
                break

            # 从刚找到的最新列表中，获取当前循环次序对应的那个项目
            current_item = item_list[i]

            # 从当前项目中提取公司名称，用于日志记录和数据核对
            # 使用 .// 开头的相对XPath，表示只在 current_item 内部查找，效率更高且更稳定
            company_name_element = current_item.find_element(By.XPATH, ".//div[contains(@class, 'itemName')]")
            company_name = company_name_element.get_attribute('title')
            print(f"   [处理中] 正在处理第 {i + 1} 个主体: {company_name}")
            time.sleep(1)

            # 在当前项目内部找到“编辑”按钮
            edit_button = current_item.find_element(By.XPATH, ".//div[contains(@class, 'btnFp')]")
            # 同样使用JS点击，确保能点上
            driver.execute_script("arguments[0].click();", edit_button)
            time.sleep(2)

            # --- 核心步骤：调用函数爬取详情页所有数据 ---
            scraped_data = scrape_all_detail_data(driver)
            # 在爬取到的数据中，额外加入一个字段，记录它在列表页的原始名称，方便后续数据核对
            scraped_data['列表页-主体名称'] = company_name
            # 将这一个企业的数据字典，追加到总的数据列表中
            all_scraped_data.append(scraped_data)
            print(f"   [成功] 资料提取完毕。")

            # --- 返回列表页 ---
            print("   [导航] 返回列表页...")
            # driver.back() 是最标准的返回上一页的方式，它会处理好浏览器的历史记录
            driver.back()

            # --- 等待列表页重新加载完成 ---
            # 这是至关重要的一步，必须确保返回后页面已准备好，才能安全地进行下一次循环
            print("   [等待] 等待列表页重新加载...")
            # 我们等待分页大小的下拉框再次可见，以此作为列表页加载成功的标志
            wait.until(EC.visibility_of_element_located((By.XPATH, pagesize_dropdown_xpath)))

        except Exception as item_error:
            # 如果在处理单个项目的过程中发生任何错误
            print(f"   [错误] 处理第 {i + 1} 个主体 ({company_name}) 时发生错误: {item_error}")
            print("   [策略] 跳过此项目，尝试继续处理下一个...")
            # 为防止错误卡在详情页，强制导航回列表页
            if "/enterprise" not in driver.current_url:
                driver.get(enterprise_url)
            # 跳过本次循环的剩余部分，直接进入下一次循环
            continue

    print("\n✅ [爬取完成] 指定数量的资料已处理完毕！")

# 捕获因等待超时而发生的错误
except Exception as e:
    print(f"\n❌ [程序错误] 发生未预期的错误: {e}")
    # 保存一张截图，用于分析程序出错时页面的状态
    screenshot_path = "error_screenshot.png"
    driver.save_screenshot(screenshot_path)
    print(f"错误发生时的页面截图已保存至: {screenshot_path}")
    input("按 Enter 键退出...")

# finally 语句块确保无论程序是成功结束还是中途报错，最后都一定会执行其中的代码
finally:
    # === 步骤 11: 将结果保存到 Excel ===
    # 检查我们是否成功爬取到了任何数据
    if all_scraped_data:
        print(f"\n➡️ [保存] 正在将 {len(all_scraped_data)} 笔资料保存至 {OUTPUT_FILENAME}...")
        try:
            # 使用pandas将数据列表转换成数据框(DataFrame)
            df = pd.DataFrame(all_scraped_data)
            # 调用 to_excel 方法将数据框写入Excel文件，index=False表示不将pandas的行索引写入文件
            df.to_excel(OUTPUT_FILENAME, index=False, engine='openpyxl')
            print(f"✅ [成功] 资料已成功保存至 {OUTPUT_FILENAME}！")
        except ImportError:
            # 如果用户的环境中没有安装 openpyxl
            print("⚠️ [警告] 未安装 'openpyxl'。请执行 'pip install openpyxl' 来支持 Excel 写入。")
        except Exception as save_error:
            # 捕获保存过程中可能发生的其他错误
            print(f"❌ [保存失败] 保存 Excel 时发生错误: {save_error}")
    else:
        # 如果列表为空，说明整个过程一条数据都没爬到
        print("\n⚠️ [注意] 没有成功爬取到任何资料，不创建 Excel 文件。")

    print("\n[结束] 正在关闭浏览器...")
    driver.quit()