from time import sleep
from pymysql import NULL
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver import ActionChains
import pymongo
import winsound


class MySeleniumCralwer:

  myClient = pymongo.MongoClient('localhost:27017')

  def __init__(self, dbName, colName, url="") -> None:
    self.driver = webdriver.Chrome()
    # 默认定位时间参数
    self.timeout = 0.5 # 最长的超时时间，单位为妙
    self.poll_frequency = 0.1 # 检测的步长，单位为秒
    # MongoDB参数
    self.db = self.myClient[dbName]
    self.col = self.db[colName]
    self.doc = {}

    if url != "":
      self.Get(url)
    
  def Get(self, url):
    try:
      print('访问链接: {}'.format(url))
      self.driver.get(url)
      return True
    except Exception as e:
      print('打开链接{}失败...等待处理(输入)... '.format(e))
      self.Alarm()
      return False 

  def Login(self, userpath, username, passpath, password, iframe='', type='XPATH'): # 要分情况考虑iframe, 后期可以加上定位方法的选择（CSS，XPATH，ID等）
    try:
      if iframe == '':
        wdw = WebDriverWait(self.driver, self.timeout)
        if type == 'XPATH':
          wdw.until(EC.presence_of_element_located((By.XPATH, userpath))).send_keys(username)
          wdw.until(EC.presence_of_element_located((By.XPATH, passpath))).send_keys(password)
        elif type == 'CSS':
          wdw.until(EC.presence_of_element_located((By.CSS_SELECTOR, userpath))).send_keys(username)
          wdw.until(EC.presence_of_element_located((By.CSS_SELECTOR, passpath))).send_keys(password)
        elif type == 'ID':
          wdw.until(EC.presence_of_element_located((By.ID, userpath))).send_keys(username)
          wdw.until(EC.presence_of_element_located((By.ID, passpath))).send_keys(password)
        elif type == 'TAG':
          wdw.until(EC.presence_of_element_located((By.TAG_NAME, userpath))).send_keys(username)
          wdw.until(EC.presence_of_element_located((By.TAG_NAME, passpath))).send_keys(password)
        else:
          print('暂不支持该定位方式,请选择XPATH、CSS、ID或TAG...等待处理(输入)...')
          self.Alarm()
          return False
      else:
        self.driver.switch_to.frame(iframe)
        wdw = WebDriverWait(self.driver, self.timeout)
        if type == 'XPATH':
          wdw.until(EC.presence_of_element_located((By.XPATH, userpath))).send_keys(username)
          wdw.until(EC.presence_of_element_located((By.XPATH, passpath))).send_keys(password)
        elif type == 'CSS':
          wdw.until(EC.presence_of_element_located((By.CSS_SELECTOR, userpath))).send_keys(username)
          wdw.until(EC.presence_of_element_located((By.CSS_SELECTOR, passpath))).send_keys(password)
        elif type == 'ID':
          wdw.until(EC.presence_of_element_located((By.ID, userpath))).send_keys(username)
          wdw.until(EC.presence_of_element_located((By.ID, passpath))).send_keys(password)
        elif type == 'TAG':
          wdw.until(EC.presence_of_element_located((By.TAG_NAME, userpath))).send_keys(username)
          wdw.until(EC.presence_of_element_located((By.TAG_NAME, passpath))).send_keys(password)
        else:
          print('暂不支持该定位方式,请选择XPATH、CSS、ID或TAG...等待处理(输入)...')
          self.Alarm()
          return False
      return True
    except Exception as e:
      print('登录失败...等待处理(输入)...')
      self.Alarm()
      return False

  def ClickElement(self, element):
    try:
      element.click()
      return True
    except Exception as e:
      print('点击失败...等待处理(输入)...')
      self.Alarm()
      return False

  def Click(self, path, poll_frequency = 0, timeout = 0, tag='', type= 'XPATH'):
    this_poll_frequency = 0
    this_timeout = 0
    if poll_frequency == 0:
      this_poll_frequency = self.poll_frequency
    else:
      this_poll_frequency = poll_frequency
    if timeout == 0:
      this_timeout = self.timeout
    else:
      this_timeout = timeout

    try:
      wdw = WebDriverWait(self.driver, timeout=this_timeout, poll_frequency=this_poll_frequency)
      if type == 'XPATH':
        wdw.until(EC.presence_of_element_located((By.XPATH, path))).click()
      elif type == 'CSS':
        wdw.until(EC.presence_of_element_located((By.CSS_SELECTOR, path))).click()
      elif type == 'ID':
        wdw.until(EC.presence_of_element_located((By.ID, path))).click()
      elif type == 'TAG':
        wdw.until(EC.presence_of_element_located((By.TAG_NAME, path))).click()
      else:
        print('暂不支持该定位方式,请选择XPATH、CSS、ID或TAG...等待处理(输入)...')
        self.Alarm()
        return False
      return True
    except Exception as e:
      print(tag + '点击失败...等待处理(输入)...')
      self.Alarm()
      return False

  def Input(self, path, poll_frequency = 0, timeout = 0, content='', tag = '', type='XPATH'):
    this_poll_frequency = 0
    this_timeout = 0
    if poll_frequency == 0:
      this_poll_frequency = self.poll_frequency
    else:
      this_poll_frequency = poll_frequency
    if timeout == 0:
      this_timeout = self.timeout
    else:
      this_timeout = timeout

    try:
      wdw = WebDriverWait(self.driver, timeout=this_timeout, poll_frequency=this_poll_frequency)
      if type == 'XPATH':
        wdw.until(EC.presence_of_element_located((By.XPATH, path))).send_keys(content)
      elif type == 'CSS':
        wdw.until(EC.presence_of_element_located((By.CSS_SELECTOR, path))).send_keys(content)
      elif type == 'ID':
        wdw.until(EC.presence_of_element_located((By.ID, path))).send_keys(content)
      elif type == 'TAG':
        wdw.until(EC.presence_of_element_located((By.TAG_NAME, path))).send_keys(content)
      else:
        print('暂不支持该定位方式,请选择XPATH、CSS、ID或TAG...等待处理(输入)...')
        self.Alarm()
        return False
      return True
    except Exception as e:
      print(tag + '输入失败...等待处理(输入)...')
      self.Alarm()
      return False

  def FindElements(self, path, poll_frequency = 0, timeout = 0, tag='', type='XPATH'): # 返回多个element
    this_poll_frequency = 0
    this_timeout = 0
    if poll_frequency == 0:
      this_poll_frequency = self.poll_frequency
    else:
      this_poll_frequency = poll_frequency
    if timeout == 0:
      this_timeout = self.timeout
    else:
      this_timeout = timeout

    try:
      wdw = WebDriverWait(self.driver, timeout=this_timeout, poll_frequency=this_poll_frequency)
      if type == 'XPATH':
        elements = wdw.until(EC.presence_of_all_elements_located((By.XPATH, path)))
      elif type == 'CSS':
        elements = wdw.until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, path)))
      elif type == 'ID':
        elements = wdw.until(EC.presence_of_all_elements_located((By.ID, path)))
      elif type == 'TAG':
        elements = wdw.until(EC.presence_of_all_elements_located((By.TAG_NAME, path)))
      else:
        print('暂不支持该定位方式,请选择XPATH、CSS、ID或TAG...等待处理(输入)...')
        self.Alarm()
        return NULL
      return elements
    except Exception as e:
      print(tag + path + '元素不存在...')
      # self.Alarm()
      return NULL
    
  def FindElement(self, path, poll_frequency = 0, timeout = 0, tag='', type='XPATH'): # 返回单个element
    this_poll_frequency = 0
    this_timeout = 0
    if poll_frequency == 0:
      this_poll_frequency = self.poll_frequency
    else:
      this_poll_frequency = poll_frequency
    if timeout == 0:
      this_timeout = self.timeout
    else:
      this_timeout = timeout

    try:
      wdw = WebDriverWait(self.driver, timeout=this_timeout, poll_frequency=this_poll_frequency)
      if type == 'XPATH':
        element = wdw.until(EC.presence_of_element_located((By.XPATH, path)))
      elif type == 'CSS':
        element = wdw.until(EC.presence_of_element_located((By.CSS_SELECTOR, path)))
      elif type == 'ID':
        element = wdw.until(EC.presence_of_element_located((By.ID, path)))
      elif type == 'TAG':
        element = wdw.until(EC.presence_of_element_located((By.TAG_NAME, path)))
      else:
        print('暂不支持该定位方式,请选择XPATH、CSS、ID或TAG...等待处理(输入)...')
        self.Alarm()
        return NULL
      return element
    except Exception as e:
      print(tag + path +  '元素不存在')
      # self.Alarm()
      return NULL
  
  def Alarm(self):
    print("触发告警，请及时处理")
    for a in range(3):
      winsound.PlaySound(f'{"14543.wav"}', winsound.SND_ALIAS)
    input()
  
  # def Current_Url(self):
  #   return self.driver.current_url

  def AddData(self, key, value):
    self.doc[key] = value
  
  def InsertOne(self):
    self.col.insert_one(self.doc)

  def ClearDoc(self):
    self.doc.clear()  

