from math import radians
from operator import length_hint
from platform import release
from turtle import pos
from my.MySeleniumCralwer import *

douban = MySeleniumCralwer(dbName='douban3', colName='top250')

for i in range(10):
  # 25个电影/页，共10页
  douban.Get('https://movie.douban.com/top250?start=' + str(i * 25) + '&filter=')
  for j in range(1, 26):
    douban.Click('//ol[@class="grid_view"]/li[' + str(j) + ']/div/div[1]/a')

    # 电影名称
    name = douban.FindElement('//*[@id="content"]/h1/span[1]', '电影名称').text
    douban.AddData('name', name)

    # 海报
    poster = douban.FindElement('//a[@class="nbgnbg"]/img', '海报').get_attribute('src')
    douban.AddData('poster', poster)

    # 导演
    director = douban.FindElement('//*[@id="info"]/span[1]/span[2]/a', '导演').text
    douban.AddData('director', director)

    # 主演
    actorElements = douban.FindElements('//span[@class="attrs"]/span/a | //span[@class="attrs"]/a', '主演')
    actors = ''
    for element in actorElements:
      if element.text  != '':
        actors = actors + str(element.text) + '/'
    douban.AddData('actors', actors)

    # 电影种类
    typesElements = douban.FindElements('//span[@property="v:genre"]', '电影种类')
    types = ''
    for element in typesElements:
      if element.text  != '':
        types = types + str(element.text) + '/'
    douban.AddData('type', types)

    # 上映日期
    release_date_elements = douban.FindElements('//span[@property="v:initialReleaseDate"]', '上映日期')
    release_date = ''
    for element in release_date_elements:
      if element.text  != '':
        release_date = release_date + str(element.text) + '/'
    douban.AddData('release_date', release_date)
      
    # 片长
    length = douban.FindElement('//span[@property="v:runtime"]', '片长').text
    douban.AddData('length', length)

    # 豆瓣评分
    score = douban.FindElement('//strong[@class="ll rating_num"]', '豆瓣评分').text
    douban.AddData('score', score)

    # 评分人数
    evaluation_number = douban.FindElement('//span[@property="v:votes"]', '评分人数').text
    douban.AddData('evaluation_number', evaluation_number)

    # 简介
    if douban.FindElement('//a[@class="j a_show_full"]', poll_frequency=0.2, timeout=3, tag='展开全部') == NULL:
      introduction = douban.FindElement('//span[@property="v:summary"]').text
      douban.AddData('introduction', introduction)
    else:
      introduction = douban.FindElement('//span[@class="all hidden"]').get_attribute('innerHTML')
      douban.AddData('introduction', introduction)
    
    # except Exception as e:
    #   pass

    douban.InsertOne()
    print(douban.doc)
    douban.ClearDoc()

    douban.driver.back()
    sleep(2)
  

sleep(100)