from urllib.parse import quote
from urllib import request
import json
import os
import xlwt
import pandas as pd
from transCoordinateSystem import gcj02_to_wgs84, gcj02_to_bd09
from ratelimit import limits, sleep_and_retry

#from shp import trans_point_to_shp

'''
版本更新说明：

2019.10.05：
    1. 数据导出格式支持CSV格式以及XLS两种格式;
    2. 支持同时采集多个城市的POI数据;
    3. 支持同时采集多个POI分类数据

2019.10.10:
    1. 数据导出支持CSV以及XLS两种格式;
    2. CSV格式数据会生成.shp文件，可以直接在ARCGIS中使用

2020.06.19:
    1.清除了poi数据写入shp文件相关操作
    2.修改为根据POI分类关键字来爬取，而不是分类编码
'''

#################################################需要修改###########################################################

# TODO 1.替换为从高德开放平台上申请申请的密钥
amap_web_key = os.getenv('AMAP_WEB_KEY', 'default_key')

# TODO 2.分类关键字,最好对照<<高德地图POI分类关键字以及编码.xlsx>>来填写对应分类关键字(不是编码)，多个用逗号隔开
keyword = ['网球场']

# TODO 3.城市，多个用逗号隔开
city = ['杭州']

# TODO 4.输出数据坐标系,1为高德GCJ20坐标系，2WGS84坐标系，3百度BD09坐标系
coord = 1

# TODO 5. 输出数据文件格式,1为默认xls格式，2为csv格式
data_file_format = 2

############################################以下不需要动#######################################################################


poi_search_url = "http://restapi.amap.com/v3/place/text"
poi_boundary_url = "https://ditu.amap.com/detail/get/detail"


# 根据城市名称和分类关键字获取poi数据
@limits(calls=1, period=1)
def getpois(cityname, keywords):
    i = 1
    poilist = []
    while True:  # 使用while循环不断分页获取数据
        result = getpoi_page(cityname, keywords, i)
        print(result)
        result = json.loads(result)  # 将字符串转换为json
        if result['count'] == '0':
            break

        hand(poilist, result)
        i = i + 1
    return poilist


# 数据写入excel
def write_to_excel(poilist, cityname, classfield):
    # 一个Workbook对象，这就相当于创建了一个Excel文件
    book = xlwt.Workbook(encoding='utf-8', style_compression=0)
    sheet = book.add_sheet(classfield, cell_overwrite_ok=True)

    # 第一行(列标题)
    sheet.write(0, 0, 'lon')
    sheet.write(0, 1, 'lat')
    sheet.write(0, 2, 'name')
    sheet.write(0, 3, 'address')
    sheet.write(0, 4, 'pname')
    sheet.write(0, 5, 'cityname')
    sheet.write(0, 6, 'business_area')
    sheet.write(0, 7, 'type')

    for i in range(len(poilist)):
        location = poilist[i].get('location')
        name = poilist[i].get('name')
        address = poilist[i].get('address')
        pname = poilist[i].get('pname')
        cityname = poilist[i].get('cityname')
        business_area = poilist[i].get('business_area')
        type = poilist[i].get('type')
        lng = str(location).split(",")[0]
        lat = str(location).split(",")[1]

        if (coord == 2):
            result = gcj02_to_wgs84(float(lng), float(lat))
            lng = result[0]
            lat = result[1]
        if (coord == 3):
            result = gcj02_to_bd09(float(lng), float(lat))
            lng = result[0]
            lat = result[1]

        # 每一行写入
        sheet.write(i + 1, 0, lng)
        sheet.write(i + 1, 1, lat)
        sheet.write(i + 1, 2, name)
        sheet.write(i + 1, 3, address)
        sheet.write(i + 1, 4, pname)
        sheet.write(i + 1, 5, cityname)
        sheet.write(i + 1, 6, business_area)
        sheet.write(i + 1, 7, type)

    # 最后，将以上操作保存到指定的Excel文件中
    book.save(r'data' + os.sep + 'poi-' + cityname + "-" + classfield + ".xls")


# 数据写入csv文件中
def write_to_csv(poilist, cityname, classfield):
    lons, lats = [], []

    ids = [poi.get('id') for poi in poilist]
    names = [poi.get('name') for poi in poilist]
    addresss = [poi.get('address') for poi in poilist]
    pnames = [poi.get('pname') for poi in poilist]
    citynames = [poi.get('cityname') for poi in poilist]
    business_areas = [poi.get('business_area') for poi in poilist]
    types = [poi.get('type') for poi in poilist]

    for poi in poilist:
        print('===================')
        print(poi)

        location = poi.get('location')
        lng = str(location).split(",")[0]
        lat = str(location).split(",")[1]

        if (coord == 2):
            result = gcj02_to_wgs84(float(lng), float(lat))
            lng = result[0]
            lat = result[1]
        if (coord == 3):
            result = gcj02_to_bd09(float(lng), float(lat))
            lng = result[0]
            lat = result[1]
        lons.append(lng)
        lats.append(lat)

    df = pd.DataFrame({
        'id': ids,
        'lon': lons,
        'lat': lats,
        'name': names,
        'address': addresss,
        'pname': pnames,
        'cityname': citynames,
        'business_area': business_areas,
        'type': types
    })

    folder_name = 'poi-' + cityname + "-" + classfield
    folder_name_full = 'data' + os.sep + folder_name + os.sep
    if os.path.exists(folder_name_full) is False:
        os.makedirs(folder_name_full)

    file_name = 'poi-' + cityname + "-" + classfield + ".csv"
    file_path = folder_name_full + file_name

    sorted_df = df.sort_values(by=['id'], ascending=True)
    sorted_df.to_csv(file_path, index=False)
    return folder_name_full, file_name


# 将返回的poi数据装入集合返回
def hand(poilist, result):
    # result = json.loads(result)  # 将字符串转换为json
    pois = result['pois']
    for i in range(len(pois)):
        poilist.append(pois[i])


# 单页获取pois
@sleep_and_retry
@limits(calls=1, period=1)
def getpoi_page(cityname, keywords, page):
    req_url = poi_search_url + "?key=" + amap_web_key + '&extensions=all&keywords=' + quote(
        keywords) + '&city=' + quote(cityname) + '&citylimit=true' + '&offset=25' + '&page=' + str(
        page) + '&output=json'
    data = ''
    print('============请求url:' + req_url)
    with request.urlopen(req_url) as f:
        data = f.read()
        data = data.decode('utf-8')
    return data


def get_areas(code):
    '''
    获取城市的所有区域
    :param code:
    :return:
    '''

    print('获取城市的所有区域：code: ' + str(code).strip())
    data = get_distrinctNoCache(code)

    print('get_distrinct result:' + data)

    data = json.loads(data)

    districts = data['districts'][0]['districts']
    # 判断是否是直辖市
    # 北京市、上海市、天津市、重庆市。
    if (code.startswith('重庆') or code.startswith('上海') or code.startswith('北京') or code.startswith('天津')):
        districts = data['districts'][0]['districts'][0]['districts']

    i = 0
    area = ""
    for district in districts:
        name = district['name']
        adcode = district['adcode']
        i = i + 1
        area = area + "," + adcode

    print(area)
    print(str(area).strip(','))
    return str(area).strip(',')


def get_data(city, keyword):
    '''
    根据城市名以及POI类型爬取数据
    :param city:
    :param keyword:
    :return:
    '''
    isNeedAreas = True
    if isNeedAreas:
        # 获取城市下所有区的行政区划编码, 如 杭州 -> 西湖区/上城区的行政区划编码
        area = get_areas(city)
    all_pois = []
    if area != None and area != "":
        area_list = str(area).split(",")
        if area_list == 0:
            area_list = str(area).split("，")

        for area in area_list:
            pois_area = getpois(area, keyword)
            print('当前城区：' + str(area) + ', 分类：' + str(keyword) + ", 总的有" + str(len(pois_area)) + "条数据")
            all_pois.extend(pois_area)
        print("所有城区的数据汇总，总数为：" + str(len(all_pois)))
        if data_file_format == 2:
            # 写入CSV
            file_folder, file_name = write_to_csv(all_pois, city, keyword)
            # 写入SHP
            #trans_point_to_shp(file_folder, file_name, 0, 1)
            return
        return write_to_excel(all_pois, city, keyword)
    else:
        pois_area = getpois(city, keyword)
        if data_file_format == 2:
            # 写入CSV
            file_folder, file_name = write_to_csv(all_pois, city, keyword)
            # 写入SHP
            #trans_point_to_shp(file_folder, file_name, 0, 1)
            return
        return write_to_excel(pois_area, city, keyword)

    return None

@sleep_and_retry
@limits(calls=1, period=1)
def get_distrinctNoCache(code):
    '''
    获取中国城市行政区划
    :return:
    '''

    url = "https://restapi.amap.com/v3/config/district?subdistrict=2&extensions=all&key=" + amap_web_key

    req_url = url + "&keywords=" + quote(code)

    print(req_url)

    with request.urlopen(req_url) as f:
        data = f.read()
        data = data.decode('utf-8')
    print(code, data)
    return data


if __name__ == '__main__':

    for ct in city:
        for type in keyword:
            get_data(ct, type)
    print('总的', len(city), '个城市, ', len(keyword), '个分类数据全部爬取完成!')

