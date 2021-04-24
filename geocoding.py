# -*- coding: utf-8 -*-
import requests
import math

x_pi = 3.14159265358979324 * 3000.0 / 180.0
pi = 3.1415926535897932384626  # π
a = 6378245.0  # 长半轴
ee = 0.00669342162296594323  # 偏心率平方


class Geocoding:
    def __init__(self, api_key):
        self.api_key = api_key

    def geocode(self, address):
        """
        利用高德geocoding服务解析地址获取位置坐标
        :param address:需要解析的地址
        :return:
        """
        geocoding = {'s': 'rsv3',
                     'key': self.api_key,
                     'city': '全国',
                     'address': address}
        # geocoding = urllib.urlencode(geocoding)
        # ret = urllib.urlopen("http://restapi.amap.com/v3/geocode/geo{}".format(geocoding))
        url = "http://restapi.amap.com/v3/geocode/geo?"
        ret = requests.get(url, params=geocoding)
        if ret.status_code == 200:
            # res = ret.json()
            # json_obj = json.loads(res)
            json_obj = ret.json()
            if json_obj['status'] == '1' and int(json_obj['count']) >= 1:
                geocodes = json_obj['geocodes'][0]
                lng = float(geocodes.get('location').split(',')[0])
                lat = float(geocodes.get('location').split(',')[1])
                return [lng, lat]
            else:
                return None
        else:
            return None


def gcj02_to_bd09(lng, lat):
    """
    火星坐标系(GCJ-02)转百度坐标系(BD-09)
    谷歌、高德——>百度
    :param lng:火星坐标经度
    :param lat:火星坐标纬度
    :return:
    """
    z = math.sqrt(lng * lng + lat * lat) + 0.00002 * math.sin(lat * x_pi)
    theta = math.atan2(lat, lng) + 0.000003 * math.cos(lng * x_pi)
    bd_lng = z * math.cos(theta) + 0.0065
    bd_lat = z * math.sin(theta) + 0.006
    return [bd_lng, bd_lat]


def bd09_to_gcj02(bd_lon, bd_lat):
    """
    百度坐标系(BD-09)转火星坐标系(GCJ-02)
    百度——>谷歌、高德
    :param bd_lat:百度坐标纬度
    :param bd_lon:百度坐标经度
    :return:转换后的坐标列表形式
    """
    x = bd_lon - 0.0065
    y = bd_lat - 0.006
    z = math.sqrt(x * x + y * y) - 0.00002 * math.sin(y * x_pi)
    theta = math.atan2(y, x) - 0.000003 * math.cos(x * x_pi)
    gg_lng = z * math.cos(theta)
    gg_lat = z * math.sin(theta)
    return [gg_lng, gg_lat]


def wgs84_to_gcj02(lng, lat):
    """
    WGS84转GCJ02(火星坐标系)
    :param lng:WGS84坐标系的经度
    :param lat:WGS84坐标系的纬度
    :return:
    """
    if out_of_china(lng, lat):  # 判断是否在国内
        return [lng, lat]
    dlat = _transformlat(lng - 105.0, lat - 35.0)
    dlng = _transformlng(lng - 105.0, lat - 35.0)
    radlat = lat / 180.0 * pi
    magic = math.sin(radlat)
    magic = 1 - ee * magic * magic
    sqrtmagic = math.sqrt(magic)
    dlat = (dlat * 180.0) / ((a * (1 - ee)) / (magic * sqrtmagic) * pi)
    dlng = (dlng * 180.0) / (a / sqrtmagic * math.cos(radlat) * pi)
    mglat = lat + dlat
    mglng = lng + dlng
    return [mglng, mglat]


def gcj02_to_wgs84(lng, lat):
    """
    GCJ02(火星坐标系)转GPS84
    :param lng:火星坐标系的经度
    :param lat:火星坐标系纬度
    :return:
    """
    if out_of_china(lng, lat):
        return [lng, lat]
    dlat = _transformlat(lng - 105.0, lat - 35.0)
    dlng = _transformlng(lng - 105.0, lat - 35.0)
    radlat = lat / 180.0 * pi
    magic = math.sin(radlat)
    magic = 1 - ee * magic * magic
    sqrtmagic = math.sqrt(magic)
    dlat = (dlat * 180.0) / ((a * (1 - ee)) / (magic * sqrtmagic) * pi)
    dlng = (dlng * 180.0) / (a / sqrtmagic * math.cos(radlat) * pi)
    mglat = lat + dlat
    mglng = lng + dlng
    return [lng * 2 - mglng, lat * 2 - mglat]


def bd09_to_wgs84(bd_lon, bd_lat):
    lon, lat = bd09_to_gcj02(bd_lon, bd_lat)
    return gcj02_to_wgs84(lon, lat)


def wgs84_to_bd09(lon, lat):
    lon, lat = wgs84_to_gcj02(lon, lat)
    return gcj02_to_bd09(lon, lat)


def _transformlat(lng, lat):
    ret = -100.0 + 2.0 * lng + 3.0 * lat + 0.2 * lat * lat + \
          0.1 * lng * lat + 0.2 * math.sqrt(math.fabs(lng))
    ret += (20.0 * math.sin(6.0 * lng * pi) + 20.0 *
            math.sin(2.0 * lng * pi)) * 2.0 / 3.0
    ret += (20.0 * math.sin(lat * pi) + 40.0 *
            math.sin(lat / 3.0 * pi)) * 2.0 / 3.0
    ret += (160.0 * math.sin(lat / 12.0 * pi) + 320 *
            math.sin(lat * pi / 30.0)) * 2.0 / 3.0
    return ret


def _transformlng(lng, lat):
    ret = 300.0 + lng + 2.0 * lat + 0.1 * lng * lng + \
          0.1 * lng * lat + 0.1 * math.sqrt(math.fabs(lng))
    ret += (20.0 * math.sin(6.0 * lng * pi) + 20.0 *
            math.sin(2.0 * lng * pi)) * 2.0 / 3.0
    ret += (20.0 * math.sin(lng * pi) + 40.0 *
            math.sin(lng / 3.0 * pi)) * 2.0 / 3.0
    ret += (150.0 * math.sin(lng / 12.0 * pi) + 300.0 *
            math.sin(lng / 30.0 * pi)) * 2.0 / 3.0
    return ret


def out_of_china(lng, lat):
    """
    判断是否在国内，不在国内不做偏移
    :param lng:
    :param lat:
    :return:
    """
    return not (lng > 73.66 and lng < 135.05 and lat > 3.86 and lat < 53.55)


if __name__ == "__main__":
    print(bd09_to_wgs84(120.661362,31.415227))
#     res = [
#   {
#     "id": "BV11280861",
#     "location": "119.695947,30.202053",
#     "name": "九州街",
#     "sequence": "1"
#   },
#   {
#     "id": "BV11280866",
#     "location": "119.706441,30.227561",
#     "name": "临安广场",
#     "sequence": "2"
#   },
#   {
#     "id": "BV11280856",
#     "location": "119.738323,30.244543",
#     "name": "农林大学",
#     "sequence": "3"
#   },
#   {
#     "id": "BV11280860",
#     "location": "119.755663,30.255795",
#     "name": "青山湖",
#     "sequence": "4"
#   },
#   {
#     "id": "BV11280864",
#     "location": "119.776595,30.260992",
#     "name": "八百里",
#     "sequence": "5"
#   },
#   {
#     "id": "BV11280858",
#     "location": "119.825131,30.257022",
#     "name": "青山湖科技城",
#     "sequence": "6"
#   },
#   {
#     "id": "BV11280859",
#     "location": "119.878316,30.249779",
#     "name": "南峰",
#     "sequence": "7"
#   },
#   {
#     "id": "BV11280862",
#     "location": "119.899603,30.246777",
#     "name": "南湖",
#     "sequence": "8"
#   },
#   {
#     "id": "BV11280865",
#     "location": "119.929447,30.243033",
#     "name": "中泰",
#     "sequence": "9"
#   },
#   {
#     "id": "BV11280863",
#     "location": "119.955554,30.248755",
#     "name": "禹航路",
#     "sequence": "10"
#   },
#   {
#     "id": "BV11280857",
#     "location": "119.968843,30.260134",
#     "name": "凤新路",
#     "sequence": "11"
#   },
#   {
#     "id": "BV10780752",
#     "location": "119.992117,30.260136",
#     "name": "绿汀路",
#     "sequence": "12"
#   }
# ]
#     for item in res:
#         pos = item.get('location').split(',')
#         print(item.get('name'))
#         new_pos = gcj02_to_wgs84(float(pos[0]), float(pos[1]))
#         print('POINT ('+str(new_pos[0])+' '+str(new_pos[1])+')')

