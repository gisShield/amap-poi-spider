import logging
import traceback

from poiSpiderName import PoiSpiderName
from poiSpiderAddr import PoiSpiderAddr


LOG_FORMAT = "%(asctime)s %(name)s %(levelname)s %(pathname)s %(message)s "  # 配置输出日志格式
DATE_FORMAT = '%Y-%m-%d  %H:%M:%S %a '  # 配置输出时间的格式，注意月份和天数不要搞乱了
logging.basicConfig(level=logging.INFO,
                    format=LOG_FORMAT,
                    datefmt=DATE_FORMAT,
                    filename=r"./logs/poi-main.log"  # 有了filename参数就不会直接输出显示到控制台，而是直接写入文件
                    )

if __name__ == '__main__':
    print("1.请输入高德API key")
    key = input()
    if key is None or key == '':
        print("输入参数错误，输入任意字符退出...")
        input()
        exit(0)
    print("2.请选择抓取类型：（1：按名称抓取；2：按标准地址抓取）")
    spider_type = input()
    if spider_type == '1':
        try:
            print("3.请输入数据所在区划代码 如：330100")
            area_code = input()
            print("4.请输入POI类型")
            type_code = input()
            print("5.请选择输入类型(1:参数输入；2：文件输入)")
            input_type = input()
            if input_type == '1':
                print("6.请输入名称")
                name = input()
                print('请稍等......')
                poi_spider = PoiSpiderName(key, area_code, type_code, input_type, name=name)
                poi_spider.get_poi_name()
                print('数据抓取完成！')

            elif input_type == '2':
                print("6.请输入文件地址 如：c:\\1.xlsx")
                file_path = input()
                print("7.请输入工作表名称")
                sheet_name = input()
                print("8.请输入数据起始行行号")
                start_row = int(input())
                print("9.请输入名称列列号")
                name_col = input()
                print("10.请输入经纬度填写列号")
                coord_col = input()
                print('请稍等......')
                poi_spider = PoiSpiderName(key, area_code, type_code, input_type, file_path=file_path, sheet_name=sheet_name,
                                           start_row=start_row,
                                           name_col=name_col, coord_col=coord_col)
                poi_spider.get_poi_file()
                print('数据抓取完成！')
            else:
                logging.error("输入参数错误，程序退出...")
                print("输入参数错误，输入任意字符退出...")
                input()
                exit(0)
        except Exception as e:
            logging.error("程序异常====> %s", e.args)
            logging.error("异常信息====> %s", traceback.format_exc())
            print("程序出现异常，输入任意字符退出...")
            input()
            exit(0)
    elif spider_type == '2':
        try:
            print("3.请输入数据所在区划代码 如：330100")
            area_code = input()
            print("4.请选择输入类型(1:参数输入；2：文件输入)")
            input_type = input()
            if input_type == '1':
                print("5.请输入标准化地址如：浙江省杭州市滨江区星耀城一期")
                print("注：地址越标准，匹配到的POI信息越精准")
                addr = input()
                print('请稍等......')
                poi_spider = PoiSpiderAddr(key, area_code, input_type, addr=addr)
                poi_spider.get_poi_addr()
                print('数据抓取完成！')

            elif input_type == '2':
                print("5.请输入文件地址 如：c:\\1.xlsx")
                file_path = input()
                print("6.请输入工作表名称")
                sheet_name = input()
                print("7.请输入数据起始行行号")
                start_row = int(input())
                print("8.请输入地址列列号")
                addr_col = input()
                print("9.请输入经纬度填写列号")
                coord_col = input()
                print('请稍等......')
                poi_spider = PoiSpiderAddr(key, area_code, input_type, file_path=file_path, sheet_name=sheet_name,
                                           start_row=start_row,
                                           addr_col=addr_col, coord_col=coord_col)
                poi_spider.get_poi_file()
                print('数据抓取完成！')
            else:
                logging.error("输入参数错误，程序退出...")
                print("输入参数错误，输入任意字符退出...")
                input()
                exit(0)
        except Exception as e:
            logging.error("程序异常====> %s", e.args)
            logging.error("异常信息====> %s", traceback.format_exc())
            print("程序出现异常，输入任意字符退出...")
            input()
            exit(0)
    else:
        print("抓取类型错误，输入任意字符退出...")
        input()
        exit(0)
