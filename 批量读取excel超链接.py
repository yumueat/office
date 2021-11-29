import xlrd
import requests
from concurrent.futures import ThreadPoolExecutor,ProcessPoolExecutor

def one_pic(url,i):
        resp=requests.get(url)
        img_name=f"./图片/{i}.jpeg"
        with open(img_name,"wb") as f:
            f.write(resp.content)
        print("over!")

def main():
    new_wb = xlrd.open_workbook("1_照片(1).xls", formatting_info=True)
    new_sh = new_wb.sheet_by_index(0)
    i=0
    with ThreadPoolExecutor(50) as t:
        while 1:
            try:
                new_sh.cell_value(i, 0)
                url1 = new_sh.hyperlink_map.get((i, 0))
                url = url1.url_or_path
                t.submit(one_pic, url, i)
                i += 1
            except IndexError:
                break
    print("all over!")

    pass

if __name__ == '__main__':
    main()





