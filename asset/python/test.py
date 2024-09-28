from datetime import date

def year_converter_to_wareki(date_obj):
    start_of_taisyo = date(1912, 7, 30) 
    start_of_syowa = date(1926, 12, 25) 
    start_of_heisei = date(1989, 1, 8) 
    start_of_reiwa = date(2019,5,1)
    year, month, day = None, None, None
    
    #-で分けてリスト化
    obj_list = date_obj.split("-")
    #2桁や1桁のときは19**にする
    if len(obj_list[0]) == 1:
        obj_list[0] = "190" + obj_list[0]
    elif len(obj_list[0]) == 2:
        obj_list[0] = "19" + obj_list[0]

    #年のみの入力は1月1日となる
    if len(obj_list) == 1:
        year = int(obj_list[0])
        input_obj = date(year, 1, 1)
    #年月の場合は1日となる
    elif len(obj_list) == 2:
        year, month = tuple(obj_list)
        year, month = int(year), int(month)
        input_obj = date(year, month, 1)
    elif len(obj_list) == 3:
        year, month, day = tuple(obj_list)
        year, month, day = int(year), int(month), int(day)
        input_obj = date(year, month, day)
    #リストに4つ以上の要素はエラー
    else:
        return None


    #出てきた年を格納する
    result_year = ""

    #年を格納
    if input_obj >=start_of_reiwa:
        nen = input_obj.year - start_of_reiwa.year + 1
        if nen == 1:
            nen = "元"
        result_year =  f"令和{nen}年"
    elif input_obj >= start_of_heisei:
        nen = input_obj.year - start_of_heisei.year + 1
        if nen == 1:
            nen = "元"
        result_year =  f"平成{nen}年"
    elif input_obj >= start_of_syowa:
        nen = input_obj.year - start_of_syowa.year + 1
        if nen == 1:
            nen = "元"
        result_year =  f"昭和{nen}年"
    elif input_obj >= start_of_taisyo:
        nen = input_obj.year - start_of_taisyo.year + 1
        if nen == 1:
            nen = "元"
        result_year =  f"大正{nen}年"
    #大正よりも前の場合
    else:
        return "昔過ぎて計算できません（大正以前）"

    #output（年月日を合算）
    result = ""

    #年のみの入力の場合は年のみで返す
    if not month:
        result = result_year
    else:
        #年月の場合
        if not day:
            result = f"{result_year}{month}月"
        #年月日の場合
        else:
            result = f"{result_year}{month}月{day}日"
    return result

#和暦から西暦へ
def year_converter_to_seireki(date_obj):
    taisyo = (1912, 7, 30)
    syowa = (1926, 12, 25)
    heisei = (1989, 1, 8)
    reiwa = (2019, 5, 1)

    year, month, day = None, None, None

#s68, h35, t50などをチェックする
    def error_checker(year_str):
        nengo, year_int = year_str[0], int(year_str[1:])

    def return_year(year_str):
        nengo, year_int = year_str[0], int(year_str[1:])
        if nengo == "r":
            return 2019 + year_int - 1
        elif nengo == "h":
            return 1989 + year_int - 1
        elif nengo == "s":
            return 1926 + year_int -1
        else:
            return 1912 + year_int -1
        

    obj_list = date_obj.split("-")
    if len(obj_list[0]) >= 4:
        return None
    if len(obj_list) > 1:
        if int(obj_list[1]) > 12:
            return None
    if len(obj_list) > 2:
        if int(obj_list[2]) > 31:
            return None
    #年のみの入力
    if len(obj_list) == 1:
        year = obj_list[0]

    #年月の場合
    elif len(obj_list) == 2:
        year, month = tuple(obj_list)

    #年月日
    elif len(obj_list) == 3:
        year, month, day = tuple(obj_list)
    else:
        return None

    result_year = f"{return_year(year)}年"

        #年のみの入力の場合は年のみで返す
    if not month:
        result = result_year
    else:
        #年月の場合
        if not day:
            result = f"{result_year}{month}月"
        #年月日の場合
        else:
            result = f"{result_year}{month}月{day}日"
    return result

if __name__ == "__main__":
    data = input("西暦・和暦のどちらかをご入力ください→ ")
    try:
        i = int(data[0])
        result = year_converter_to_wareki(data)
    #そうでないとき
    except:
        #t, s, h, rから始まるとき
        if data[0] in ["t", "s", "h", "r"]:
            result = year_converter_to_seireki(data)
        #そうでないとき
        else:
            print("正しい値を入力してください")
    #resultに値が返っているとき
    if result:
        print(result)
    else:
        print("正しい値を入力してください")