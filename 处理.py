import xlrd

Book = xlrd.open_workbook("2018招生计划上册完整版.xlsx")
CollegeAndMajor = {}
ThisCollege = ""
for i in range(18, 238):
    ThisSheet = Book.sheet_by_name("Table " + str(i))
    #获取sheet索引
    RowsNumber = ThisSheet.nrows
    #获取该sheet所有行
    for Index in range(RowsNumber):
        ThisRow = ThisSheet.row_values(Index)
        if type(ThisRow[0]) == float:
            if ThisRow[0] >= 100:
                if ThisRow[0] < 1000:
                    Number = "0" + str(int(ThisRow[0]))
                else:
                    Number = str(int(ThisRow[0]))
                try:
                    ThisCollege = (Number, ThisRow[1], int(ThisRow[2]))
                except ValueError:
                    ThisCollege = (Number, ThisRow[1], ThisRow[2])
                CollegeAndMajor[ThisCollege] = []
            if ThisRow[0] < 100:
                if ThisRow[0] < 10:
                    Number = "0" + str(int(ThisRow[0]))
                else:
                    Number = str(int(ThisRow[0]))
                try:
                    if ThisRow[5] == "√":
                        ZhengZhi = True
                    else:
                        ZhengZhi = False
                except IndexError:
                    ZhengZhi = False
                try:
                    if ThisRow[6] == "√":
                        DiLi = True
                    else:
                        DiLi = False
                except IndexError:
                    DiLi = False
                try:
                    if ThisRow[7] == "√":
                        HuaXue = True
                    else:
                        HuaXue = False
                except IndexError:
                    HuaXue = False
                try:
                    if ThisRow[8] == "√":
                        ShengWu = True
                    else:
                        ShengWu = False
                except IndexError:
                    ShengWu = False
                try:
                    CollegeAndMajor[ThisCollege].append((ThisRow[0], ThisRow[1], int(ThisRow[2]), ThisRow[3],
                                                         ThisRow[4], ZhengZhi, DiLi, HuaXue, ShengWu))
                except KeyError:
                    pass
                except ValueError:
                    CollegeAndMajor[ThisCollege].append((ThisRow[0], ThisRow[1], ThisRow[2], ThisRow[3],
                                                         ThisRow[4], ZhengZhi, DiLi, HuaXue, ShengWu))
File = open("理科二本（2）.csv", "w")
File.write("代号,院校名称,计划数,选修要求,选测科目\n")
# File.write("代号,名称,计划数,学制,学费,不支持的选测科目\n")
while True:
    ThisNumber = input("输入学校代号：")
    for ThisCollege in CollegeAndMajor:
        if ThisCollege[0] == ThisNumber:
            try:
                NameOfCollege = ThisCollege[1][: ThisCollege[1].index("\n")]
                GetIndex = ThisCollege[1].index("\n") + 1
                print(NameOfCollege)
            except ValueError:
                print("院校名称出错。")
                NameOfCollege = input("请重新输入：")
            NewStudents = input("请输入计划数：")
            try:
                print("该学校的选测科目要求是：" + ThisCollege[1][ThisCollege[1].index("求") + 2: ThisCollege[1].index("\n", GetIndex)])
            except Exception:
                print("未找到，请手动输入选测科目要求：")
            Quest = input("请修改：")
            if Quest:
                File.write("{},{},{},{}\n".format(ThisNumber, NameOfCollege, NewStudents, Quest))
            else:
                File.write("{},{},{},{}\n".format(ThisNumber, NameOfCollege, NewStudents,
                                                  ThisCollege[1][ThisCollege[1].index("求") + 2: ThisCollege[1].index("\n", GetIndex)]))
            break
    else:
        print("无" + ThisNumber)
        NameOfCollege = input("请输入院校名称：")
        NewStudents = input("请输入计划数：")
        Quest = input("请输入选测要求：")
        File.write("{},{},{},{}\n".format(ThisNumber, NameOfCollege, NewStudents, Quest))
    if ThisNumber == "":
        break
File.close()
