#### 2024攀岩秋季赛成绩处理包含：成绩处理xlsx文件，包含宏的*.xlsm文件

#### WPS+VBA插件.rar 包含:插件安装程序和安装指南


 # 攀岩赛成绩处理裁判指南

恭喜你成为本次攀岩赛成绩处理的裁判,请阅读本攀岩成绩处理指南.

## 成绩处理任务:

 1.根据比赛要求处理出男子女子参赛表格

2.完善攀岩成绩处理系统xxxx年版本

3.比赛时对成绩进行记录和排序处理

4.成绩公布


## 任务1 根据比赛要求处理出参赛表格

直接对xlsx表格进行按时间顺序进行挑选即可,此处不做详细说明,记得去重


## 任务2.完善攀岩处理系统这个表格

<span style="color: red;">(注意尽量不要修改表格的名字和每个sheet结构,如有修改需要对函数需要做对应的调正) 

### 步骤1 观察表格

本"成绩处理系统"一共包含1个子sheet,本别为:

`1得分条`,`2女子记录表裁判`,`3男子记录表裁判`,`3.1男子成绩统计`,`2.1女子成绩统计`,`男子参赛表`,`女子参赛表`,`男子得分条`,`女子得分条`,`男子得分条备份`,`女子得分条备份` (共11个子sheet)

展示excle中表格样式,sheet表列表显示在下边
![image](https://github.com/Loueor/SUSTech-A-guide-to-dealing-with-climbing-results/blob/main/picture/%E5%B1%8F%E5%B9%95%E6%88%AA%E5%9B%BE%202024-10-04%20211642-1.png)
#### 注意到

在拿到的表格中
![image](https://github.com/Loueor/SUSTech-A-guide-to-dealing-with-climbing-results/blob/main/picture/%E5%B1%8F%E5%B9%95%E6%88%AA%E5%9B%BE%202024-09-22%20145206.png)
![image](https://github.com/Loueor/SUSTech-A-guide-to-dealing-with-climbing-results/blob/main/picture/%E5%B1%8F%E5%B9%95%E6%88%AA%E5%9B%BE%202024-10-04%20211912-1.png)

![image](https://github.com/Loueor/SUSTech-A-guide-to-dealing-with-climbing-results/blob/main/picture/%E5%B1%8F%E5%B9%95%E6%88%AA%E5%9B%BE%202024-10-04%20211806-1.png)

#### 这些时间信息需要修改!

#### 并且需要将处理出的男子女子参赛名单分别更新到该xlsx文件中的`"男子参赛表"`和`"女子参赛表"`这个子sheet中

### 步骤2 更改所有表中参赛号码和姓名(表格初始化)

#### 在微软excle中

可以使用`F11`+`Alt`键打开excle编程语言VBA, 如果是在笔记本中`F11`键可能和功能键重合,则需要按下`F11`+`Fn`+`Alt`(如果不能打开,请网络搜索"如何打开excle Microsoft Visual Basic for Application")

打开如下界面,点击任何一个sheet,打开编辑框,选择通用模式
![image](https://github.com/Loueor/SUSTech-A-guide-to-dealing-with-climbing-results/blob/main/picture/%E5%B1%8F%E5%B9%95%E6%88%AA%E5%9B%BE%202024-09-22%20150534.png)
![image](https://github.com/Loueor/SUSTech-A-guide-to-dealing-with-climbing-results/blob/main/picture/%E5%B1%8F%E5%B9%95%E6%88%AA%E5%9B%BE%202024-09-22%20150710.png)

#### 将下面代码复制放入后,点击左上角的绿色小三角    
![image](https://github.com/Loueor/SUSTech-A-guide-to-dealing-with-climbing-results/blob/main/picture/%E5%B1%8F%E5%B9%95%E6%88%AA%E5%9B%BE%202024-09-22%20150856.png)
#### 在WPS中
需要先下载`wps.vab`插件,同EXCLE操作.


该函数的功能是读取`男子参赛表`和`女子参赛表`的姓名列,分别填入`2女子记录表裁判`,`3男子记录表裁判`,`3.1男子成绩统计`,`2.1女子成绩统计`,`男子得分条`,`女子得分条`的姓名列,并且根据姓名的个数给出参赛号码,男子是101开头,女子是201开头.
```VBA
Sub CopyNamesAndGenerateID()
    Dim wsSource1 As Worksheet
    Dim wsSource2 As Worksheet
    Dim wsTarget1 As Worksheet
    Dim wsTarget2 As Worksheet
    Dim wsTarget3 As Worksheet
    Dim wsTarget4 As Worksheet
    Dim sourceLastRow As Long
    Dim targetStartRow1 As Long
    Dim targetStartRow2 As Long
    Dim targetStartRow3 As Long
    Dim targetStartRow4 As Long
    Dim targetStartRow5 As Long
    Dim targetStartRow6 As Long
    Dim i1 As Long
    Dim i2 As Long
    Dim idNumber1 As Long
    Dim idNumber2 As Long

    ' 设置工作表
    Set wsSource1 = ThisWorkbook.Sheets("男子参赛表") ' 来源工作表
    Set wsSource2 = ThisWorkbook.Sheets("女子参赛表") '
    Set wsTarget1 = ThisWorkbook.Sheets("3男子记录表裁判") ' 目标工作表
    Set wsTarget2 = ThisWorkbook.Sheets("2女子记录表裁判") ' 目标工作表
    Set wsTarget3 = ThisWorkbook.Sheets("3.1男子成绩统计") ' 目标工作表
    Set wsTarget4 = ThisWorkbook.Sheets("2.1女子成绩统计") ' 目标工作表
    Set wsTarget6 = ThisWorkbook.Sheets("女子得分条统计") ' 目标工作表
    Set wsTarget5 = ThisWorkbook.Sheets("男子得分条统计") ' 目标工作表



 ' 清空指定列（）从第四行开始的内容
    wsTarget1.Range("A4:G" & wsTarget1.Rows.Count).ClearContents
    wsTarget2.Range("A4:G" & wsTarget2.Rows.Count).ClearContents
    wsTarget3.Range("A4:I" & wsTarget3.Rows.Count).ClearContents
    wsTarget4.Range("A4:I" & wsTarget4.Rows.Count).ClearContents
    wsTarget5.Range("A3:P" & wsTarget5.Rows.Count).ClearContents
    wsTarget6.Range("A3:P" & wsTarget6.Rows.Count).ClearContents
    
    ' 找到男生女生参赛名单中的最后一行
    sourceLastRow1 = wsSource1.Cells(wsSource1.Rows.Count, "A").End(xlUp).Row
    sourceLastRow2 = wsSource2.Cells(wsSource2.Rows.Count, "A").End(xlUp).Row

    ' 设置目标工作表中开始复制名字的行号 (假设从第4行开始)
    targetStartRow1 = 4
    targetStartRow2 = 4


    ' 设置编号初始值
    idNumber1 = 101
    idNumber2 = 201
    xuhao1 = 1
    xuhao2 = 1
    
    '先解决男生的问题
    
    ' 将男子参赛表中的姓名列复制到SheetB中从第4行开始的某列（假设是C列）
    For i1 = 2 To sourceLastRow1 ' 假设男子参赛表的名字从第2行开始
    
        wsTarget1.Cells(targetStartRow1, 3).Value = wsSource1.Cells(i1, 1).Value ' 复制名字到3男子记录表裁判的C列
        wsTarget1.Cells(targetStartRow1, 2).Value = idNumber1 ' 在B列生成对应的编号
        wsTarget1.Cells(targetStartRow1, 1).Value = xuhao1 ' 在A列生成对应的序号
        
        wsTarget3.Cells(targetStartRow1, 3).Value = wsSource1.Cells(i1, 1).Value ' 复制名字到3男子记录表裁判的C列
        wsTarget3.Cells(targetStartRow1, 2).Value = idNumber1 ' 在B列生成对应的编号
        wsTarget3.Cells(targetStartRow1, 1).Value = xuhao1 ' 在A列生成对应的序号
        
        wsTarget5.Cells(targetStartRow1 - 1, 3).Value = wsSource1.Cells(i1, 1).Value ' 复制名字到3男子记录表裁判的C列
        wsTarget5.Cells(targetStartRow1 - 1, 2).Value = idNumber1 ' 在B列生成对应的编号
        wsTarget5.Cells(targetStartRow1 - 1, 1).Value = xuhao1 ' 在A列生成对应的序号
        
        targetStartRow1 = targetStartRow1 + 1
        idNumber1 = idNumber1 + 1
        xuhao1 = xuhao1 + 1
    Next i1
    
    '解决女生问题
    
     For i2 = 2 To sourceLastRow2 ' 假设男子参赛表的名字从第2行开始
    
        wsTarget2.Cells(targetStartRow2, 3).Value = wsSource2.Cells(i2, 1).Value ' 复制名字到3男子记录表裁判的C列
        wsTarget2.Cells(targetStartRow2, 2).Value = idNumber2 ' 在B列生成对应的编号
        wsTarget2.Cells(targetStartRow2, 1).Value = xuhao2 ' 在A列生成对应的序号
        
        wsTarget4.Cells(targetStartRow2, 3).Value = wsSource2.Cells(i2, 1).Value ' 复制名字到3男子记录表裁判的C列
        wsTarget4.Cells(targetStartRow2, 2).Value = idNumber2 ' 在B列生成对应的编号
        wsTarget4.Cells(targetStartRow2, 1).Value = xuhao2 ' 在A列生成对应的序号
        
        wsTarget6.Cells(targetStartRow2 - 1, 3).Value = wsSource2.Cells(i2, 1).Value ' 复制名字到3男子记录表裁判的C列
        wsTarget6.Cells(targetStartRow2 - 1, 2).Value = idNumber2 ' 在B列生成对应的编号
        wsTarget6.Cells(targetStartRow2 - 1, 1).Value = xuhao2 ' 在A列生成对应的序号
        
        
        targetStartRow2 = targetStartRow2 + 1
        idNumber2 = idNumber2 + 1
        xuhao2 = xuhao2 + 1
    Next i2

    MsgBox "姓名已成功复制并生成对应的编号。"
End Sub
```

#### 所有表格已经根据输入的男子女子参赛表成功初始化

### 步骤3 导出pdf

导出`1得分条`,`2女子记录表裁判`,`3男子记录表裁判`这三个表的pdf用于现场成绩统计,导出`男子成绩统计`和`女子成绩统计`当作签到表

## 任务三 比赛时和比赛后成绩处理

### 步骤1 完善 `男子成绩统计`和`女子成绩统计`这两个sheet

比赛中成绩处理裁判会拿到每个选手的成绩条,将成绩条成绩分别填入`男子成绩统计`和`女子成绩统计`这两个sheet.<span style="color: red;">注意:每条线路的完攀数>=1,奖励点数也>=1,不可能出现完攀>=1,奖励点数 =0 的情况.

<span style="color: red;">当场加入到的选手就依次在`男子参赛表`,`女子参赛表`,`男子成绩统计`和`女子成绩统计`,`男子得分条`,`女子得分条`这4个表格后添加名字和序号即可.

#### 因为程序编写,得分条的正确将直接影响最后成绩统计的正确性.

然后打开VBA运行下段程序得出男子女子成绩排名(选择跟上边的函数不同的VBA窗口).

该程序的功能是:通过`女子得分条`和`男子得分条`sheet,计算每位选手的得分情况,然后填入到`男子成绩统计`和`女子成绩统计`中:

    Top line是计算所有线路非零完攀数

    Z计算所有线路非零奖励点数

    Top times 是计算所有线路的完攀数之和

    Z times是计算所有线路非零奖励点数之和

排序功能是:

    先按照Top line的数量降序,如果Top line数量相同,按照Z 降序排列,如果Z的数量仍相同,则按照Top times的数量升序排列,如果Top times的数量还相同,就按照Z times升序排列

完赛标准:

    该程序设定的完赛标准是Z times >=1 即可

得奖情况:

    根据赛事要求自行排出1,2,3等奖.

<span style="color: red;">以上要求如有变化请自行更改函数
``` VBA

 Sub UpdateAndSortStatistics(ByVal sourceSheetName As String, ByVal targetSheetName As String, ByVal isMale As Boolean)
     Dim wsSource As Worksheet
     Dim wsTarget As Worksheet
     Dim lastRow As Long
     Dim i As Long, j As Long
     Dim topLineCount As Long, zCount As Long, topTimesSum As Long, zTimesSum As Long
     Dim rank As Long

     ' 设置工作表
     Set wsSource = ThisWorkbook.Sheets(sourceSheetName) ' 数据来源的表
     Set wsTarget = ThisWorkbook.Sheets(targetSheetName) ' 结果要写入的表

     ' 找到最后一行
     lastRow = wsTarget.Cells(wsTarget.Rows.Count, "A").End(xlUp).Row

     ' 遍历目标表中需要填写的行
     For i = 4 To lastRow ' 从第4行开始遍历（假设前3行为标题）
         topLineCount = 0
         zCount = 0
         topTimesSum = 0
         zTimesSum = 0

         ' 遍历来源表中的每条线路的奖励点和完攀情况（假设数据从D到O列）
         For j = 1 To 6 ' 六条线路
             If wsSource.Cells(i-1, 4 + (j - 1) * 2 + 1).Value > 0 Then ' 完攀列 D, F, H, J, L, N
                 topLineCount = topLineCount + 1 ' 完攀计数
                 topTimesSum = topTimesSum + wsSource.Cells(i-1, 4 + (j - 1) * 2 + 1).Value ' 完攀数值和
             End If
             If wsSource.Cells(i-1, 4 + (j - 1) * 2).Value > 0 Then ' 奖励点列 C, E, G, I, K, M
                 zCount = zCount + 1 ' 奖励点计数
                 zTimesSum = zTimesSum + wsSource.Cells(i-1, 4 + (j - 1) * 2).Value ' 奖励点数值和
             End If
         Next j

         ' 将结果写入目标表
         wsTarget.Cells(i, 4).Value = topLineCount ' 写入Top Line
         wsTarget.Cells(i, 5).Value = zCount ' 写入Z
         wsTarget.Cells(i, 6).Value = topTimesSum ' 写入TopTimes
         wsTarget.Cells(i, 7).Value = zTimesSum ' 写入Z Times
     Next i

     ' 排序设置
     With wsTarget.Sort
         .SortFields.Clear
         ' 按照 Top Line 排序 (降序)
         .SortFields.Add Key:=wsTarget.Range("D2:D" & lastRow), Order:=xlDescending
         ' 如果 Top Line 相同，按照 Z 排序 (降序)
         .SortFields.Add Key:=wsTarget.Range("E2:E" & lastRow), Order:=xlDescending
         ' 如果 Z 相同，按照 Top Times 排序 (升序)
         .SortFields.Add Key:=wsTarget.Range("F2:F" & lastRow), Order:=xlAscending
         ' 如果 Top Times 相同，按照 Z Times 排序 (升序)
         .SortFields.Add Key:=wsTarget.Range("G2:G" & lastRow), Order:=xlAscending

         ' 设置排序范围和方式，从第四行开始排序
         .SetRange wsTarget.Range("B4:G" & lastRow)
         .Header = xlNo ' 假设有两行标题
         .Apply
     End With

     ' 更新“是否完赛”列，根据 Z 列的值判断
     For i = 3 To lastRow
         If wsTarget.Cells(i, 5).Value >= 1 Then
             wsTarget.Cells(i, 8).Value = "完赛加分"
         Else
             wsTarget.Cells(i, 8).Value = "没有完赛"
         End If
     Next i

     ' 生成 Rank 列，从第4行开始
     rank = 1
     For i = 4 To lastRow
         wsTarget.Cells(i, 1).Value = rank ' 将 Rank 写入 A 列
         rank = rank + 1
     Next i

     ' 显示完成提示
     If isMale Then
         MsgBox "男生成绩统计和排序完成。"
     Else
         MsgBox "女生成绩统计和排序完成。"
     End If
 End Sub

 Sub UpdateMaleAndFemale()
    ' 更新男生成绩统计并排序
     UpdateAndSortStatistics "男子得分条统计", "3.1男子成绩统计", True

    ' 更新女生成绩统计并排序
     UpdateAndSortStatistics "女子得分条统计", "2.1女子成绩统计", False
 End Sub

```

## 任务四 成绩公布

在参赛群中公布`男子成绩统计`,`女子成绩统计`,`男子得分条`和`女子得分条`检验是否有失误.
