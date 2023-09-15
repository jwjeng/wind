'檔名：風速資料轉換
'用途：將風速計取得的原始數據加以整理，計算平均及累計值，再以隨機檔儲存以供日後運算

Imports System.IO                                   '要加這句才能執行許多預設指令

Public Class Form1
    Public Dir, FName, SName As String                     ' 宣告公用變數
    Public DataType As Single

    Structure WindData
        <VBFixedString(21)> Public DatTim As String         ' 字串要加前面那個<VBFixedString()>指定長度
        Public SNumber As Integer
        Public Speed, Direct, Avg, Var, VarX, VarY, DirAvg As Double
        Public DirMode As Integer
        Public ResSpeed, ResDir, Persistence, PosX, PosY, Dist As Double
        Public Temperature, RelHumidity, AtmPressure As Double
    End Structure

    Public Wind As WindData
    Public DataLen As Short = Len(Wind)
    Const Radians As Double = Math.PI / 180                 'Radians=弳度=PI/180
    Public TimeFreq(24, 2)                                  '設定每小時分布1數量、2換算頻度



    '  表單初始設定
    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Timer1.Interval = 1000
        Timer1.Enabled = False

        '  TabPage1-原始資料試算
        CheckBox4.Checked = True

        '  TabPage2-從隨機檔匯出資料
        TextBox7.Text = 0 : TextBox8.Text = 23                  '預先設定篩選時間為0-23時
        CheckedListBox1.SetItemChecked(0, True)                 '預先設定勾選日期、原始風速、風向
        CheckedListBox1.SetItemChecked(1, True)
        CheckedListBox1.SetItemChecked(2, True)
        RadioButton1.Checked = True
        CheckBox3.Checked = False

        '  TabPage4-基本資料試算
        RadioButton10.Checked = True
        RadioButton12.Checked = True
        RadioButton13.Checked = True
        RadioButton14.Checked = True


        '加入這段可隱藏特定標籤頁
        'For Each tabpg As TabPage In TabControl1.TabPages
        'If tabpg.Text.StartsWith("落點") Then TabControl1.TabPages.Remove(tabpg) '以標籤內文Text搜尋
        'If tabpg.Name.StartsWith("TabPage3") Then TabControl1.TabPages.Remove(tabpg) '以標籤名稱Name搜尋  
        'Next
        '語法結束

    End Sub

    ' ===== 程式結束 =====
    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click, Button5.Click, Button8.Click, Button14.Click, Button18.Click
        End
    End Sub

    ' ===== 從原始資料進行試算轉檔 =====

    ' 從原始資料進行試算轉檔-設定輸出入檔名
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim I, N As Integer
        Dim S1, S2 As String
        If OpenFileDialog1.ShowDialog() = Windows.Forms.DialogResult.OK Then            '啟用檔案對話框，輸入檔名FName
            FName = OpenFileDialog1.FileName
            N = Len(FName)
            For I = N To 1 Step -1                                                      '將檔名分割為路徑Dir和檔名FName
                If Mid(FName, I, 1) = "\" Then
                    Dir = Microsoft.VisualBasic.Left(FName, I - 1)
                    FName = Mid(FName, I + 1)
                End If
            Next
            Label1.Text = "目錄：" & Dir
            Label2.Text = "檔名：" & FName
            TextBox62.Text = Dir                                    '預設輸出目錄
            TextBox1.Text = FName + "a"                             '預設輸出檔名
            TextBox2.Text = "記錄長度：" & DataLen & vbCrLf
            TextBox3.Text = 60
            TextBox4.Text = 8
            TextBox19.Text = -1                 ' 造成風向變化最低風速，以下忽略無影響

            '判斷資料類型
            FileOpen(1, Dir + "\" + FName, OpenMode.Input)
            If CheckBox4.Checked = True Then
                For I = 1 To 4                      ' 檔頭資料不需要
                    S1 = LineInput(1)
                Next
            End If
            S1 = LineInput(1)
            S2 = LineInput(1)
            If S1.Length > 45 Then ST2.Checked = True Else ST1.Checked = True
            TextBox2.Text = "最初兩筆資料：" & vbCrLf & S1 & vbCrLf & S2
            If Microsoft.VisualBasic.Mid(S1, 19, 2) <> "00" Or Microsoft.VisualBasic.Mid(S2, 19, 2) <> "00" Then
                RB1.Checked = True
            Else
                If Microsoft.VisualBasic.Mid(S1, 17, 4) <> "0:00" Or Microsoft.VisualBasic.Mid(S2, 17, 4) <> "0:00" Then
                    RB1min.Checked = True
                Else
                    If Microsoft.VisualBasic.Mid(S1, 16, 5) <> "00:00" Or Microsoft.VisualBasic.Mid(S2, 16, 5) <> "00:00" Then
                        RB10m.Checked = True
                    Else
                        RB1hr.Checked = True
                    End If
                End If
            End If
            FileClose(1)

        End If
    End Sub

    '讀取原始檔
    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Dim Dir1, S As String
        Dim I, N, SN, DataPos, AvgNo, SumNo, AvgPos, SumPos As Integer
        Dim LastDir, MinSP, SX, SX2, XV, YV, XVector, YVector, XV2, YV2, RX, RX1, RY, RY1, DirX, DirY, TempX, RHX, APX As Double

        AvgNo = Val(TextBox3.Text)
        SumNo = Val(TextBox4.Text)
        LastDir = 0
        TextBox18.Text = ""
        MinSP = Val(TextBox19.Text)         ' MinSP = 造成風向變化最低風速
        Dim AvgSpeed(AvgNo), AvgDir(AvgNo) As Double                        '暫存風速風向
        Dim ResX(AvgNo), ResY(AvgNo) As Double                              '合成風向
        Dim AvgTemp(AvgNo), AvgRH(AvgNo), AvgAP(AvgNo) As Double            '暫存溫度、濕度、氣壓
        Dim SumSpeed(SumNo), SumDir(SumNo), Dx(SumNo), Dy(SumNo) As Double  '風速風向總和
        Dim Dir16(15) As Integer

        Dir1 = TextBox62.Text
        SName = TextBox1.Text
        FileOpen(1, Dir + "\" + FName, OpenMode.Input)
        FileOpen(2, Dir1 + "\" + SName, OpenMode.Random, , , DataLen)

        If CheckBox4.Checked = True Then
            For I = 1 To 4                      ' 檔頭資料不需要
                S = LineInput(1)
            Next
        End If

        DataPos = 0
        Wind.Temperature = 0
        Wind.RelHumidity = 0
        Wind.AtmPressure = 0
        Do Until EOF(1)
            Application.DoEvents()          ' 加這個才可以在中途印出txtbox
            DataPos += 1
            Input(1, Wind.DatTim)           ' 原始資料四個欄位:時間、序號、風速、風向
            If CheckBox4.Checked = True Then Input(1, Wind.SNumber)
            Input(1, Wind.Speed)
            Input(1, Wind.Direct)
            If Wind.Speed <= MinSP Then     ' 風速小於 MinSP 視為風向不變 
                Wind.Direct = LastDir
            Else
                LastDir = Wind.Direct
            End If
            If Wind.Direct < 0 Then Wind.Direct += 360
            If ST2.Checked = True Then      ' ST2再讀取三欄位：溫度、濕度、氣壓
                Input(1, Wind.Temperature)
                Input(1, Wind.RelHumidity)
                Input(1, Wind.AtmPressure)
            End If
            ' 篩選漏行
            If CheckBox4.Checked = True Then
                If DataPos > 1 And Wind.SNumber <> SN + 1 Then TextBox18.Text += "第" & DataPos + 4 & "行：" & Wind.DatTim & vbCrLf
                SN = Wind.SNumber
            End If

            AvgPos += 1                                                     ' 暫存平均位置指標
            If AvgPos > AvgNo Then AvgPos = 1
            SX = SX + Wind.Speed                                            ' 計算風速平均及變方
            SX2 = SX2 + Wind.Speed ^ 2                                      ' SX: 累計風速, SX2=風速平方和

            XV = Math.Cos(Wind.Direct * Radians)                            ' 計算風向平均及變方
            YV = Math.Sin(Wind.Direct * Radians)                            ' XV, YV: 風向分向量
            XVector += XV : YVector += YV                                   ' X,YVector: 風向分向量累計
            XV2 += XV ^ 2 : YV2 += YV ^ 2                                   ' XV2,YV2: 風向分向量平方和

            Dir16(D16(Wind.Direct)) += 1                                    ' 計算盛行風向(眾數)

            RX1 = XV * Wind.Speed : RX += RX1                               ' 計算合成風速風向
            RY1 = YV * Wind.Speed : RY += RY1

            If ST2.Checked = True Then                                      ' ST2：溫度、濕度、氣壓 累加
                TempX += Wind.Temperature
                RHX += Wind.RelHumidity
                APX += Wind.AtmPressure
            End If

            If DataPos <= AvgNo Then                                        '  未達基本量
                N = DataPos
            Else
                SX -= AvgSpeed(AvgPos)                                      '  累計值減去被踢出累計範圍的部分
                SX2 -= AvgSpeed(AvgPos) ^ 2
                XVector -= Math.Cos(AvgDir(AvgPos) * Radians)
                YVector -= Math.Sin(AvgDir(AvgPos) * Radians)
                XV2 -= Math.Cos(AvgDir(AvgPos) * Radians) ^ 2
                YV2 -= Math.Sin(AvgDir(AvgPos) * Radians) ^ 2
                Dir16(D16(AvgDir(AvgPos))) -= 1
                RX -= ResX(AvgPos) : RY -= ResY(AvgPos)
                If ST2.Checked = True Then                                  ' ST2：溫度、濕度、氣壓 扣回N次前資料
                    TempX -= AvgTemp(AvgPos)
                    RHX -= AvgRH(AvgPos)
                    APX -= AvgAP(AvgPos)
                End If
            End If

            Wind.Avg = SX / N : Wind.Var = (SX2 - SX ^ 2 / N) / N                   ' 得到風速平均及變方
            AvgSpeed(AvgPos) = Wind.Speed
            AvgDir(AvgPos) = Wind.Direct
            Wind.DirAvg = Math.Atan(YVector / XVector) / Radians                    ' 得到風向平均
            If XVector < 0 Then Wind.DirAvg += 180
            If Wind.DirAvg < 0 Then Wind.DirAvg += 360
            Wind.VarX = (XV2 - XVector ^ 2 / N) / N                                 ' 得到風向X,Y分向量變方
            Wind.VarY = (YV2 - YVector ^ 2 / N) / N

            'For I = 0 To 15                                                        ' 得到盛行風向(眾數)
            'If Dir16(Wind.DirMode) < Dir16(I) Then Wind.DirMode = I
            'Next
            If Dir16(Wind.DirMode) < Dir16(D16(Wind.Direct)) Then Wind.DirMode = D16(Wind.Direct)

            ResX(AvgPos) = RX1 : ResY(AvgPos) = RY1                                 ' 得到合成風速風向
            Wind.ResSpeed = Math.Sqrt(RX ^ 2 + RY ^ 2) / AvgNo
            Wind.ResDir = Math.Atan(RY / RX) / Radians
            If RX < 0 Then Wind.ResDir += 180
            If Wind.ResDir < 0 Then Wind.ResDir += 360

            Wind.Persistence = Wind.ResSpeed / Wind.Avg                             ' 得到持續性Persistence

            If ST2.Checked = True Then
                AvgTemp(AvgPos) = Wind.Temperature          ' ST2：溫度、濕度、氣壓 暫存
                AvgRH(AvgPos) = Wind.RelHumidity
                AvgAP(AvgPos) = Wind.AtmPressure
                Wind.Temperature = TempX / N                ' ST2：溫度、濕度、氣壓 平均
                Wind.RelHumidity = RHX / N
                Wind.AtmPressure = APX / N
            End If

            If DataType < 2 Then                                                    ' 0.1秒,1秒
                SumPos += 1                                                         ' 暫存累計位置指標
                If SumPos > SumNo Then SumPos = 1
                DirX = Math.Cos(Wind.Direct * Radians) * Wind.Speed : Wind.PosX += DirX * DataType
                DirY = Math.Sin(Wind.Direct * Radians) * Wind.Speed : Wind.PosY += DirY * DataType
                If DataPos > SumNo Then Wind.PosX -= Dx(SumPos) : Wind.PosY -= Dy(SumPos) ' 得到累計位置 PosX,Y
                Dx(SumPos) = DirX * DataType : Dy(SumPos) = DirY * DataType
            Else                                                                    ' 1分、10分、1時
                Wind.PosX = Math.Cos(Wind.Direct * Radians) * Wind.Speed * SumNo
                Wind.PosY = Math.Sin(Wind.Direct * Radians) * Wind.Speed * SumNo
            End If

            FilePut(2, Wind, DataPos)                                   '存入隨機檔
            If DataPos > SumNo Then
                DirX = Wind.PosX : DirY = Wind.PosY
                FileGet(2, Wind, DataPos - SumNo + 1)
                Wind.PosX = DirX : Wind.PosY = DirY
                Wind.Dist = Math.Sqrt(DirX ^ 2 + DirY ^ 2)
                FilePut(2, Wind, DataPos - SumNo + 1)
            End If

            If DataPos Mod 1000 = 0 Then                                '每1000筆印出一次
                TextBox2.Text = "記錄：" & DataPos & vbCrLf
                TextBox2.Text += "時間：" & Wind.DatTim & vbCrLf
                TextBox2.Text += "風速：" & Wind.Speed & vbCrLf
                TextBox2.Text += "風向：" & Wind.Direct & vbCrLf
                TextBox2.Text += "落點：X = " & Wind.PosX & vbCrLf & "　　　Y = " & Wind.PosY & vbCrLf
                TextBox2.Text += "距離：" & Wind.Dist & vbCrLf
                TextBox2.Text += "溫度：" & Wind.Temperature & vbCrLf
                TextBox2.Text += "濕度：" & Wind.RelHumidity & vbCrLf
                TextBox2.Text += "氣壓：" & Wind.AtmPressure & vbCrLf
            End If
        Loop
        TextBox2.Text += "轉檔結束"
        If TextBox18.Text = "" Then
            TextBox18.Text = "一切ok，沒有缺漏"
        Else
            TextBox18.Text = "中間有問題，請檢查!" & vbCrLf & vbCrLf & TextBox18.Text
        End If
        TextBox2.Text += vbCrLf & "轉檔完成! 資料筆數：" & DataPos
        FileClose(1) : FileClose(2)
    End Sub

    '函數：轉16方位角
    Function D16(ByVal WindDir As Single)
        D16 = CShort(WindDir / 22.5)
        If D16 = 16 Then D16 = 0
    End Function

    '資料類別：0.1秒、1秒、1分、10分、1時
    Private Sub RB01_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RB01.CheckedChanged, ST1.CheckedChanged
        DataType = 0.1
    End Sub

    Private Sub RB1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RB1.CheckedChanged, ST2.CheckedChanged
        DataType = 1
    End Sub

    Private Sub RB1min_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RB1min.CheckedChanged
        DataType = 2
    End Sub

    Private Sub RB10m_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RB1min.CheckedChanged
        DataType = 3
    End Sub

    Private Sub RB1hr_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RB1hr.CheckedChanged, RB10m.CheckedChanged
        DataType = 4
    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        TextBox2.Text = "記錄：" & Wind.SNumber & vbCrLf
        TextBox2.Text += "時間：" & Wind.DatTim & vbCrLf
        TextBox2.Text += "風速：" & Wind.Speed & vbCrLf
        TextBox2.Text += "風向：" & Wind.Direct & vbCrLf
        TextBox2.Text += "落點：X = " & Wind.PosX & "Y = " & Wind.PosY
    End Sub


    ' ===== 從彙整後的隨機檔匯出資料 =====

    ' 從彙整後的隨機檔匯出資料-設定資料檔來源及輸出檔名
    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        Dim I, N As Integer
        ' TextBox49.Text = 1 : TextBox61.Text = 1

        If OpenFileDialog1.ShowDialog() = Windows.Forms.DialogResult.OK Then            '啟用檔案對話框，輸入檔名FName
            FName = OpenFileDialog1.FileName
            N = Len(FName)
            For I = N To 1 Step -1                                                      '將檔名分割為路徑Dir和檔名FName
                If Mid(FName, I, 1) = "\" Then
                    Dir = Microsoft.VisualBasic.Left(FName, I - 1)
                    FName = Mid(FName, I + 1)
                End If
            Next
            Label7.Text = "目錄：" & Dir
            Label8.Text = "檔名：" & FName
            TextBox63.Text = Dir
            TextBox5.Text = FName + "_out.txt"
            TextBox6.Text = "" : TextBox20.Text = ""

            Dim Min, Max As Integer                                 ' 設定輸出資料區間
            FileOpen(1, Dir + "\" + FName, OpenMode.Random, , , DataLen)
            Min = 1 : Max = LOF(1) \ DataLen                        '預先設定資料範圍為全部
            FileClose(1)
            TrackBar1.Minimum = 1 : TrackBar1.Maximum = Max
            TrackBar1.TickFrequency = Max \ 20 : TrackBar1.LargeChange = Max \ 200
            TrackBar2.Minimum = 1 : TrackBar2.Maximum = Max
            TrackBar2.TickFrequency = Max \ 20 : TrackBar2.LargeChange = Max \ 200
            TrackBar1.Value = 1 : TrackBar2.Value = Max
            TextBox49.Text = TrackBar1.Value : TextBox61.Text = TrackBar2.Value
            TextBox17.Text = "資料數量：" & Max & "筆"

            'FileGet(1, Wind, TrackBar1.Value)
            'Label11.Text = "起始時間：" & Wind.DatTim
            'FileGet(1, Wind, TrackBar2.Value)
            'Label10.Text = "終止時間：" & Wind.DatTim
            'FileClose(1)

        End If
    End Sub

    ' 從彙整後的隨機檔匯出資料-設定起始和終止時間～不需輸出全區間
    Private Sub TrackBar1_Scroll(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TrackBar1.Scroll
        FileOpen(1, Dir + "\" + FName, OpenMode.Random, , , DataLen)
        FileGet(1, Wind, TrackBar1.Value)
        Label11.Text = "起始時間：" & Wind.DatTim
        FileClose(1)
        TextBox49.Text = TrackBar1.Value
    End Sub

    Private Sub TrackBar2_Scroll(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TrackBar2.Scroll
        FileOpen(1, Dir + "\" + FName, OpenMode.Random, , , DataLen)
        FileGet(1, Wind, TrackBar2.Value)
        Label10.Text = "終止時間：" & Wind.DatTim
        FileClose(1)
        TextBox61.Text = TrackBar2.Value
    End Sub

    Private Sub TextBox49_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox49.TextChanged, TextBox61.TextChanged
        If Val(TextBox49.Text) * Val(TextBox61.Text) <> 0 Then
            TrackBar1.Value = Val(TextBox49.Text)
            TrackBar2.Value = Val(TextBox61.Text)
            FileOpen(1, Dir + "\" + FName, OpenMode.Random, , , DataLen)
            FileGet(1, Wind, TrackBar1.Value)
            Label11.Text = "起始時間：" & Wind.DatTim
            FileGet(1, Wind, TrackBar2.Value)
            Label10.Text = "終止時間：" & Wind.DatTim
            FileClose(1)
        End If
    End Sub

    ' 將溫濕度氣壓資料併入隨機檔匯出資料-設定資料檔來源及輸出檔名
    Private Sub Button10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button10.Click
        Dim FName1, S1, S2 As String
        If OpenFileDialog1.ShowDialog() = Windows.Forms.DialogResult.OK Then            '啟用檔案對話框，輸入檔名FName1
            FName1 = OpenFileDialog1.FileName
            TextBox20.Text = FName1
            '顯示最初兩筆資料
            FileOpen(1, FName1, OpenMode.Input)
            For I = 1 To 4                      ' 檔頭資料不需要
                S1 = LineInput(1)
            Next
            S1 = LineInput(1)
            S2 = LineInput(1)
            TextBox17.Text = "最初2筆資料：" & vbCrLf & S1 & vbCrLf & S2 & vbCrLf
            FileClose(1)
        End If
    End Sub

    ' 從彙整後的隨機檔匯出資料-開始轉檔
    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        Dim F, F1, I, J, N As Integer
        Dim W1, W2, W3 As Double
        Dim Temp_A As String = ""
        Dim S As String = ""
        Dim Dir1 As String = ""
        SName = TextBox5.Text
        Dir1 = TextBox63.Text
        FileOpen(1, Dir + "\" + FName, OpenMode.Random, , , DataLen)
        FileOpen(2, Dir1 + "\" + SName, OpenMode.Output)
        If Microsoft.VisualBasic.Left(TextBox20.Text, 2) <> "" Then
            FileOpen(3, TextBox20.Text, OpenMode.Input)
            For I = 1 To 4                      ' 檔頭資料不需要
                S = LineInput(3)
            Next
            F1 = 1
        End If
        For I = TrackBar1.Value To TrackBar2.Value                          ' 隨機檔可以直接從特定位置讀資料
            Application.DoEvents()                                          ' 加這個才可以在中途印出txtbox
            FileGet(1, Wind, I) : F = 0
            J = Val(Microsoft.VisualBasic.Mid(Wind.DatTim, 12, 2))
            '依選取的頻率輸出
            If RadioButton1.Checked = True Or
              (RadioButton2.Checked = True And Microsoft.VisualBasic.Mid(Wind.DatTim, 20, 2) = "  ") Or
              (RadioButton3.Checked = True And Microsoft.VisualBasic.Mid(Wind.DatTim, 18, 2) = "00") Or
              (RadioButton24.Checked = True And Microsoft.VisualBasic.Mid(Wind.DatTim, 16, 4) = "0:00") Or
              (RadioButton4.Checked = True And Microsoft.VisualBasic.Mid(Wind.DatTim, 15, 5) = "00:00") Then F = 1
            If F1 = 1 And F = 1 Then
                F1 = 0
                Do Until F1 = 1 Or EOF(3)
                    If CheckBox3.Checked = True Then
                        S = LineInput(3)    '要表頭，整行讀入
                        If Microsoft.VisualBasic.Mid(S, 2, 19) = Microsoft.VisualBasic.Mid(Wind.DatTim, 1, 19) Then F1 = 1
                    Else
                        Input(3, S)        '不用表頭，依序讀入各欄位資料
                        Input(3, W1) : Input(3, W1) : Input(3, W1)
                        Input(3, W1)
                        Input(3, W2)
                        Input(3, W3)
                        If S = Microsoft.VisualBasic.Mid(Wind.DatTim, 1, 19) Then F1 = 1
                    End If
                Loop
                F1 = 1
            End If

            '看哪個有勾才輸出哪個
            If F = 1 And J >= Val(TextBox7.Text) And J <= Val(TextBox8.Text) Then
                If CheckedListBox1.GetItemChecked(0) Then Temp_A = Chr(34) & Wind.DatTim & Chr(34)
                If CheckedListBox1.GetItemChecked(1) Then Temp_A &= " , " & Wind.Speed
                If CheckedListBox1.GetItemChecked(2) Then Temp_A &= " , " & Wind.Direct
                If CheckedListBox1.GetItemChecked(3) Then Temp_A &= " , " & Wind.Avg
                If CheckedListBox1.GetItemChecked(4) Then Temp_A &= " , " & Wind.Var
                If CheckedListBox1.GetItemChecked(5) Then Temp_A &= " , " & Wind.DirAvg
                If CheckedListBox1.GetItemChecked(6) Then Temp_A &= " , " & Wind.VarX
                If CheckedListBox1.GetItemChecked(7) Then Temp_A &= " , " & Wind.VarY
                If CheckedListBox1.GetItemChecked(8) Then Temp_A &= " , " & Wind.DirMode
                If CheckedListBox1.GetItemChecked(9) Then Temp_A &= " , " & Wind.ResSpeed
                If CheckedListBox1.GetItemChecked(10) Then Temp_A &= " , " & Wind.ResDir
                If CheckedListBox1.GetItemChecked(11) Then Temp_A &= " , " & Wind.Persistence
                If CheckedListBox1.GetItemChecked(12) Then Temp_A &= " , " & Wind.PosX
                If CheckedListBox1.GetItemChecked(13) Then Temp_A &= " , " & Wind.PosY
                If CheckedListBox1.GetItemChecked(14) Then Temp_A &= " , " & Wind.Dist
                If CheckedListBox1.GetItemChecked(15) Then Temp_A &= " , " & Wind.Temperature
                If CheckedListBox1.GetItemChecked(16) Then Temp_A &= " , " & Wind.RelHumidity
                If CheckedListBox1.GetItemChecked(17) Then Temp_A &= " , " & Wind.AtmPressure
                If F1 = 1 Then
                    If CheckBox3.Checked = 1 Then
                        Temp_A &= " , " & S                             '要表頭，整行加入
                    Else
                        Temp_A &= " , " & W1 & " , " & W2 & " , " & W3  '不用表頭，加入後三項資料
                    End If
                End If
                PrintLine(2, Temp_A) : N += 1
                If F = 1 Then
                    If (RadioButton34.Checked And RadioButton3.Checked) Or (RadioButton36.Checked And RadioButton4.Checked) = True Then I += 59
                    If RadioButton34.Checked And RadioButton4.Checked = True Then I += 3599
                    If RadioButton34.Checked And RadioButton24.Checked = True Then I += 599
                    If RadioButton36.Checked And RadioButton24.Checked = True Then I += 9
                End If

            End If
            If (I Mod 1000 = 0) Or (N Mod 1000 = 0) Then TextBox6.Text = "記錄：" & I & "  輸出資料筆數：" & N : TextBox17.Text = Temp_A
        Next
        FileClose(1)
        FileClose(2)
        If F1 = 1 Then FileClose(3)
        TextBox6.Text = "轉檔完成! 輸出資料筆數：" & N
    End Sub

    '=====  落點頻度分布計算  =====

    '設定輸出入檔名
    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        Dim I, N, Total, NumX, NumY As Integer
        Dim MinX, MinY, MinX0, MinY0, MaxX, MaxY, X, Y, StepX, StepY As Decimal
        Dim SX, SX2, SY, SY2 As Double
        If OpenFileDialog1.ShowDialog() = Windows.Forms.DialogResult.OK Then            '啟用檔案對話框，輸入檔名FName
            FName = OpenFileDialog1.FileName
            N = Len(FName)
            For I = N To 1 Step -1                                                      '將檔名分割為路徑Dir和檔名FName
                If Mid(FName, I, 1) = "\" Then
                    Dir = Microsoft.VisualBasic.Left(FName, I - 1)
                    FName = Mid(FName, I + 1)
                End If
            Next
            Label14.Text = "目錄：" & Dir
            Label15.Text = "檔名：" & FName
            TextBox64.Text = Dir
            TextBox9.Text = FName + "_freq.txt"
            Application.DoEvents()          ' 加這個才可以在中途印出txtbox

            FileOpen(1, Dir + "\" + FName, OpenMode.Input)          '預讀資料：取得資料數量及上下界
            Do Until EOF(1)
                Total += 1
                Input(1, Wind.DatTim)                     ' 時間, X, Y
                Input(1, X)
                Input(1, Y)
                SX += X : SX2 += X ^ 2 : SY += Y : SY2 += Y ^ 2
                If Total = 1 Then
                    MinX = X : MinY = Y : MaxX = X : MaxY = Y
                Else
                    If MaxX < X Then MaxX = X ' 計算上下界
                    If MaxY < Y Then MaxY = Y
                    If MinX > X Then MinX = X
                    If MinY > Y Then MinY = Y
                End If
                If Total Mod 1000 = 0 Then TextBox16.Text = "預讀中，記錄：" & Total : Application.DoEvents() ' 加這個才可以在中途印出txtbox
            Loop
            TextBox16.Text = "整理完畢，可以開始分類了" & vbCrLf & "目的目錄和檔名還沒改請趁早"
            FileClose(1)

            '預設分布組界、組距
            Label17.Text = "基本資料: " & vbCrLf &
                           "資料數量：" & Total & vbCrLf &
                           "X軸：" & vbCrLf &
                           "最小：" & MinX & "    " & "最大：" & MaxX & vbCrLf &
                           "平均：" & SX / Total & "    " & "變方：" & (SX2 - SX ^ 2 / Total) / Total & vbCrLf &
                           "Y軸：" & vbCrLf &
                           "最小：" & MinY & "    " & "最大：" & MaxY & vbCrLf &
                           "平均：" & SY / Total & "    " & "變方：" & (SY2 - SY ^ 2 / Total) / Total & vbCrLf
            'TextBox16.Text &= vbCrLf & Label17.Text
            ' StepX = 0.1 : StepY = 0.1
            StepX = Int((MaxX - MinX) * 100) / 1000
            StepY = Int((MaxY - MinY) * 100) / 1000
            NumX = Int((MaxX - MinX) / StepX) + 1
            NumY = Int((MaxY - MinY) / StepY) + 1
            MinX0 = Int(MinX / StepX) * StepX
            MinY0 = Int(MinY / StepY) * StepY
            TextBox10.Text = MinX0 : TextBox11.Text = MinX0 + (NumX + 1) * StepX : TextBox12.Text = StepX
            TextBox13.Text = MinY0 : TextBox14.Text = MinY0 + (NumY + 1) * StepY : TextBox15.Text = StepY
            Application.DoEvents()          ' 加這個才可以在中途印出txtbox
        End If
    End Sub

    '計算併輸出頻度分布資料
    Private Sub Button9_Click(sender As System.Object, e As System.EventArgs) Handles Button9.Click
        Dim I, J, Total, NumX, NumY, A, B As Integer
        Dim MinX, MinY, MaxX, MaxY, X, Y, StepX, StepY As Decimal
        Dim S, Dir1 As String
        MinX = TextBox10.Text : MaxX = TextBox11.Text : StepX = TextBox12.Text
        NumX = (MaxX - MinX) / StepX : MaxX = MinX + (NumX + 1) * StepX : TextBox11.Text = MaxX
        MinY = TextBox13.Text : MaxY = TextBox14.Text : StepY = TextBox15.Text
        NumY = (MaxY - MinY) / StepY : MaxY = MinY + (NumY + 1) * StepY : TextBox14.Text = MaxY
        Dim Freq(NumX, NumY) As Integer
        SName = TextBox9.Text
        Dir1 = TextBox64.Text
        FileOpen(1, Dir + "\" + FName, OpenMode.Input)          '讀取資料
        Do Until EOF(1)
            Input(1, Wind.DatTim)                     ' 時間, X, Y
            Input(1, X)
            Input(1, Y)
            If X >= MinX And X <= MaxX And Y >= MinY And Y <= MaxY Then
                A = Int((X - MinX) / StepX)
                B = Int((Y - MinY) / StepY)
                Freq(A, B) += 1
            End If
            Total += 1
            If Total Mod 1000 = 0 Then TextBox16.Text = "計算中，記錄：" & Total : Application.DoEvents() ' 加這個才可以在中途印出txtbox
        Loop
        FileClose(1)
        FileOpen(2, Dir1 + "\" + SName, OpenMode.Output)
        PrintLine(2, "輸出檔名：" & SName)
        PrintLine(2, vbCrLf & Label17.Text)
        S = "組界：" & vbCrLf &
            "X軸：" & MinX & " - " & MaxX & vbCrLf &
            "Y軸：" & MinY & " - " & MaxY
        PrintLine(2, S)
        S = "組距：" & vbCrLf &
            "X軸：" & StepX & vbCrLf &
            "Y軸：" & StepY
        PrintLine(2, S)
        S = "            : "
        For I = 0 To NumX
            If I > 0 Then S &= " , "
            S &= MinX + I * StepX
        Next
        PrintLine(2, S)
        For J = 0 To NumY
            S = (MinY + J * StepY) & " : "
            For I = 0 To NumX
                '        S = (MinX + I * StepX) & "," & (MinY + J * StepY) & "," & Freq(I, J)
                S &= Freq(I, J) & " , "
            Next
            PrintLine(2, S)
        Next
        FileClose(2)
        TextBox16.Text &= vbCrLf & "OK!"
    End Sub

    '===== 5. 產生模擬數據 =====

    '產生模擬數據-讀取檔案
    Private Sub Button13_Click(sender As System.Object, e As System.EventArgs) Handles Button13.Click
        Dim I, N, Type As Integer
        Dim S1, S2 As String

        If OpenFileDialog1.ShowDialog() = Windows.Forms.DialogResult.OK Then            '啟用檔案對話框，輸入檔名FName
            FName = OpenFileDialog1.FileName
            N = Len(FName)
            For I = N To 1 Step -1                                                      '將檔名分割為路徑Dir和檔名FName
                If Mid(FName, I, 1) = "\" Then
                    Dir = Microsoft.VisualBasic.Left(FName, I - 1)
                    FName = Mid(FName, I + 1)
                End If
            Next
            Label46.Text = "目錄：" & Dir
            Label47.Text = "檔名：" & FName
            TextBox65.Text = Dir
            Dim Min, Max As Integer                                 ' 設定輸出資料區間
            FileOpen(1, Dir + "\" + FName, OpenMode.Random, , , DataLen)
            Min = 1 : Max = LOF(1) \ DataLen                        '預先設定資料範圍為全部
            TrackBar4.Minimum = 1 : TrackBar4.Maximum = Max
            TrackBar4.TickFrequency = Max \ 20 : TrackBar4.LargeChange = Max \ 200
            TrackBar3.Minimum = 1 : TrackBar3.Maximum = Max
            TrackBar3.TickFrequency = Max \ 20 : TrackBar3.LargeChange = Max \ 200
            TrackBar4.Value = 1 : TrackBar3.Value = Max
            TextBox54.Text = "資料數量：" & Max & "筆"
            FileGet(1, Wind, TrackBar4.Value)
            Label44.Text = "起始時間：" & Wind.DatTim
            FileGet(1, Wind, TrackBar3.Value)
            Label45.Text = "終止時間：" & Wind.DatTim

            '    RadioButton18.Checked = True
            Application.DoEvents()                                          ' 加這個才可以在中途印出txtbox

            '判斷資料類型
            FileGet(1, Wind, 1) : S1 = Wind.DatTim : Type = 0
            FileGet(1, Wind, 2) : S2 = Wind.DatTim
            TextBox54.Text &= vbCrLf & S1 & " , " & S2

            If Microsoft.VisualBasic.Mid(S1, 18, 2) <> "00" Or Microsoft.VisualBasic.Mid(S2, 18, 2) <> "00" Then
                Type = 1
            Else
                If Microsoft.VisualBasic.Mid(S1, 15, 5) <> "00:00" Or Microsoft.VisualBasic.Mid(S2, 15, 5) <> "00:00" Then
                    Type = 2
                Else
                    Type = 3
                End If
            End If
            If Type = 1 Then RadioButton18.Checked = True : TextBox54.Text &= vbCrLf & "原始檔已是每秒數據!"
            If Type = 2 Then RadioButton19.Checked = True
            If Type = 3 Then RadioButton17.Checked = True
            FileClose(1)
        End If
    End Sub

    ' 設定輸出檔名-有勾選才設
    Private Sub CheckBox1_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles CheckBox1.CheckedChanged, CheckBox2.CheckedChanged
        If CheckBox1.Checked = True Then TextBox46.Text = FName + "-data.txt" Else TextBox46.Text = ""
        If CheckBox2.Checked = True Then TextBox55.Text = FName + "-fall.txt" Else TextBox55.Text = ""
    End Sub

    '產生模擬數據-設定起始和終止時間～不需輸出全區間
    Private Sub TrackBar4_Scroll(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TrackBar4.Scroll
        FileOpen(1, Dir + "\" + FName, OpenMode.Random, , , DataLen)
        FileGet(1, Wind, TrackBar4.Value)
        Label44.Text = "起始時間：" & Wind.DatTim
        FileClose(1)
    End Sub

    Private Sub TrackBar3_Scroll(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TrackBar3.Scroll
        FileOpen(1, Dir + "\" + FName, OpenMode.Random, , , DataLen)
        FileGet(1, Wind, TrackBar3.Value)
        Label45.Text = "終止時間：" & Wind.DatTim
        FileClose(1)
    End Sub

    ' 設定各時段花粉產生數量，並自動轉換為頻度
    Private Sub TextBox47_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox47.TextChanged
        Dim I As Integer
        Dim Sum, J As Single
        Dim TF() As String
        TF = Split(TextBox47.Text, vbCrLf)
        TextBox48.Text = ""
        If TF.Length = 24 Then
            For I = 0 To 23
                TimeFreq(I, 0) = Val(TF(I)) : Sum += TimeFreq(I, 0)
            Next
            For I = 0 To 23
                TimeFreq(I, 1) = TimeFreq(I, 0) / Sum : J = TimeFreq(I, 1)
                TextBox48.Text &= J.ToString("0.000000") & vbCrLf
            Next
        Else                                                                ' 若行數非24行表示有問題
            TextBox48.Text = "數量不對" & vbCrLf & "請檢查資料有無增刪"
        End If
    End Sub

    ' 檢查三種係數行數是否正確
    Private Sub TextBox50_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox50.TextChanged, TextBox51.TextChanged, TextBox52.TextChanged
        Dim TF() As String
        TextBox54.Text = ""
        TF = Split(TextBox52.Text, vbCrLf)
        If TF.Length <> 6 Then TextBox54.Text &= "風速" & vbCrLf
        TF = Split(TextBox50.Text, vbCrLf)
        If TF.Length <> 6 Then TextBox54.Text &= "風向x向量" & vbCrLf
        TF = Split(TextBox51.Text, vbCrLf)
        If TF.Length <> 6 Then TextBox54.Text &= "風向y向量" & vbCrLf
        If TextBox54.Text <> "" Then                                        ' 若係數行數非6行表示有問題
            TextBox54.Text &= "係數有問題," & vbCrLf & "請檢查資料有無增刪"
        End If
    End Sub

    ' 設定是否顯示輸入粒子總數和時段比例欄位
    Private Sub RadioButton20_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles RadioButton20.CheckedChanged, RadioButton21.CheckedChanged, RadioButton22.CheckedChanged, RadioButton23.CheckedChanged
        Dim Mult As Integer                         'Mult : 放大倍率
        If RadioButton18.Checked = True Then Mult = 1
        If RadioButton19.Checked = True Then Mult = 60
        If RadioButton25.Checked = True Then Mult = 600
        If RadioButton17.Checked = True Then Mult = 3600
        If RadioButton20.Checked = True Or RadioButton21.Checked = True Then        ' 若選每秒一顆，自然就不用再輸入時段比例，粒子總數 = 分x60 = 時x3600
            Label51.Text = "產生粒子總數量："
            GroupBox11.Visible = False
            TextBox53.Text = (TrackBar3.Value - TrackBar4.Value + 1) * Mult
        Else
            Label51.Text = "每日粒子產生量："
            GroupBox11.Visible = True
            TextBox53.Visible = True
        End If
    End Sub

    ' 顯示三種線性模式的說明文字
    Private Sub Button15_Click_1(sender As System.Object, e As System.EventArgs) Handles Button15.Click
        Label50.Visible = Not Label50.Visible       '直接轉換顯示屬性 開 <-> 關
    End Sub

    ' 開始產生模擬數據及落點位置
    Private Sub Button16_Click(sender As System.Object, e As System.EventArgs) Handles Button16.Click
        Randomize(Guid.NewGuid().GetHashCode())                 '使用Guid函數當作亂數種子
        Dim I, J, F, ST, EN, DN, Hr, No, Mult, BSMode As Integer 'Mult-放大倍率, BSMode-取樣方式, DN-日夜, Hr-小時
        Dim Total, TotalD As Integer                            'Total-粒子總數, TotalD-一日粒子總數
        Dim Coef(5, 3), Tmp, RMult, Max As Double               'Coef : 模式係數, Tmp: 暫存, Max: 時段比例最大值
        Dim DirX, DirY, Deg, PosX, PosY As Double               'DirX.Y:風向XY向量, PosX.Y:落點XY座標
        Dim SigmaSpeed, SigmaDirX, SigmaDirY As Double          'Sigma : 風速風向標準差
        Dim BSSpeed, BSDirX, BSDirY, BSSpeed1, BSDirX1, BSDirY1 As Double                   'BS: 模擬風速風向
        Dim TF(), S, Dir1 As String                             'TF() : 暫存陣列
        Dim D1, D2 As Date                                      'D1, D2: 暫存日期

        If RadioButton18.Checked = True Then Mult = 1 'Mult-放大倍率(重複取樣次數)
        If RadioButton19.Checked = True Then Mult = 60
        If RadioButton25.Checked = True Then Mult = 600
        If RadioButton17.Checked = True Then Mult = 3600
        If RadioButton20.Checked = True Then BSMode = 1 'BSMode-取樣方式
        If RadioButton21.Checked = True Then BSMode = 2
        If RadioButton22.Checked = True Then BSMode = 3
        If RadioButton23.Checked = True Then BSMode = 4
        ST = TrackBar4.Value                                    'ST: 起始時間比數
        EN = TrackBar3.Value                                    'EN: 終止時間
        TotalD = Val(TextBox53.Text)                            'TotalD: 一天總量
        'RMult = Math.Sqrt(Mult)                                'RMulit: 變異數需要乘以sqrt(n)嗎?
        RMult = 1
        TF = Split(TextBox52.Text, vbCrLf)                      '風速係數
        For I = 0 To 5 : Coef(I, 0) = Val(TF(I)) : Next
        TF = Split(TextBox50.Text, vbCrLf)                      '風向x向量係數
        For I = 0 To 5 : Coef(I, 1) = Val(TF(I)) : Next
        TF = Split(TextBox51.Text, vbCrLf)                      '風向y向量係數
        For I = 0 To 5 : Coef(I, 2) = Val(TF(I)) : Next
        Dir1 = TextBox65.Text
        FileOpen(1, Dir + "\" + FName, OpenMode.Random, , , DataLen)
        If CheckBox1.Checked = True Then FileOpen(2, Dir + "\" + TextBox46.Text, OpenMode.Output)
        If CheckBox2.Checked = True Then FileOpen(3, Dir + "\" + TextBox55.Text, OpenMode.Output)

        Select Case BSMode

            Case 1, 3                                        '「順序」產生  1-每秒固定一個，3-設定總量依比例
                For I = ST To EN
                    FileGet(1, Wind, I)
                    DirX = Math.Cos(Wind.Direct) : DirY = Math.Sin(Wind.Direct)
                    S = Wind.DatTim                                 'S-資料日期
                    Hr = Val(Microsoft.VisualBasic.Mid(S, 15, 2))   '取得小時Hr
                    DN = 0 : If Hr > 5 Or Hr < 18 Then DN = 1 '判定日夜DN
                    If Microsoft.VisualBasic.Mid(S, 15, 5) = "00:00" Then
                        TextBox54.Text = S
                        Application.DoEvents()               ' 加這個才可以在中途印出txtbox
                    End If
                    If BSMode = 1 Then                          '計算每個原始數據需要產生多少模擬數據(No)
                        No = Mult
                    Else
                        No = TimeFreq(Hr, 1) * Total
                        If Mult = 60 Then No /= Mult
                    End If
                    If No > 1 Then                                                      '除了每秒一顆以外的狀況
                        '計算風速風向變異數
                        SigmaSpeed = Coef(0, 0) + Coef(1, 0) * Wind.Speed + Coef(2, 0) * Wind.Temperature + Coef(3, 0) * Wind.RelHumidity + Coef(4, 0) * Wind.AtmPressure + Coef(5, 0) * DN
                        SigmaDirX = Coef(0, 1) + Coef(1, 1) * Wind.Speed + Coef(2, 1) * Wind.Temperature + Coef(3, 1) * Wind.RelHumidity + Coef(4, 1) * Wind.AtmPressure + Coef(5, 1) * DN
                        SigmaDirY = Coef(0, 2) + Coef(1, 2) * Wind.Speed + Coef(2, 2) * Wind.Temperature + Coef(3, 2) * Wind.RelHumidity + Coef(4, 2) * Wind.AtmPressure + Coef(5, 2) * DN
                        For J = 1 To No                                               '每分或每小時數據
                            '分別取常態逢機數字
                            Normal(Wind.Speed, SigmaSpeed * RMult, BSSpeed, BSSpeed1)
                            Normal(DirX, SigmaDirX * RMult, BSDirX, BSDirX1)
                            Normal(DirY, SigmaDirY * RMult, BSDirY, BSDirY1)
                            Dist(BSSpeed, BSDirX, BSDirY, Deg, PosX, PosY)              '輸出模擬數據2次
                            If CheckBox1.Checked = True Then PrintLine(2, S & " , " & BSSpeed & " , " & Deg)
                            If CheckBox2.Checked = True Then PrintLine(3, S & " , " & PosX & " , " & PosY)
                            Dist(BSSpeed1, BSDirX1, BSDirY1, Deg, PosX, PosY)
                            If CheckBox1.Checked = True Then PrintLine(2, S & " , " & BSSpeed1 & " , " & Deg)
                            If CheckBox2.Checked = True Then PrintLine(3, S & " , " & PosX & " , " & PosY)
                            J += 1
                        Next
                    Else                                '每秒一顆
                        '計算風速風向變異數
                        SigmaSpeed = Coef(0, 0) + Coef(1, 0) * Wind.Speed + Coef(2, 0) * Wind.Temperature + Coef(3, 0) * Wind.RelHumidity + Coef(4, 0) * Wind.AtmPressure
                        SigmaDirX = Coef(0, 1) + Coef(1, 1) * Wind.Speed + Coef(2, 1) * Wind.Temperature + Coef(3, 1) * Wind.RelHumidity + Coef(4, 1) * Wind.AtmPressure
                        SigmaDirY = Coef(0, 2) + Coef(1, 2) * Wind.Speed + Coef(2, 2) * Wind.Temperature + Coef(3, 2) * Wind.RelHumidity + Coef(4, 2) * Wind.AtmPressure
                        '分別取常態逢機數字
                        Normal(Wind.Speed, SigmaSpeed * RMult, BSSpeed, BSSpeed1)
                        Normal(DirX, SigmaDirX * RMult, BSDirX, BSDirX1)
                        Normal(DirY, SigmaDirY * RMult, BSDirY, BSDirY1)
                        Dist(BSSpeed, BSDirX, BSDirY, Deg, PosX, PosY)                   '輸出模擬數據2次
                        If CheckBox1.Checked = True Then PrintLine(2, S & " , " & BSSpeed & " , " & Deg)
                        If CheckBox2.Checked = True Then PrintLine(3, S & " , " & PosX & " , " & PosY)
                        Dist(BSSpeed1, BSDirX1, BSDirY1, Deg, PosX, PosY)
                        If CheckBox1.Checked = True Then PrintLine(2, S & " , " & BSSpeed1 & " , " & Deg)
                        If CheckBox2.Checked = True Then PrintLine(3, S & " , " & PosX & " , " & PosY)
                        I += 1

                        '每秒一顆，直接輸出不用拔靴
                        'If CheckBox1.Checked = True Then                                '輸出模擬數據
                        '    PrintLine(2, S & " , " & Wind.Speed & " , " & Math.Atan2(DirY, DirX) / Radians & " , " & DirX & " , " & DirY)
                        'End If
                        'Dist(Wind.Speed, DirX, DirY, PosX, PosY)                        '輸出模擬落點
                        'If CheckBox2.Checked = True Then PrintLine(3, S & " , " & PosX & " , " & PosY)
                    End If
                Next

            Case 2, 4                                       '「逢機」產生, 2-每秒固定一個，4-設定總量依比例

                For I = 0 To 23                             '設定各時段比例最大值Max
                    If Max < TimeFreq(I, 1) Then Max = TimeFreq(I, 1)
                Next
                If BSMode = 2 Then                          '計算需要逢機幾次(Total)
                    Total = (EN - ST + 1) * Mult
                Else
                    FileGet(1, Wind, ST) : D1 = Wind.DatTim
                    FileGet(1, Wind, EN) : D2 = Wind.DatTim
                    Total = DateDiff("D", D1, D2) * TotalD          '計算全程完整日發散粒子
                    For I = Hour(D1) To 23                          '累計起始日剩餘時數發散粒子
                        Total += TimeFreq(I, 1) * TotalD
                    Next
                    For I = 0 To Hour(D2)                           '累計結束日剩餘時數發散粒子
                        Total += TimeFreq(I, 1) * TotalD
                    Next
                End If

                For I = 1 To Total                                  '判斷選中時段是否符合選取機率，否則重選
                    F = 0
                    Do Until F = 1
                        Tmp = Rnd()
                        If Tmp < Max Then
                            J = Int(Rnd() * (EN - ST + 1)) + ST
                            FileGet(1, Wind, J)
                            If Tmp <= TimeFreq(Hour(Wind.DatTim), 1) Then F = 1
                        End If
                    Loop

                    DirX = Math.Cos(Wind.Direct) : DirY = Math.Sin(Wind.Direct)
                    If I Mod 1000 = 1 Then
                        TextBox54.Text = I - 1 & " - " & Wind.DatTim
                        Application.DoEvents()               ' 加這個才可以在中途印出txtbox
                    End If
                    '計算風速風向變異數
                    SigmaSpeed = Coef(0, 0) + Coef(1, 0) * Wind.Speed + Coef(2, 0) * Wind.Temperature + Coef(3, 0) * Wind.RelHumidity + Coef(4, 0) * Wind.AtmPressure
                    SigmaDirX = Coef(0, 1) + Coef(1, 1) * Wind.Speed + Coef(2, 1) * Wind.Temperature + Coef(3, 1) * Wind.RelHumidity + Coef(4, 1) * Wind.AtmPressure
                    SigmaDirY = Coef(0, 2) + Coef(1, 2) * Wind.Speed + Coef(2, 2) * Wind.Temperature + Coef(3, 2) * Wind.RelHumidity + Coef(4, 2) * Wind.AtmPressure
                    '分別取常態逢機數字
                    Normal(Wind.Speed, SigmaSpeed * RMult, BSSpeed, BSSpeed1)
                    Normal(DirX, SigmaDirX * RMult, BSDirX, BSDirX1)
                    Normal(DirY, SigmaDirY * RMult, BSDirY, BSDirY1)
                    Dist(BSSpeed, BSDirX, BSDirY, Deg, PosX, PosY)                   '輸出模擬數據2次
                    If CheckBox1.Checked = True Then PrintLine(2, Wind.DatTim & " , " & BSSpeed & " , " & Deg)
                    If CheckBox2.Checked = True Then PrintLine(3, Wind.DatTim & " , " & PosX & " , " & PosY)
                    Dist(BSSpeed1, BSDirX1, BSDirY1, Deg, PosX, PosY)
                    If CheckBox1.Checked = True Then PrintLine(2, Wind.DatTim & " , " & BSSpeed1 & " , " & Deg)
                    If CheckBox2.Checked = True Then PrintLine(3, Wind.DatTim & " , " & PosX & " , " & PosY)
                    I += 1
                Next
        End Select
        TextBox54.Text &= vbCrLf & "轉檔結束"
        FileClose(1)
        If CheckBox1.Checked = True Then FileClose(2)
        If CheckBox2.Checked = True Then FileClose(3)

    End Sub

    ' 由氣象資料轉換變異數
    Sub Formula_Y(Coef(), X, ByRef Y)
        Y = Coef(0) + Coef(1) * X + Coef(2) * Wind.Speed + Coef(3) * Wind.Temperature + Coef(4) * Wind.RelHumidity + Coef(5) * Wind.AtmPressure
    End Sub

    ' 由平均風速及XY向量估算落點
    Sub Dist(SP, DirX, DirY, ByRef Deg, ByRef PosX, ByRef PosY)
        Static SumPos, SumNo As Integer
        If SumPos = 0 Then
            SumNo = Val(TextBox56.Text)
        End If
        Static Dx(SumNo), Dy(SumNo) As Double
        SumPos += 1 : If SumPos > SumNo Then SumPos = 1 ' 暫存累計位置指標
        Deg = Math.Atan2(DirX, DirY) / Radians
        If Deg < 0 Then Deg += 360
        DirX = SP * Math.Cos(Deg * Radians)
        DirY = SP * Math.Sin(Deg * Radians)
        PosX += DirX : PosY += DirY
        PosX -= Dx(SumPos) : PosY -= Dy(SumPos) ' 得到累計位置 PosX,Y
        Dx(SumPos) = DirX : Dy(SumPos) = DirY
    End Sub


    '======  6. 基本資料試算   ======

    '======  6-1. 基本資料標準化轉換   ======
    Private Sub Button11_Click(sender As System.Object, e As System.EventArgs) Handles Button11.Click
        Application.DoEvents()          ' 加這個才可以在中途印出txtbox
        Dim Exp As Integer
        If RadioButton10.Checked = True Then Exp = -6 Else Exp = -3
        TextBox24.Text = Val(TextBox21.Text) * 10 ^ Exp
        If RadioButton12.Checked = True Then Exp = -6
        If RadioButton7.Checked = True Then Exp = -3
        If RadioButton9.Checked = True Then Exp = -2
        TextBox25.Text = Val(TextBox22.Text) * 10 ^ Exp
        If RadioButton14.Checked = True Then Exp = 3 Else Exp = -3
        TextBox26.Text = Val(TextBox23.Text) * 10 ^ Exp
        If RadioButton13.Checked = True Then
            TextBox28.Text = Val(TextBox27.Text) * 3.6 : Label28.Text = "km / hr"
        Else
            TextBox28.Text = Val(TextBox27.Text) / 3.6 : Label28.Text = "m / sec"
        End If
    End Sub

    '======  6-2. 微粒子相關資料試算   ======

    Private Sub Button12_Click(sender As System.Object, e As System.EventArgs) Handles Button12.Click
        ' Dim  As Integer
        Dim Dp, Pp, Pf, G, Mu, Ap, V, Vt1, Vt2, Vt3, Rep As Single
        '      Dim S_temp As String
        Dp = Val(TextBox29.Text)                                       'Dp = 粒徑
        Pf = Val(TextBox34.Text)                                       'Pf = 流體密度
        G = Val(TextBox41.Text)                                        'G  = 重力加速度
        Mu = Val(TextBox42.Text)                                       'Mu = 流體黏滯係數
        Ap = Math.PI * Dp ^ 2 / 4 : TextBox32.Text = Ap                'Ap = 截面積
        V = Math.PI * Dp ^ 3 / 6 : TextBox33.Text = V                  'V  = 體積
        If Val(TextBox30.Text) > 0 Then                                '由粒重計算粒子密度Pp
            Pp = Val(TextBox30.Text) / V : TextBox31.Text = Pp
        End If
        If Val(TextBox31.Text) > 0 Then                                '由粒子密度Pp計算粒重
            Pp = Val(TextBox31.Text) : TextBox30.Text = Pp * V
        End If

        Vt1 = G * Dp ^ 2 * (Pp - Pf) / (18 * Mu) : TextBox35.Text = Vt1               'Vt1= 終端速度(微粒子)
        'Vt2 = (4 / 225 * (Pp - Pf) / (Pf * Mu)) ^ (1 / 3) * Dp : TextBox36.Text = Vt2 'Vt2= 終端速度(中粒子)
        Vt2 = ((Pp - Pf) * G / 18 * Math.Sqrt(Dp ^ 3 / (Pf * Mu))) ^ (2 / 3)       'Vt2= 終端速度(中粒子)
        TextBox36.Text = Vt2
        Vt3 = Math.Sqrt(3 * G * (Pp - Pf) / Pf * Dp) : TextBox37.Text = Vt3      'Vt3= 終端速度(粗粒子)
        Rep = Dp * Vt1 * Pf / Mu : TextBox38.Text = Rep                   'Rep= 雷諾數，以微粒子計算
        If Rep < 1 Then TextBox43.Text = "微粒子"
        If Rep >= 1 And Rep <= 500 Then TextBox43.Text = "中粒子"
        If Rep > 500 Then TextBox43.Text = "粗粒子"
        Rep = Dp * Vt2 * Pf / Mu : TextBox39.Text = Rep                   'Rep= 雷諾數，以中粒子計算
        If Rep < 1 Then TextBox44.Text = "微粒子"
        If Rep >= 1 And Rep <= 500 Then TextBox44.Text = "中粒子"
        If Rep > 500 Then TextBox44.Text = "粗粒子"
        Rep = Dp * Vt3 * Pf / Mu : TextBox40.Text = Rep                   'Rep= 雷諾數，以粗粒子計算
        If Rep < 1 Then TextBox45.Text = "微粒子"
        If Rep >= 1 And Rep <= 500 Then TextBox45.Text = "中粒子"
        If Rep > 500 Then TextBox45.Text = "粗粒子"
    End Sub

    ' ===== 6-3.測試常態逢機變數產生效率 =====
    Private Sub Button17_Click(sender As System.Object, e As System.EventArgs) Handles Button17.Click
        Randomize(Guid.NewGuid().GetHashCode())                 '使用Guid函數當作亂數種子
        Dim I, N As Integer
        Dim Mean, Sigma, T3, N1, N2, S, S2 As Double
        'Dim S, S2 As Decimal
        Dim T1, T2 As String
        Mean = Val(TextBox57.Text)
        Sigma = Val(TextBox58.Text)
        N = Val(TextBox59.Text)
        TextBox60.Text = "測試開始，產生" & N & "個常態逢機變數" & vbCrLf

        ' 原始Box-Muller法
        TextBox60.Text &= "原始Box-Muller法" & vbCrLf : T1 = Now
        TextBox60.Text &= "起始時間：" & T1 & vbCrLf
        S = 0 : S2 = 0
        For I = 1 To N / 2
            Normal(Mean, Sigma, N1, N2)
            S += N1 + N2
            S2 += N1 ^ 2 + N2 ^ 2
            '       If I Mod 1000 = 0 Then TextBox60.Text &= N1 & ", " & N2 & vbCrLf
        Next
        S = S / N
        S2 = (S2 - S ^ 2 * N) / (N - 1)
        '     TextBox60.Text &= N1 & ", " & N2 & vbCrLf
        T2 = Now
        TextBox60.Text &= "平均：" & S & vbCrLf
        TextBox60.Text &= "變方：" & S2 & vbCrLf
        TextBox60.Text &= "結束時間：" & T2 & vbCrLf
        TextBox60.Text &= "使用時間：" & DateDiff("S", T1, T2) & "秒" & vbCrLf & vbCrLf
        T3 = DateDiff("S", T1, T2)
        Application.DoEvents()                                          ' 加這個才可以在中途印出txtbox

        ' 修正Box-Muller法
        TextBox60.Text &= "修正Box-Muller法" & vbCrLf : T1 = Now
        TextBox60.Text &= "起始時間：" & T1 & vbCrLf
        S = 0 : S2 = 0
        For I = 1 To N / 2
            Normal1(Mean, Sigma, N1, N2)
            S += N1 + N2
            S2 += N1 ^ 2 + N2 ^ 2
            '        If I Mod 1000 = 0 Then TextBox60.Text &= N1 & ", " & N2 & vbCrLf
        Next
        S = S / N
        S2 = (S2 - S ^ 2 * N) / (N - 1)
        T2 = Now
        TextBox60.Text &= "平均：" & S & vbCrLf
        TextBox60.Text &= "變方：" & S2 & vbCrLf
        TextBox60.Text &= "結束時間：" & T2 & vbCrLf
        TextBox60.Text &= "使用時間：" & DateDiff("S", T1, T2) & "秒" & vbCrLf & vbCrLf
        TextBox60.Text &= "相差：" & T3 - DateDiff("S", T1, T2) & "秒"
    End Sub

    'Box-Muller法產生常態逢機數字
    Sub Normal(Mean, Sigma, ByRef N1, ByRef N2)
        Dim U1, U2 As Double
        U1 = Math.Sqrt(-2 * Math.Log(Rnd()))
        U2 = 2 * Math.PI * Rnd()
        N1 = U1 * Math.Cos(U2) * Sigma + Mean
        N2 = U1 * Math.Sin(U2) * Sigma + Mean
    End Sub

    ' Box-Muller法產生常態逢機數字~修正版 (不用三角函數)
    Sub Normal1(Mean, Sigma, ByRef N1, ByRef N2)
        Dim S, U1, U2 As Double : S = 2
        Do Until S <= 1
            U1 = 2 * Rnd() - 1 : U2 = 2 * Rnd() - 1
            S = U1 ^ 2 + U2 ^ 2
        Loop
        S = Math.Sqrt(-2 * Math.Log(S) / S)
        N1 = U1 * S * Sigma + Mean
        N2 = U2 * S * Sigma + Mean
    End Sub

    ' ===== 6.分割資料檔案 =====
    Private Sub Button19_Click(sender As System.Object, e As System.EventArgs) Handles Button19.Click
        Dim I, N, F As Integer
        Dim S As String
        If OpenFileDialog1.ShowDialog() = Windows.Forms.DialogResult.OK Then            '啟用檔案對話框，輸入檔名FName
            FName = OpenFileDialog1.FileName
            N = Len(FName)
            For I = N To 1 Step -1                                                      '將檔名分割為路徑Dir和檔名FName
                If Mid(FName, I, 1) = "\" Then
                    Dir = Microsoft.VisualBasic.Left(FName, I - 1)
                    FName = Mid(FName, I + 1)
                End If
            Next
            Label68.Text = "目錄：" & Dir
            Label69.Text = "檔名：" & FName
            TextBox66.Text = Dir
            TextBox69.Text = FName + "_###.txt"

            '預讀資料
            FileOpen(1, Dir + "\" + FName, OpenMode.Input)
            TextBox70.Text = "預讀最初10筆資料：" & vbCrLf
            For I = 1 To 10
                S = LineInput(1)
                If I = 1 And Mid(S, 2, 2) = "20" Then F = 1 : CheckBox5.Checked = True
                TextBox70.Text &= S & vbCrLf
            Next
            FileClose(1)
        End If
    End Sub

    ' 分割資料檔案 - 開始轉檔
    Private Sub Button20_Click(sender As System.Object, e As System.EventArgs) Handles Button20.Click
        Dim F, F1, I, SP, M As Integer
        Dim Temp_A As String = ""
        Dim S As String = ""
        Dim S1(4), MM, DD, HH, MM1, DD1, HH1 As String
        Dim Dir1 As String = ""
        '分割模式
        If RadioButton26.Checked = True Then SP = 1
        If RadioButton27.Checked = True Then SP = 2
        If RadioButton28.Checked = True Then SP = 3
        If RadioButton29.Checked = True Then SP = 4
        '合併模式
        If RadioButton30.Checked = True Then M = 1
        If RadioButton32.Checked = True Then M = 2
        If RadioButton33.Checked = True Then M = 3

        Dir1 = TextBox63.Text
        FileOpen(1, Dir + "\" + FName, OpenMode.Input)
        If CheckBox5.Checked = True Then
            F = 1
            For I = 1 To 4                      ' 預讀檔頭資料
                S1(I) = LineInput(1)
            Next
        End If
        F1 = 0
        Do Until EOF(1)
            S = LineInput(1)
            MM = Microsoft.VisualBasic.Mid(S, 6, 2)
            DD = Microsoft.VisualBasic.Mid(S, 9, 2)
            HH = Microsoft.VisualBasic.Mid(S, 12, 2)
            If F1 = 0 Then MM1 = MM : DD1 = DD : HH1 = HH : F1 = 1
            FileOpen(2, Dir1 + "\" + SName, OpenMode.Output)
            If Microsoft.VisualBasic.Left(TextBox20.Text, 2) <> "" Then
                FileOpen(3, TextBox20.Text, OpenMode.Input)
                F1 = 1
            End If
        Loop

    End Sub

End Class