Imports System.IO
Imports System.Web
Imports MySql.Data.MySqlClient
Imports NPOI.SS.UserModel
Imports NPOI.XSSF.UserModel

Module Module1
    Public CuDr As String       ''exeの動くカレントディレクトリを格納
    Public mysqlCon As New MySqlConnection
    Public sqlCommand As New MySqlCommand

    Sub sql_st()
        ''データベースに接続

        Dim Builder = New MySqlConnectionStringBuilder()
        ' データベースに接続するために必要な情報をBuilderに与える。データベース情報はGitに乗せないこと。
        Builder.Server = ""
        Builder.Port =
        Builder.UserID = ""
        Builder.Password = ""
        Builder.Database = ""

        Dim ConStr = Builder.ToString()

        mysqlCon.ConnectionString = ConStr
        mysqlCon.Open()

    End Sub

    Sub sql_cl()
        ' データベースの切断
        mysqlCon.Close()
    End Sub

    Function sql_result_return(ByVal query As String) As DataTable
        ''データセットを返すSELECT系のSQLを処理するコード

        Dim dt As New DataTable()
        Try
            ' 4.データ取得のためのアダプタの設定
            Dim Adapter = New MySqlDataAdapter(query, mysqlCon)
            Adapter.SelectCommand.CommandTimeout = 600

            ' 5.データを取得
            Dim Ds As New DataSet
            Adapter.Fill(dt)

            Return dt
        Catch ex As Exception

            Return dt
        End Try

    End Function

    Function sql_result_no(ByVal query As String)
        ''データセットを返さない、DELETE、UPDATE、INSERT系のSQLを処理するコード

        Try
            sqlCommand.CommandTimeout = 600
            sqlCommand.Connection = mysqlCon
            sqlCommand.CommandText = query
            sqlCommand.ExecuteNonQuery()

            Return "Complete"
        Catch ex As Exception

            Return ex.Message
        End Try

    End Function


    Function dt2unepocht(ByVal vbdate As DateTime) As Long
        ''VBで使用出来る日付を入力すると、UNIX エポック秒に変換する
        vbdate = vbdate.ToUniversalTime()

        Dim dt1 As New DateTime(1970, 1, 1, 0, 0, 0, 0)
        Dim elapsedTime As TimeSpan = vbdate - dt1

        Return CType(elapsedTime.TotalSeconds, Long)

    End Function
    Sub close_save()
        ''設定用ファイルの保存

        Dim dtx1 As String = ""

        For lp1 As Integer = 1 To 1
            Dim tbxn1 As String = "Cf_TextBox" & lp1.ToString

            Dim cs As Control() = Form1.Controls.Find(tbxn1, True)
            If cs.Length > 0 Then
                dtx1 &= CType(cs(0), TextBox).Text
                dtx1 &= vbCrLf
            End If
        Next

        Dim stCurrentDir As String = System.IO.Directory.GetCurrentDirectory()
        CuDr = stCurrentDir

        Dim excsv1 As IO.StreamWriter
        excsv1 = New IO.StreamWriter(CuDr & "\config.ini", False, System.Text.Encoding.GetEncoding("shift_jis"))
        excsv1.Write(dtx1)
        excsv1.Close()
        excsv1.Dispose()

    End Sub

    Sub excel_DataRead01(ByVal ofn As String)
        ''運賃明細をデータベースに取り込む

        ''データベースと接続
        Call sql_st()

        Dim sql1 As New System.Text.StringBuilder()
        sql1.Clear()

        sql1.Append("SELECT ")
        sql1.Append(" Max(`配送日`)")
        sql1.Append("FROM `gcn_shipping_slip`")

        Dim dTb1 As DataTable = sql_result_return(sql1.ToString)

        If dTb1.Rows.Count = 0 Then
            MsgBox(dTb1.ToString & "データがありません。異常です", MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "警告")
            Debug.Print(sql1.ToString)
        Else
            For Each DRow As DataRow In dTb1.Rows

                Form1.Cf_TextBox100.Text = DRow.Item(0)

            Next
        End If


        Dim sqlh As New System.Text.StringBuilder()

        sqlh.Clear()
        sqlh.Append("INSERT INTO `gcn_shipping_slip` (")
        sqlh.Append(" `伝票種別`")
        sqlh.Append(",`配送日`")
        sqlh.Append(",`お問い合せＮｏ`")
        sqlh.Append(",`都道府県`")
        sqlh.Append(",`個数`")
        sqlh.Append(",`税区分`")
        sqlh.Append(",`運賃`")
        sqlh.Append(",`中継料`")
        sqlh.Append(",`合計`")
        sqlh.Append(",`配送区分`")
        sqlh.Append(" ) VALUES ")

        sql1.Clear()

        Dim rfs As FileStream = File.OpenRead(ofn)
        Dim book01 As IWorkbook = New XSSFWorkbook(rfs)
        rfs.Close()

        Dim sqlcu As Integer = 0

        For numb1 As Integer = 0 To book01.NumberOfSheets - 1 Step 1

            '番号指定でシートを取得する（番号は０～）
            Dim sheet1 As ISheet = book01.GetSheetAt(numb1)

            Dim c1 As Integer = 1

            Dim lpc1 As Integer = 0

            Dim com1 As String = ""
            Dim shname As String = book01.GetSheetName(numb1)
            If shname.IndexOf("BtoC") > 0 Then
                com1 = "BtoC"
            ElseIf shname.IndexOf("BtoB") > 0 Then
                com1 = "BtoB"
            Else
                com1 = "unknown"
            End If

            For r1 As Integer = 0 To sheet1.LastRowNum

                c1 = 1
                Dim dty1 As String = ""

                Try
                    dty1 = sheet1.GetRow(r1).GetCell(c1).ToString()
                    dty1 = dty1.Replace("＜", "")
                    dty1 = dty1.Replace("＞", "")

                    If dty1.IndexOf("払") > 0 Then
                        sql1.Append("(")
                        sql1.Append("'")
                        sql1.Append(dty1)
                        sql1.Append("'")
                    Else
                        Dim chk1 As String = sql1.ToString
                        If chk1.IndexOf("),") > 0 Then

                            'データベースにSQLを発行
                            Dim rst1 = sql_regist1(sqlh.ToString & sql1.ToString)
                            If rst1 <> "Complete" Then
                                MsgBox("SQLが正常に反映されませんでした", MsgBoxStyle.Critical Or MsgBoxStyle.OkOnly, "警告")
                                Debug.Print(sqlh.ToString & sql1.ToString)
                            Else
                                sql1.Clear()
                            End If
                        End If

                        Continue For

                    End If

                Catch ex As System.NullReferenceException
                    sql1.Append("NULL")
                End Try

                sql1.Append(",")

                Dim dty2 As String = sql_numeric_type(sheet1, r1, 2)

                If dty2 = "NULL" Then
                    Continue For
                Else
                    Dim ln1 As Integer = dty2.Length
                    ln1 = ln1 - 2
                    dty2 = dty2.Insert(ln1, "/")

                    Dim m As Integer = dty2.IndexOf("/")
                    If m = 0 Then
                    Else
                        Dim mn As String = dty2.Substring(0, m)

                        Select Case mn
                            Case "1"
                                dty2 = "2018/" & dty2

                            Case "2"
                                dty2 = "2018/" & dty2

                            Case "3"
                                dty2 = "2018/" & dty2

                            Case "4"
                                dty2 = "2018/" & dty2

                            Case "5"
                                dty2 = "2018/" & dty2

                            Case "6"
                                dty2 = "2018/" & dty2

                            Case "7"
                                dty2 = "2018/" & dty2

                            Case "8"
                                dty2 = "2018/" & dty2

                            Case "9"
                                dty2 = "2018/" & dty2

                            Case "10"
                                dty2 = "2018/" & dty2

                            Case "11"
                                dty2 = "2018/" & dty2

                            Case "12"
                                dty2 = "2017/" & dty2

                            Case Else
                                Debug.Print("")

                        End Select


                    End If

                    Dim dt2 As DateTime = DateTime.Parse(dty2)

                    sql1.Append("'")
                    sql1.Append(dt2.Year)
                    sql1.Append("-")
                    sql1.Append(dt2.Month)
                    sql1.Append("-")
                    sql1.Append(dt2.Day)
                    sql1.Append("'")

                End If

                sql1.Append(",")
                sql1.Append(sql_string_type(sheet1, r1, 3))
                sql1.Append(",")
                sql1.Append(sql_string_type(sheet1, r1, 4))
                sql1.Append(",")
                sql1.Append(sql_numeric_type(sheet1, r1, 5))
                sql1.Append(",")
                sql1.Append(sql_string_type(sheet1, r1, 6))
                sql1.Append(",")
                sql1.Append(sql_numeric_type(sheet1, r1, 7))
                sql1.Append(",")
                sql1.Append(sql_numeric_type(sheet1, r1, 8))
                sql1.Append(",")
                sql1.Append(sql_numeric_type(sheet1, r1, 9))
                sql1.Append(",'")
                sql1.Append(com1)
                sql1.Append("'),")

                lpc1 += 1

                Dim chn1 As Integer = Math.Floor(lpc1 / 1000)
                Dim chn2 As Integer = Math.Ceiling(lpc1 / 1000)

                If chn1 = chn2 Then
                    'データベースにSQLを発行
                    Dim rst1 = sql_regist1(sqlh.ToString & sql1.ToString)
                    If rst1 <> "Complete" Then
                        MsgBox("SQLが正常に反映されませんでした", MsgBoxStyle.Critical Or MsgBoxStyle.OkOnly, "警告")
                        Debug.Print(sqlh.ToString & sql1.ToString)
                    End If

                    sql1.Clear()

                    lpc1 = 0
                End If

            Next

            If sql1.ToString = "" Then
            Else

                'データベースにSQLを発行
                Dim rst1 = sql_regist1(sqlh.ToString & sql1.ToString)
                If rst1 <> "Complete" Then
                    MsgBox("SQLが正常に反映されませんでした", MsgBoxStyle.Critical Or MsgBoxStyle.OkOnly, "警告")
                    Debug.Print(sqlh.ToString & sql1.ToString)
                End If
            End If

        Next


        sql1.Clear()

        sql1.Append("SELECT ")
        sql1.Append(" Max(`配送日`)")
        sql1.Append("FROM `gcn_shipping_slip`")

        dTb1 = sql_result_return(sql1.ToString)

        If dTb1.Rows.Count = 0 Then
            MsgBox(dTb1.ToString & "データがありません。異常です", MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "警告")
            Debug.Print(sql1.ToString)
        Else
            For Each DRow As DataRow In dTb1.Rows

                Form1.Cf_TextBox101.Text = DRow.Item(0)

            Next
        End If

        ''切断
        Call sql_cl()

        MessageBox.Show("Excelシート取り込み完了", "完了", MessageBoxButtons.OK, MessageBoxIcon.Information)


    End Sub

    Sub data_processing()

        ''運賃明細リスト

        ''データベースと接続
        Call sql_st()

        Dim sql1 As String = "INSERT INTO `global_fare_details`("
        sql1 &= "`受注番号`,"
        sql1 &= "`購入品商品CODE`,"
        sql1 &= "`購入品名`,"
        sql1 &= "`購入品価格`,"
        sql1 &= "`購入品数量`,"
        sql1 &= "`顧客分類`,"
        sql1 &= "`発送日`,"
        sql1 &= "`発送伝票番号`"
        sql1 &= ") SELECT"
        sql1 &= "`sale_summary`.`受注番号`,"
        sql1 &= "`sale_details`.`購入品商品CODE`,"
        sql1 &= "`sale_details`.`購入品名`,"
        sql1 &= "`sale_details`.`購入品価格`,"
        sql1 &= "`sale_details`.`購入品数量`,"
        sql1 &= "`sale_summary`.`顧客分類`,"
        sql1 &= "`sale_summary`.`発送日`,"
        sql1 &= "`sale_summary`.`発送伝票番号`"
        sql1 &= " FROM `sale_summary` INNER JOIN `sale_details`"
        sql1 &= " On `sale_summary`.`受注番号` = `sale_details`.`受注番号`"
        sql1 &= " WHERE `sale_details`.`Data_source` = 'crm'"
        sql1 &= " AND `sale_summary`.`Data_source`='crm'"
        sql1 &= " AND `発送日` > '" & Form1.Cf_TextBox100.Text & "'"
        sql1 &= " AND `発送日` <= '" & Form1.Cf_TextBox101.Text & "'"
        sql1 &= ";"

        Dim rst1 = sql_regist1(sql1.ToString)
        If rst1 <> "Complete" Then
            MsgBox("初期データが正常に反映されませんでした", MsgBoxStyle.Critical Or MsgBoxStyle.OkOnly, "警告")
            Debug.Print(sql1.ToString)
        End If

        ''データに含めることになったため、除去ブロックを削除(2018/2/22)
        'sql1 = "DELETE FROM `global_fare_details`"
        'sql1 &= " WHERE `受注番号` IN (SELECT FOO.`受注番号` FROM("
        'sql1 &= "SELECT `受注番号` FROM `global_fare_details` GROUP BY `受注番号` HAVING Count(`Serial`) > 1"
        'sql1 &= ") AS FOO );"

        'rst1 = sql_regist1(sql1.ToString)
        'If rst1 <> "Complete" Then
        'MsgBox("合わせ買い商品が正常に削除されませんでした", MsgBoxStyle.Critical Or MsgBoxStyle.OkOnly, "警告")
        'Debug.Print(sql1.ToString)
        'End If

        'sql1 = "DELETE FROM `global_fare_details`"
        'sql1 &= " WHERE `購入品数量` > 1"
        'sql1 &= ";"

        'rst1 = sql_regist1(sql1.ToString)
        'If rst1 <> "Complete" Then
        'MsgBox("1個以上購入の商品が正常に削除されませんでした", MsgBoxStyle.Critical Or MsgBoxStyle.OkOnly, "警告")
        'Debug.Print(sql1.ToString)
        'End If

        sql1 = "DELETE FROM `global_fare_details` WHERE `発送伝票番号` is null;"
        rst1 = sql_regist1(sql1.ToString)
        If rst1 <> "Complete" Then
            MsgBox("発送伝票番号の無い商品が正常に削除されませんでした", MsgBoxStyle.Critical Or MsgBoxStyle.OkOnly, "警告")
            Debug.Print(sql1.ToString)
        End If


        sql1 = "SELECT"
        sql1 &= " `地方公共団体コード`," '0
        sql1 &= " `都道府県`,"          '1
        sql1 &= " `佐川地域コード`,"       '2
        sql1 &= " `佐川地域名`"          '3
        sql1 &= " FROM `shipping_area_data`;"

        Dim dTb1 As DataTable = sql_result_return(sql1)

        If dTb1.Rows.Count = 0 Then
            MsgBox(dTb1.ToString & "データがありません。異常です", MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "警告")
            Debug.Print(sql1)
        Else
            For Each DRow As DataRow In dTb1.Rows

                Dim sql2 As String = "SELECT `お問い合せＮｏ`"
                sql2 &= " FROM `gcn_shipping_slip`"
                sql2 &= " WHERE `都道府県` Like '"
                sql2 &= DRow.Item(1)
                sql2 &= "%'"
                sql2 &= " AND `配送日` > '" & Form1.Cf_TextBox100.Text & "'"
                sql2 &= " AND `配送日` <= '" & Form1.Cf_TextBox101.Text & "'"
                sql2 &= ";"

                Dim sql3 As String = "UPDATE `global_fare_details` SET `都道府県`='"
                sql3 &= DRow.Item(1)
                sql3 &= "',"
                sql3 &= "`地方公共団体コード`='"
                sql3 &= DRow.Item(0)
                sql3 &= "',"
                sql3 &= "`佐川地域コード`='"
                sql3 &= DRow.Item(2)
                sql3 &= "',"
                sql3 &= "`佐川地域名`='"
                sql3 &= DRow.Item(3)
                sql3 &= "' WHERE `発送伝票番号` in ( "

                Dim dTb2 As DataTable = sql_result_return(sql2)

                If dTb2.Rows.Count = 0 Then
                    MsgBox(dTb2.ToString & "データがありません。異常です", MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "警告")
                    Debug.Print(sql2)
                Else
                    For Each DRow2 As DataRow In dTb2.Rows

                        sql3 &= "'"
                        sql3 &= DRow2.Item(0)
                        sql3 &= "',"

                    Next
                End If

                Dim i As Long = Len(sql3)
                sql3 = Left(sql3, i - 1)

                sql3 &= ");"

                rst1 = sql_regist1(sql3)
                If rst1 <> "Complete" Then
                    MsgBox("都道府県情報更新SQLが正常に反映されませんでした", MsgBoxStyle.Critical Or MsgBoxStyle.OkOnly, "警告")
                    Debug.Print(sql1.ToString)
                End If


            Next
        End If

        sql1 = "DELETE FROM `global_fare_details` WHERE `都道府県` is null;"
        rst1 = sql_regist1(sql1)
        If rst1 <> "Complete" Then
            MsgBox("都道府県情報の無い商品が正常に削除されませんでした", MsgBoxStyle.Critical Or MsgBoxStyle.OkOnly, "警告")
            Debug.Print(sql1.ToString)
        End If

        '送料の修正
        sql1 = "SELECT `合計`"
        sql1 &= " FROM `gcn_shipping_slip`"
        sql1 &= " WHERE `配送日` > '" & Form1.Cf_TextBox100.Text & "'"
        sql1 &= " AND `配送日` <= '" & Form1.Cf_TextBox101.Text & "'"
        sql1 &= " GROUP BY `合計`"
        sql1 &= ";"

        dTb1 = sql_result_return(sql1)
        If dTb1.Rows.Count = 0 Then
            MsgBox(dTb1.ToString & "データがありません。異常です", MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "警告")
            Debug.Print(sql1)
        Else
            For Each DRow As DataRow In dTb1.Rows

                Dim sql2 As String = "SELECT"
                sql2 &= " `お問い合せＮｏ`"
                sql2 &= " FROM `gcn_shipping_slip`"
                sql2 &= " WHERE `運賃` = "
                sql2 &= DRow.Item(0)

                sql2 &= " AND `配送日` > '" & Form1.Cf_TextBox100.Text & "'"
                sql2 &= " AND `配送日` <= '" & Form1.Cf_TextBox101.Text & "'"

                sql2 &= ";"

                Dim sql3 As String = "UPDATE `global_fare_details` SET `合計`='"
                sql3 &= DRow.Item(0)
                sql3 &= "' WHERE `発送伝票番号` in ( "

                Dim dTb2 As DataTable = sql_result_return(sql2)

                If dTb2.Rows.Count = 0 Then
                    MsgBox(dTb2.ToString & "データがありません。異常です", MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "警告")
                    Debug.Print(sql2)
                Else
                    For Each DRow2 As DataRow In dTb2.Rows

                        sql3 &= "'"
                        sql3 &= DRow2.Item(0)
                        sql3 &= "',"

                    Next
                End If

                Dim i2 As Long = Len(sql3)
                sql3 = Left(sql3, i2 - 1)
                sql3 &= ");"

                rst1 = sql_regist1(sql3)
                If rst1 <> "Complete" Then
                    MsgBox("送料合計が正常に反映されませんでした", MsgBoxStyle.Critical Or MsgBoxStyle.OkOnly, "警告")
                    Debug.Print(sql1.ToString)
                End If
            Next

            sql1 = "SELECT `購入品商品CODE` FROM `global_fare_details`"
            sql1 &= " WHERE `global_fare_details`.`発送日` > '" & Form1.Cf_TextBox100.Text & "'"
            sql1 &= " AND `global_fare_details`.`発送日` <= '" & Form1.Cf_TextBox101.Text & "'"
            sql1 &= " GROUP BY `購入品商品CODE`;"

            dTb1 = sql_result_return(sql1)

            If dTb1.Rows.Count = 0 Then
                MsgBox(dTb1.ToString & "データがありません。異常です", MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "警告")
                Debug.Print(sql1)
            Else
                For Each DRow As DataRow In dTb1.Rows

                    Dim sql2 As String = "SELECT"
                    sql2 &= " `商品ページコード`"
                    sql2 &= " FROM `Conve_data_total`"
                    sql2 &= " WHERE `枝付` = "
                    sql2 &= "'"
            sql2 &= DRow.Item(0)
                    sql2 &= "'"
                    sql2 &= ";"

                    Dim dTb2 As DataTable = sql_result_return(sql2)

                    If dTb2.Rows.Count = 0 Then
                        MsgBox(dTb2.ToString & "データがありません。異常です", MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "警告")
                        Debug.Print(sql2)
                    Else
                        For Each DRow2 As DataRow In dTb2.Rows

                            Dim sql3 As String = "UPDATE `global_fare_details` SET `商品ページ`='"
                            sql3 &= DRow2.Item(0)
                            sql3 &= "' WHERE `購入品商品CODE` = "
                            sql3 &= "'"
                            sql3 &= DRow.Item(0)
                            sql3 &= "'"
                            sql3 &= ";"

                            rst1 = sql_regist1(sql3)
                            If rst1 <> "Complete" Then
                                MsgBox("商品ページ情報が正常に反映されませんでした", MsgBoxStyle.Critical Or MsgBoxStyle.OkOnly, "警告")
                                Debug.Print(sql1.ToString)
                            End If


                        Next
                    End If

                Next
            End If


        End If

        Dim sno As Long = 1

        sql1 = "SELECT `商品ページ` FROM `global_fare_details`"
        sql1 &= " GROUP BY `商品ページ` ORDER BY `商品ページ`;"
        dTb1 = sql_result_return(sql1)

        If dTb1.Rows.Count = 0 Then
            MsgBox(dTb1.ToString & "データがありません。異常です", MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "警告")
            Debug.Print(sql1)
        Else
            For Each DRow As DataRow In dTb1.Rows

                Dim sql3 As String = "UPDATE  `global_fare_details` SET `Sort`="
                sql3 &= sno
                sql3 &= " WHERE `商品ページ` = "
                sql3 &= "'"
                sql3 &= DRow.Item(0)
                sql3 &= "'"
                sql3 &= ";"

                rst1 = sql_regist1(sql3)
                If rst1 <> "Complete" Then
                    MsgBox("ソート用番号が正常に反映されませんでした", MsgBoxStyle.Critical Or MsgBoxStyle.OkOnly, "警告")
                    Debug.Print(sql1.ToString)
                End If

                sno += 1

            Next
        End If

        sql1 = "SELECT"
        sql1 &= " `Serial`"
        sql1 &= ",`発送伝票番号`"
        sql1 &= " FROM `global_fare_details`"
        sql1 &= " WHERE `global_fare_details`.`発送日` > '" & Form1.Cf_TextBox100.Text & "'"
        sql1 &= " AND `global_fare_details`.`発送日` <= '" & Form1.Cf_TextBox101.Text & "'"

        dTb1 = sql_result_return(sql1)

        If dTb1.Rows.Count = 0 Then
            MsgBox(dTb1.ToString & "データがありません。異常です", MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "警告")
            Debug.Print(sql1)
        Else
            For Each DRow As DataRow In dTb1.Rows
                Dim sql2 As String = "SELECT"
                sql2 &= " `配送区分`"
                sql2 &= " FROM `gcn_shipping_slip`"
                sql2 &= " WHERE `お問い合せＮｏ` = '"
                sql2 &= DRow.Item(1)
                sql2 &= "'"
                sql2 &= " AND `配送日` > '" & Form1.Cf_TextBox100.Text & "'"
                sql2 &= " AND `配送日` <= '" & Form1.Cf_TextBox101.Text & "'"
                sql2 &= ";"

                Dim dTb2 As DataTable = sql_result_return(sql2)

                If dTb2.Rows.Count = 0 Then
                    MsgBox(dTb2.ToString & "データがありません。異常です", MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "警告")
                    Debug.Print(sql2)
                Else
                    For Each DRow2 As DataRow In dTb2.Rows
                        Dim sql3 As String = "UPDATE `global_fare_details` SET `配送区分`='"
                        sql3 &= DRow2.Item(0)
                        sql3 &= "' WHERE `Serial` = "
                        sql3 &= DRow.Item(0)
                        sql3 &= ";"

                        rst1 = sql_regist1(sql3)
                        If rst1 <> "Complete" Then
                            MsgBox("配送区分正常に反映されませんでした", MsgBoxStyle.Critical Or MsgBoxStyle.OkOnly, "警告")
                            Debug.Print(sql1.ToString)
                        End If
                    Next
                End If

            Next
        End If

        ''データベース切断
        Call sql_cl()

        MessageBox.Show("集計／計算処理完了", "完了", MessageBoxButtons.OK, MessageBoxIcon.Information)

    End Sub


    Function sql_string_type(ByVal sheet1 As ISheet, ByVal r1 As Integer, c1 As Integer) As String

        Dim dty1 As String = ""

        Try
            dty1 = sheet1.GetRow(r1).GetCell(c1).ToString()

            Return "'" & dty1 & "'"

        Catch ex As System.NullReferenceException
            Return "NULL"

        End Try

    End Function

    Function sql_numeric_type(ByVal sheet1 As ISheet, ByVal r1 As Integer, c1 As Integer) As String
        Dim dty1 As String = ""

        Try
            dty1 = sheet1.GetRow(r1).GetCell(c1).ToString()

            Return dty1

        Catch ex As System.NullReferenceException
            Return "NULL"

        End Try

    End Function

    Function sql_regist1(ByVal sql0 As String) As String

        Dim ln1 As Long = sql0.Length
        sql0 = sql0.Substring(0, ln1 - 1)
        sql0 &= ";"

        Return sql_result_no(sql0)

    End Function

End Module
