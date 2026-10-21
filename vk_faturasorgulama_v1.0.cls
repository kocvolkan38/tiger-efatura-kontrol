VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sayfa3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub CommandButton1_Click()
    Dim conn As Object
    Dim rs As Object
    Dim sql As String
    Dim row As Integer
    Dim sunucu As String, veritabani As String, kullanici As String, sifre As String
    Dim isyeri As String
    row = 7

    'Baðlantý bilgileri Ayar sayfasýndan okunur (K2:K5)
    sunucu = Trim(CStr(Sayfa1.Range("K2").Value))
    veritabani = Trim(CStr(Sayfa1.Range("K3").Value))
    kullanici = Trim(CStr(Sayfa1.Range("K4").Value))
    sifre = CStr(Sayfa1.Range("K5").Value)
    If sunucu = "" Or veritabani = "" Or kullanici = "" Then
        MsgBox "Ayar sayfasýnda sunucu IP, veritabaný adý ve kullanýcý adý alanlarýný doldurunuz.", vbExclamation, "github.com/kocvolkan38"
        Exit Sub
    End If

    'Ýþyeri (BRANCH) filtresi Rapor sayfasý J3 hücresinden okunur; boþ ise tüm iþyerleri listelenir
    isyeri = Trim(CStr(Sayfa3.Cells(3, 10).Value))
    If isyeri <> "" And Not IsNumeric(isyeri) Then
        MsgBox "Ýþyeri numarasý sayýsal olmalýdýr (örn. 0, 1, 2).", vbExclamation, "github.com/kocvolkan38"
        Exit Sub
    End If

    Set conn = CreateObject("ADODB.Connection")
    Set rs = CreateObject("ADODB.Recordset")
    
    Range("B7:N10000").ClearContents
    Cells(7, 2).Select
    conn.Open "Provider=SQLOLEDB;Data Source=" & sunucu & ";Initial Catalog=" & veritabani & ";User Id=" & kullanici & ";Password=" & sifre & ";"

    'Logo tarafý
    sql = " set dateformat dmy "
    sql = sql + " SELECT F.DATE_, CASE F.TRCODE "
    sql = sql + " WHEN 1 THEN 'Mal Alým Faturasý' "
    sql = sql + " WHEN 3 THEN 'Toptan Satýþ Ýade' "
    sql = sql + " WHEN 4 THEN 'Alýnan Hizmet Ft.' "
    sql = sql + " WHEN 6 THEN 'Alým Ýade Ft.' "
    sql = sql + " WHEN 8 THEN 'Toptan Satýþ Ft.' "
    sql = sql + " WHEN 9 THEN 'Verilen Hizmet Ft.' "
    sql = sql + " WHEN 13 THEN 'Alýnan Fiyat Farký' "
    sql = sql + " WHEN 14 THEN 'Verilen Fiyat Farký' "
    sql = sql + " END AS FATURA_TURU, "
    sql = sql + " REPLACE(F.FICHENO, '/','0') AS FISNO, C.DEFINITION_, F.NETTOTAL, " '17/11 Ayný numaraya sahip faturalarda bir 0 silinerek / karakteri ekleniyor. Bu karakter için dönüþüm satýrý eklendi.
    'sql = sql + " F.FICHENO, C.DEFINITION_, F.NETTOTAL, " '17/11 Güncelleme öncesi
    sql = sql & " CASE F.ACCOUNTED "
    sql = sql & " WHEN 1 THEN 'Muhasebeleþmiþ' "
    sql = sql & " WHEN 0 THEN 'Muhasebeleþmemiþ' "
    sql = sql & " END AS MUHASEBE_DURUMU "
    sql = sql & " FROM LG_" & Sayfa3.Cells(5, 8).Value & "_01_INVOICE F "
    sql = sql + " LEFT JOIN LG_" & Sayfa3.Cells(5, 8).Value & "_CLCARD C ON C.LOGICALREF = F.CLIENTREF "
    sql = sql + " WHERE F.DATE_ >= '" & Sayfa3.Cells(2, 3).Value & "' "
    sql = sql + " AND F.DATE_ <= '" & Sayfa3.Cells(3, 3).Value & "' "
    If isyeri <> "" Then
        sql = sql + " AND F.BRANCH = " & CLng(isyeri) & " "
    End If
    sql = sql + " AND F.FROMBANK = 0 "
'    sql = sql + " AND F.DEPARTMENT <> '000' "
'    sql = sql + " AND F.DEPARTMENT <> '001' "
'    sql = sql + " AND F.DEPARTMENT <> '005' "
'    sql = sql + " AND F.DEPARTMENT <> '006' "
'    sql = sql + " AND F.DEPARTMENT <> '099' "
    If Sayfa3.Cells(2, 8) = 1 Then
        sql = sql + " AND ( F.TRCODE IN(1,4,2,3)) "
    ElseIf Sayfa3.Cells(2, 8) = 2 Then
        sql = sql + " AND ( F.TRCODE IN(7,8,9,14,6,13)) "
    End If
    If Sayfa3.Cells(3, 8) <> 3 Then
        sql = sql + " AND F.ACCOUNTED = '" & Sayfa3.Cells(3, 8).Value & "' "
    End If
    If Sayfa3.Cells(2, 5) <> "" Then
        sql = sql + " AND C.DEFINITION_ LIKE '%" & Sayfa3.Cells(2, 5).Value & "%' "
    End If
    If Sayfa3.Cells(3, 5) <> "" Then
        sql = sql + " AND F.FICHENO LIKE '%" & Sayfa3.Cells(3, 5).Value & "%' "
    End If
    sql = sql + " ORDER BY F.DATE_ "

    rs.Open sql, conn
    
    Range("B7").CopyFromRecordset rs
    rs.Close
    
    'e-fatura tarafý
    sql = " set dateformat dmy "
    sql = sql + " SELECT DOCDATE, DOCNR, SENDERTITLE, "
    sql = sql + " CASE DOCCURRENCY WHEN 'TRY' THEN DOCTOTAL ELSE DOCTOTALLOCAL END AS TOPLAM,"
    sql = sql + " CASE STATUS WHEN 10 THEN 'RET' END AS RET "
    'sql = sql + " , STATUS, RECTYPE, TRCODE, EINVOICETYP "
    sql = sql + " FROM LG_" & Sayfa3.Cells(5, 8).Value & "_01_APPROVAL "
    sql = sql + " WHERE DOCDATE >= '" & Sayfa3.Cells(2, 3).Value & "' "
    sql = sql + " AND DOCDATE <= '" & Sayfa3.Cells(3, 3).Value & "' "
    sql = sql + " AND DOCNR NOT LIKE '' "
    sql = sql + " AND RECTYPE NOT LIKE 114 "
    sql = sql + " AND RECTYPE NOT LIKE 115 "
    sql = sql + " AND RECTYPE NOT LIKE 125 "
    sql = sql + " AND RECTYPE NOT LIKE 127 "
    
    If Sayfa3.Cells(2, 5) <> "" Then
        sql = sql + " AND SENDERTITLE LIKE '%" & Sayfa3.Cells(2, 5).Value & "%' "
    End If
    If Sayfa3.Cells(3, 5) <> "" Then
        sql = sql + " AND DOCNR LIKE '%" & Sayfa3.Cells(3, 5).Value & "%' "
    End If
    sql = sql + " ORDER BY DOCDATE "
    
    rs.Open sql, conn
    
    Range("I7").CopyFromRecordset rs
    
    SON = Sayfa3.Range("J" & Rows.Count).End(xlUp).row
    For i = 7 To SON
        On Error Resume Next
        Sayfa3.Cells(i, 13).Value = Application.WorksheetFunction.VLookup(Sayfa3.Cells(i, 10).Value, Sayfa3.Range("D7:F1000"), 3, 0) - Sayfa3.Cells(i, 12).Value
        On Error GoTo 0
    Next i

    For i = 7 To SON
        On Error Resume Next
        Sayfa3.Cells(i, 14).Value = Application.WorksheetFunction.VLookup(Sayfa3.Cells(i, 10).Value, Sayfa3.Range("D7:G10000"), 4, 0)
        If Err.Number <> 0 Then
            Sayfa3.Cells(i, 14).Value = "Ýþlenmemiþ"
            Err.Clear
        End If
        On Error GoTo 0
    Next i
    
    rs.Close
    conn.Close
    Set rs = Nothing
    Set conn = Nothing
    MsgBox "Faturalar listelendi!", vbOKOnly, "github.com/kocvolkan38"
End Sub
