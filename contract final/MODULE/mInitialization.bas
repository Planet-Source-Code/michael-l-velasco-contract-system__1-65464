Attribute VB_Name = "mInitialization"
Global UserName
Global Rights1_Add
Global Rights1_Edit
Global Rights1_Delete
Global Rights2_Tables
Global Rights2_Service_Crew
Global Rights2_Ingredients
Global Rights2_Menu
Global Rights2_Supplier
Global Rights2_SalesOrders
Global Rights2_PurchaseOrders
Global Rights2_ReceivingOrders
Global Rights2_Post_SalesOrders
Global Rights2_Post_ReceivingOrders
Global Rights2_Inventory_Report
Global Rights2_Sales_Report
Global Rights2_Critical_Report
Global Rights3_Backup
Global Rights3_Restore
Global companycode As String
Global pili As String
Global Rights3_Password_Security
Global Rights3_CarwtConse
Global OrderEntryOpen As Boolean
Global OrderEntryModule As String
Global EditClicked As Boolean
Global xWidth As Integer
Global xHeight As Integer
Global t1 As String
Global t2 As String
Global t3 As String
Global t4 As String
Global t5 As String
Global t6 As String
Global t7 As String
Global t8 As String
Global t9 As String
Global t10 As String
Global t11 As String
Global t12 As String
Global t13 As String
Global t14 As String
Global t15 As String
Global t16 As String
Global t17 As String
Global t18 As String
Global t19 As String
Global t20 As String
Global t21 As String
Global t22 As String
Global t23 As String
Global t24 As String
Global t25 As String
Global t26 As String
Global secdep As String
Global adrent As String
Global strDB As String
Global FileNameTXT As String
Global dd As String
Public con As New ADODB.Connection
Global CRReport1 As CRAXDRT.Report
Global CRApp As New CRAXDRT.Application
Global db As New Connection

Public Declare Sub GlobalMemoryStatus Lib "kernel32" (lpBuffer As meminfo_status)

Public Type meminfo_status
    dwlength As Long
    dwmemoryload As Long
    dwtotalphy As Long
    dwavaiphy As Long
    dwtotalpagefile As Long
    dwavaipagefile As Long
    dwtotalvirtual As Long
    dwavailabelvirtual As Long
End Type

Public meminfo As meminfo_status

' REPORT RESOLUTION FIXER
Function ReportResolution()
    If IsResolution(640, 480) Then
        xWidth = 640
        xHeight = 480
    ElseIf IsResolution(800, 600) Then
        xWidth = 800
        xHeight = 600
    ElseIf IsResolution(1024, 768) Then
        xWidth = 1024
        xHeight = 768
    ElseIf IsResolution(1280, 1024) Then
        xWidth = 1280
        xHeight = 1024
    ElseIf IsResolution(1600, 1200) Then
        xWidth = 1600
        xHeight = 1200
    End If
End Function

' FOR RESOLUTION VERIFIER
Function IsResolution(Width As Integer, Height As Integer) As Boolean
    If (Screen.Width / Screen.TwipsPerPixelX = Width) And (Screen.Height / Screen.TwipsPerPixelY = Height) Then
        IsResolution = True
    Else
        IsResolution = False
    End If
End Function

' DECODE PASSWORD.
Function Decode_Pass(p_str As String) As String
    For i = 1 To Len(p_str) Step 1
        strs = strs + Chr(Asc(Mid(p_str, i, 1)) * 2)
    Next i
        Decode_Pass = strs
End Function

' UNCODE PASSWORD.
Function UnCode_Pass(p_str As String) As String
    For i = 1 To Len(p_str) Step 1
        strs = strs + Chr(Asc(Mid(p_str, i, 1)) / 2)
    Next i
        UnCode_Pass = strs
End Function
Sub Main()
FileNameTXT = App.Path + "\contract"
strDB = FileNameTXT + ";Jet OLEDB:Database Password=mykpogi;"

    Set db = New Connection
        db.CursorLocation = adUseClient
        db.Open "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=" & strDB
Login.Show
End Sub

Function Amt2Words(nInAmount As Double) As String
    Dim sInWords As String, sNum As String, nCent As Double, sThree As String
    Dim sNum1 As String, nCtr As Integer, sWord As String, lcont As Boolean
    
    Dim aTens(9) As String, aOnes(9) As String, aCValue(9) As String
    Dim nLen As Integer, x As Integer, nSingle As Integer
    
    
    aOnes(1) = "One"
    aOnes(2) = "Two"
    aOnes(3) = "Three"
    aOnes(4) = "Four"
    aOnes(5) = "Five"
    aOnes(6) = "Six"
    aOnes(7) = "Seven"
    aOnes(8) = "Eight"
    aOnes(9) = "Nine"

    aTens(1) = "Ten"
    aTens(2) = "Twenty"
    aTens(3) = "Thirty"
    aTens(4) = "Forty"
    aTens(5) = "Fifty"
    aTens(6) = "Sixty"
    aTens(7) = "Seventy"
    aTens(8) = "Eigthy"
    aTens(9) = "Ninety"

    aCValue(1) = "Eleven"
    aCValue(2) = "Twelve"
    aCValue(3) = "Thirteen"
    aCValue(4) = "Fourteen"
    aCValue(5) = "Fifteen"
    aCValue(6) = "Sixteen"
    aCValue(7) = "Seventeen"
    aCValue(8) = "Eighteen"
    aCValue(9) = "Nineteen"
    
    nInAmount = Abs(nInAmount)
    sNum = Trim(Str(Int(nInAmount)))
    nCent = 0
    If Val(sNum) > 0 Then
        nCent = nInAmount - Val(sNum)
    Else
        nCent = nInAmount
    End If

    nCent = nCent * 100
    nLen = Len(sNum)
    If nLen < 12 Then
        sNum1 = Stuff(sNum, 1, "0", 12 - Len(sNum))
    Else
        sNum1 = sNum
    End If
    sInWords = ""
    
    nCtr = 1
    Do While True
        sThree = Mid(sNum1, nCtr, 3)
        sWord = ""
        For x = 1 To 3
            nSingle = Val(Mid(sThree, x, 1))
            lcont = True
            If nSingle > 0 Then
                If x = 1 Then
                    sWord = sWord + aOnes(nSingle) + " Hundred "
                End If
                If x = 2 Then
                    If nSingle = 1 And Val(Mid(sThree, 3, 1)) > 0 Then
                        sWord = sWord + " " + aCValue(Val(Mid(sThree, 3, 1)))
                        lcont = False
                    Else
                        If nSingle > 0 Then
                            sWord = sWord + " " + aTens(nSingle)
                        End If
                    End If
                End If
            
                If Not lcont Then
                    Exit For
                End If
                If x = 3 Then
                    sWord = sWord + " " + aOnes(nSingle)
                End If
            End If
        Next x
    
        sInWords = sInWords + " " + sWord
        If nCtr = 1 And Len(Trim(sInWords)) > 1 Then
            sInWords = sInWords + " " + "Billion"
        End If
    
        If nCtr = 4 And Len(Trim(sInWords)) > 1 Then
            sInWords = sInWords + " " + "Million"
        End If
    
        If nCtr = 7 And Len(Trim(sInWords)) > 1 Then
            sInWords = sInWords & " " & "Thousand"
        End If
    
        nCtr = nCtr + 3
        If nCtr > 13 Then
            Exit Do
        End If
    
    Loop
    
    'I use Peso coz its our currency name in the Philippines
    'Just change it whatever currency word you have...
    If pili = "ITO" Then
            If Val(sNum) > 1 Then
                sInWords = sInWords & ""
            End If
            
            If Val(sNum) = 1 Then
                sInWords = sInWords + ""
            End If
            
            nCent = Format(nCent, "0.00")
            
            If nCent > 0 And Val(sNum) > 1 Then
                sInWords = sInWords + " " + "and" + " " + Trim(Str(nCent))
            End If
        
            If nCent > 0 And Val(sNum) = 0 Then
                sInWords = sInWords + " " + Trim(Str(nCent))
            End If
            
            sInWords = sInWords + " "
            Amt2Words = Trim(sInWords)
            pili = ""
    Else
            If Val(sNum) > 1 Then
                sInWords = sInWords & "" & "Pesos"
            End If
            
            If Val(sNum) = 1 Then
                sInWords = sInWords + "" + "Pesos"
            End If
            
            nCent = Format(nCent, "0.00")
            
            If nCent > 0 And Val(sNum) > 1 Then
                sInWords = sInWords + " " + "and" + " " + Trim(Str(nCent)) + "/100"
            End If
        
            If nCent > 0 And Val(sNum) = 0 Then
                sInWords = sInWords + " " + Trim(Str(nCent)) + "/100"
            End If
            
            sInWords = sInWords + " " + "Only"
            Amt2Words = Trim(sInWords)
    End If
End Function

'Parameters: 1. sStr : String to be stuff
'            2. cPos : Position where it is inserted
'                      1 : Left
'                      2 : Right
'            3. cStuff: Character to be stuff
'            4. nNo   : how many times

Function Stuff(sStr, cPos As Byte, cStuff As String, nNo As Byte) As String

    Dim sString As String, x As Byte
    sString = ""
    For x = 1 To nNo
        sString = sString & cStuff
    Next x
    If cPos = 1 Then
        sString = sString & sStr
    End If
    
    If cPos = 2 Then
        sString = sStr & sString
    End If
    
    Stuff = sString
    
End Function
