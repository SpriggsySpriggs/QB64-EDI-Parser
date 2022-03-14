Option Explicit

'$VersionInfo metacommands used to enable comctl32 v6
$VersionInfo:CompanyName=SpriggsySpriggs
$VersionInfo:FileDescription=Testing EDI parsing
$VersionInfo:Comments=Testing EDI parsing
$VersionInfo:LegalCopyright=(c)2022 SpriggsySpriggs
$VersionInfo:ProductName=QB64 EDI Parser
$VersionInfo:InternalName=parseEDI
$VersionInfo:PRODUCTVERSION#=0,0,0,2
$VersionInfo:FILEVERSION#=0,0,0,2
$VersionInfo:ProductVersion=0.0.0.2
$VersionInfo:FileVersion=0.0.0.2
$VersionInfo:OriginalFilename=parseEDI

$If VERSION < 2.1 OR WINDOWS = 0 OR 32BIT Then
    $Error Requires Windows 64 bit and QB64 version 2.1 or higher
$End If

On Error GoTo ERRHANDLE
ERRHANDLE:
If Err Then
    If InclErrorLine Then
        CriticalError Err, InclErrorLine, InclErrorFile$
        Resume Next
    Else
        CriticalError Err, ErrorLine, ""
        Resume Next
    End If
End If

$NoPrefix
$Console:Only

$ExeIcon:'edi_icon.ico'
Icon

ConsoleTitle "EDI Parser"


'Type AddressInfo
'    As String ContactName, CompanyName, Street, City, State, Zip
'End Type

'Type Header
'    As String DocType, Version, Sender, Recipient, SendDate, SendTime
'    As AddressInfo AddressInfo
'End Type

Type LineItem830
    'As Header Header
    As String PO, PartNumber, PartDescription, ShipQty, ShipDates, ShipTerms
End Type

'$Include:'OpenSave.BI'

Const DEFEXT = "edi"

ReDim Shared As String EDI(0)
Dim As String EDI
Dim Shared As String delimiter, terminator
Dim As Long hEDI: hEDI = FreeFile
Dim As String file: file = ComDlgFileName("Select EDI Document", Dir$("documents"), "EDI Documents|*.EDI|All Files|*.*", 1, OFN_HIDEREADONLY)

If file <> "" Then
    If checkExtension(file) Then
        Open "B", hEDI, file
        EDI = Space$(LOF(hEDI))
        Get hEDI, , EDI
        Close hEDI
        ConsoleTitle "EDI Parser: Viewing - " + Mid$(file, InStrRev(file, "\") + 1)
    Else
        SoftError "Parse failed", "Extension is not valid"
        System
    End If
Else End
End If

If InStr(EDI, "ISA") <> 0 And InStr(EDI, "GS") <> 0 Then
    delimiter = Mid$(EDI, InStr(EDI, "ISA") + 3, 1)
    terminator = Mid$(EDI, InStr(EDI, "GS" + delimiter) - 1, 1)
Else
    SoftError "Parse failed", "A valid header could not be found"
    System
End If

denullify EDI, delimiter
If InStr(EDI, Chr$(10)) Then
    tokenize EDI, Chr$(13) + Chr$(10) + terminator, EDI()
Else
    tokenize EDI, terminator, EDI() 'splitting EDI into an array using the newline characters as the delimiters
End If

ReDim Shared As String ISA(0), GS(0), ST(0), N1(0), N2(0), N3(0)

tokenize EDI(findNext("ISA", EDI(), 0)), delimiter + terminator, ISA()
tokenize EDI(findNext("GS", EDI(), 0)), delimiter + terminator, GS()
tokenize EDI(findNext("ST", EDI(), 0)), delimiter + terminator, ST()

If isKnownDoc(ST(1)) And isKnownVersion(ISA(12)) Then
    If findNext("N1", EDI(), 0) And findNext("N2", EDI(), 0) And findNext("N3", EDI(), 0) Then
        tokenize EDI(findNext("N1", EDI(), 0)), delimiter + terminator, N1()
        tokenize EDI(findNext("N2", EDI(), 0)), delimiter + terminator, N2()
        tokenize EDI(findNext("N3", EDI(), 0)), delimiter + terminator, N3()
    End If
    doHeader
    Select Case ST(1)
        Case "830"
            ReDim Shared As LineItem830 LineItem(0)
            do830
    End Select
Else
    SoftError "Parse failed", "Document type is not in list of known types." + Chr$(10) + "Type: " + ST(1) + Chr$(10) + "Version: " + ISA(12)
    System
End If

'Function countElement~& (element As String, EDI() As String)
'    Dim As Unsigned Long i
'    For i = LBound(EDI) To UBound(EDI)

'    Next
'End Function

Function findNext~& (element As String, EDI() As String, last As Unsigned Long)
    Dim As Unsigned Long i
    For i = last To UBound(EDI)
        If Mid$(EDI(i), InStr(EDI(i), element + delimiter), Len(element)) = element Then
            findNext = i
            Exit Function
        End If
    Next
End Function

Sub doHeader
    Print "Doc Type : "; ST(1); " - "; docFriendlyName(ST(1))
    Print "Version  : "; ISA(12)
    Print "Sender   : "; ISA(6)
    Print "Recipient: "; ISA(8)
    Print "Send Date: "; ISA(9)
    Print "Send Time: "; ISA(10)
    If ST(1) = "830" Then
        Print "Address  : "; N2(1)
        Print , "  "; N2(2)
        Print , "  "; N3(1)
        Print , "  "; N3(2)
    End If
    Print String$(100, "-")
End Sub

Sub do830
    ReDim As LineItem830 LineItem(0)
    Dim As Unsigned Long linecount, nextitem, olditem, lastitem
    Do
        nextitem = findNext("LIN", EDI(), olditem + 1)
        If nextitem Then
            ReDim Preserve As LineItem830 LineItem(UBound(LineItem) + 1)
            olditem = nextitem
            ReDim As String LIN(0)
            tokenize EDI(nextitem), delimiter + terminator, LIN()
            Dim As Integer i
            LineItem(linecount).PartNumber = LIN(3)
            For i = LBound(Lin) To UBound(Lin)
                Select Case LIN(i)
                    Case "PD"
                        LineItem(linecount).PartDescription = LIN(i + 1)
                    Case "PO"
                        LineItem(linecount).PO = LIN(i + 1)
                End Select
            Next
            Print "Part Number: "; LineItem(linecount).PartNumber, "PO Number: "; LineItem(linecount).PO, "Description: "; LineItem(linecount).PartDescription
            Print String$(100, "-")
            lastitem = olditem + 1
            ReDim As String FST(0)
            nextitem = findNext("FST", EDI(), olditem)
            tokenize EDI(nextitem), delimiter + terminator, FST()
            For i = LBound(FST) + 1 To UBound(FST)
                If i = UBound(FST) Then Print FST(i) Else Print FST(i),
            Next
            LineItem(linecount).ShipQty = LineItem(linecount).ShipQty + FST(1) + Chr$(3)
            LineItem(linecount).ShipDates = LineItem(linecount).ShipDates + FST(4) + Chr$(3)
            LineItem(linecount).ShipTerms = LineItem(linecount).ShipTerms + FST(2) + FST(3) + Chr$(3)
            Do
                olditem = nextitem
                nextitem = findNext("FST", EDI(), olditem + 1)
                If nextitem <> olditem + 1 Then Exit Do
                ReDim As String FST(0)
                tokenize EDI(nextitem), delimiter + terminator, FST()
                For i = LBound(FST) + 1 To UBound(FST)
                    If i = UBound(FST) Then Print FST(i) Else Print FST(i),
                Next
                LineItem(linecount).ShipQty = LineItem(linecount).ShipQty + FST(1) + Chr$(3)
                LineItem(linecount).ShipDates = LineItem(linecount).ShipDates + FST(4) + Chr$(3)
                LineItem(linecount).ShipTerms = LineItem(linecount).ShipTerms + FST(2) + FST(3) + Chr$(3)
            Loop While nextitem = olditem + 1
            Print String$(100, "-")
        End If
        linecount = linecount + 1
    Loop While findNext("LIN", EDI(), lastitem + 1)
    ReDim Preserve As LineItem830 LineItem(UBound(LineItem) - 1)
End Sub

Sub loadData (DataName As String, DataArray() As String)
    '$Include:'EDI_Data.BI'
    Dim As Unsigned Long i, lowerbound
    lowerbound = LBound(DataArray)
    i = lowerbound
    Dim As String EDI
    Select Case UCase$(DataName)
        Case "SEPARATORS"
            Restore Separators
        Case "HEADER"
            Restore Header
        Case "TRAILER"
            Restore Trailer
        Case "DETAILSEGMENTS"
            Restore DetailSegments
        Case "LOOPS"
            Restore Loops
        Case "DOCUMENTS"
            Restore Documents
        Case "VERSIONS"
            Restore Versions
        Case "DOCUMENTTYPES"
            Restore DocumentTypes
    End Select

    Read EDI
    While EDI <> "EOD"
        ReDim Preserve DataArray(lowerbound To UBound(DataArray) + 1)
        DataArray(i) = EDI
        i = i + 1
        Read EDI
    Wend
    ReDim Preserve DataArray(UBound(DataArray) - 1)
End Sub

Function isKnownDoc%% (DocType As String)
    ReDim As String Docs(0)
    loadData "Documents", Docs()
    Dim As Long i
    For i = LBound(Docs) To UBound(Docs)
        If Docs(i) = DocType Then
            isKnownDoc = -1
            Exit Function
        End If
    Next
End Function

Function isKnownVersion%% (Version As String)
    ReDim As String Versions(0)
    loadData "Versions", Versions()
    Dim As Long i
    For i = LBound(Versions) To UBound(Versions)
        If Versions(i) = Version Then
            isKnownVersion = -1
            Exit Function
        End If
    Next
End Function

Function docFriendlyName$ (DocType As String)
    ReDim As String DocumentTypes(0)
    loadData "DocumentTypes", DocumentTypes()
    Dim As Long i
    For i = LBound(DocumentTypes) To UBound(DocumentTypes)
        If Left$(DocumentTypes(i), 3) = DocType Then
            docFriendlyName = Mid$(DocumentTypes(i), 5)
            Exit Function
        End If
    Next
End Function

$If PTRTOSTRING = UNDEFINED Then
    $Let PTRTOSTRING = TRUE
    Function pointerToString$ (pointer As Offset)
        Declare CustomType Library
            Function strlen%& (ByVal ptr As Unsigned Offset)
        End Declare
        Dim As Offset length: length = strlen(pointer)
        If length Then
            Dim As MEM pString: pString = Mem(pointer, length)
            Dim As String ret: ret = Space$(length)
            MemGet pString, pString.OFFSET, ret
            MemFree pString
        End If
        pointerToString = ret
    End Function
$End If

Sub tokenize (toTokenize As String, delimiters As String, StorageArray() As String)
    Declare CustomType Library
        Function strtok%& (ByVal str As Offset, delimiters As String)
    End Declare
    Dim As Offset tokenized
    Dim As String tokCopy: tokCopy = toTokenize + Chr$(0)
    Dim As String delCopy: delCopy = delimiters + Chr$(0)
    Dim As Unsigned Long lowerbound: lowerbound = LBound(StorageArray)
    Dim As Unsigned Long i: i = lowerbound
    tokenized = strtok(Offset(tokCopy), delCopy)
    While tokenized <> 0
        ReDim Preserve StorageArray(lowerbound To UBound(StorageArray) + 1)
        StorageArray(i) = pointerToString(tokenized)
        tokenized = strtok(0, delCopy)
        i = i + 1
    Wend
    ReDim Preserve StorageArray(UBound(StorageArray) - 1)
End Sub

Sub denullify (arg As String, delimiter As String)
    Do
        arg = String.Insert(arg, " ", InStr(arg, delimiter + delimiter) + 1)
    Loop While InStr(arg, delimiter + delimiter)
End Sub

Function String.Insert$ (toChange As String, insert As String, position As Long)
    Dim As String newchange
    newchange = toChange
    newchange = Mid$(newchange, 1, position - 1) + insert + Mid$(newchange, position, Len(newchange) - position + 1)
    String.Insert = newchange
End Function

Function checkExtension%% (file As String)
    If UCase$(Mid$(file, InStrRev(file, "."), 4)) = ".EDI" Then checkExtension = -1 Else checkExtension = 0
End Function

Function String.StartsWith (check As String, contains As String)
    If Mid$(check, 1, Len(contains)) = contains Then String.StartsWith = -1 Else String.StartsWith = 0
End Function

Function String.Replace$ (a As String, b As String, c As String)
    Dim j, r$
    j = InStr(a, b)
    If j > 0 Then
        r$ = Left$(a, j - 1) + c + String.Replace(Right$(a, Len(a) - j + 1 - Len(b)), b, c)
    Else
        r$ = a
    End If
    String.Replace = r$
End Function

'$Include:'OpenSave.BM'
'$Include:'MessageBox.BM'
