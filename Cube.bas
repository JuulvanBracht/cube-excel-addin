Attribute VB_Name = "modCube"
Option Explicit

'=====================================================================
' Cube add-in
' Cube-add-in met Ribbon-acties:
'   Verversen    -> Selectie / Huidig werkblad / Volledige werkmap
'   Vastzetten   -> Selectie / Huidig werkblad / Volledige werkmap
'   Initialiseren -> Maakt de werkmap-brede LAMBDA-naam TPCUBE aan
'=====================================================================

Private Const TITEL As String = "Cube"
Private Const VRAAG_BEVESTIGING As Boolean = True   ' False = geen waarschuwing bij "Naar waarden"

Private mOudeBerekening As XlCalculation

'---------------------------------------------------------------------
' Ribbon callbacks (namen moeten overeenkomen met customUI14.xml)
'---------------------------------------------------------------------
Public Sub Cube_Refresh(control As IRibbonControl)
    Select Case control.ID
        Case "btnRefreshSelectie": RefreshSelectie
        Case "btnRefreshTabblad": RefreshTabblad
        Case "btnRefreshWerkbestand": RefreshWerkbestand
    End Select
End Sub

Public Sub Cube_NaarWaarden(control As IRibbonControl)
    Select Case control.ID
        Case "btnWaardeSelectie": WaardenSelectie
        Case "btnWaardeTabblad": WaardenTabblad
        Case "btnWaardeWerkbestand": WaardenWerkbestand
    End Select
End Sub


'---------------------------------------------------------------------
' INITIALISEREN
'---------------------------------------------------------------------
Public Sub Cube_Initialize(control As IRibbonControl)

    Dim wb As Workbook
    Dim nm As Name
    Dim naam As Name
    Dim lambdaNaam As String
    Dim lambdaDefinitie As String

    On Error GoTo Fout

    If ActiveWorkbook Is Nothing Then
        MsgBox "Er is geen actieve werkmap beschikbaar.", vbExclamation, TITEL
        Exit Sub
    End If

    Set wb = ActiveWorkbook

    ' Hardcoded LAMBDA-waarden
    lambdaNaam = "TPCUBE"
    lambdaDefinitie = "LAMBDA(RAND())"

    ' Zoek naar een bestaande naam met dezelfde naam (lettergrootte-onsgevoelig).
    For Each naam In wb.Names
        If StrComp(naam.Name, lambdaNaam, vbTextCompare) = 0 Then
            Set nm = naam
            Exit For
        End If
    Next naam

    If nm Is Nothing Then
        ' Maak de werkmap-brede LAMBDA-functie aan.
        Set nm = wb.Names.Add( _
            Name:=lambdaNaam, _
            RefersTo:="=" & lambdaDefinitie, _
            Visible:=True)
    Else
        ' Werk de bestaande werkmap-brede definitie bij.
        nm.RefersTo = "=" & lambdaDefinitie
    End If

    MsgBox "Cube is succesvol geïnitialiseerd." & vbNewLine & vbNewLine & _
           "Naam: " & lambdaNaam & vbNewLine & _
           "Definitie: =" & lambdaDefinitie & vbNewLine & _
           "Gebruik: =" & lambdaNaam & "()", _
           vbInformation, TITEL

    Exit Sub

Fout:
    MsgBox "Cube initialiseren mislukt:" & vbNewLine & _
           Err.Description, vbExclamation, TITEL

End Sub

'---------------------------------------------------------------------
' REFRESH
'---------------------------------------------------------------------
Private Sub RefreshSelectie()
    Dim rng As Range
    Set rng = GeselecteerdBereik()
    If rng Is Nothing Then Exit Sub

    On Error GoTo Fout
    Application.ScreenUpdating = False
    BerekenBereik rng
    Application.CalculateUntilAsyncQueriesDone
    Application.ScreenUpdating = True
    Exit Sub
Fout:
    Application.ScreenUpdating = True
    MsgBox "Refresh mislukt: " & Err.Description, vbExclamation, TITEL
End Sub

Private Sub RefreshTabblad()
    Dim ws As Worksheet
    Set ws = ActiefBlad()
    If ws Is Nothing Then Exit Sub

    On Error GoTo Fout
    Application.ScreenUpdating = False
    BerekenBereik ws.UsedRange
    Application.CalculateUntilAsyncQueriesDone
    Application.ScreenUpdating = True
    Exit Sub
Fout:
    Application.ScreenUpdating = True
    MsgBox "Refresh mislukt: " & Err.Description, vbExclamation, TITEL
End Sub

Private Sub RefreshWerkbestand()
    Dim wb As Workbook, ws As Worksheet
    If ActiveWorkbook Is Nothing Then Exit Sub
    Set wb = ActiveWorkbook

    On Error GoTo Fout
    Application.ScreenUpdating = False
    wb.RefreshAll                                   ' dataverbindingen, draaitabellen, cube
    Application.CalculateUntilAsyncQueriesDone
    For Each ws In wb.Worksheets
        BerekenBereik ws.UsedRange
    Next ws
    Application.CalculateUntilAsyncQueriesDone
    Application.ScreenUpdating = True
    Exit Sub
Fout:
    Application.ScreenUpdating = True
    MsgBox "Refresh mislukt: " & Err.Description, vbExclamation, TITEL
End Sub

Private Sub BerekenBereik(ByVal rng As Range)
    Dim doel As Range
    Set doel = Application.Intersect(rng, rng.Worksheet.UsedRange)
    If doel Is Nothing Then Exit Sub
    doel.Dirty
    doel.Calculate
End Sub

'---------------------------------------------------------------------
' NAAR WAARDEN (formules -> platte tekst/getallen)
'---------------------------------------------------------------------
Private Sub WaardenSelectie()
    Dim rng As Range
    Set rng = GeselecteerdBereik()
    If rng Is Nothing Then Exit Sub
    If Not Bevestig("de selectie") Then Exit Sub

    Dim n As Long, geblokkeerd As Long
    On Error GoTo Fout
    StartBatch
    Application.CalculateUntilAsyncQueriesDone      ' eerst laten uitrekenen, anders wordt #GETTING_DATA vastgezet
    n = ZetBladVast(rng.Worksheet, rng)
    If n < 0 Then geblokkeerd = 1: n = 0
    EindBatch
    Klaar n, geblokkeerd
    Exit Sub
Fout:
    EindBatch
    MsgBox "Omzetten mislukt: " & Err.Description, vbExclamation, TITEL
End Sub

Private Sub WaardenTabblad()
    Dim ws As Worksheet
    Set ws = ActiefBlad()
    If ws Is Nothing Then Exit Sub
    If Not Bevestig("het tabblad '" & ws.Name & "'") Then Exit Sub

    Dim n As Long, geblokkeerd As Long
    On Error GoTo Fout
    StartBatch
    Application.CalculateUntilAsyncQueriesDone
    n = ZetBladVast(ws)
    If n < 0 Then geblokkeerd = 1: n = 0
    EindBatch
    Klaar n, geblokkeerd
    Exit Sub
Fout:
    EindBatch
    MsgBox "Omzetten mislukt: " & Err.Description, vbExclamation, TITEL
End Sub

Private Sub WaardenWerkbestand()
    Dim ws As Worksheet, n As Long, r As Long, geblokkeerd As Long
    If ActiveWorkbook Is Nothing Then Exit Sub
    If Not Bevestig("het hele werkbestand") Then Exit Sub

    On Error GoTo Fout
    StartBatch
    Application.CalculateUntilAsyncQueriesDone
    For Each ws In ActiveWorkbook.Worksheets
        r = ZetBladVast(ws)
        If r < 0 Then geblokkeerd = geblokkeerd + 1 Else n = n + r
    Next ws
    EindBatch
    Klaar n, geblokkeerd
    Exit Sub
Fout:
    EindBatch
    MsgBox "Omzetten mislukt: " & Err.Description, vbExclamation, TITEL
End Sub

' Zet alle formules in 'bereik' (of het hele blad) om naar waarden.
' Geeft het aantal formulecellen terug, of -1 als het blad beveiligd is.
Private Function ZetBladVast(ByVal ws As Worksheet, Optional ByVal bereik As Range) As Long
    Dim doel As Range, formules As Range, area As Range, cel As Range

    If ws.ProtectContents Then ZetBladVast = -1: Exit Function
    If bereik Is Nothing Then Set bereik = ws.UsedRange

    Set doel = Application.Intersect(bereik, ws.UsedRange)
    If doel Is Nothing Then Exit Function

    On Error Resume Next
    Set formules = doel.SpecialCells(xlCellTypeFormulas)
    On Error GoTo 0
    If formules Is Nothing Then Exit Function

    ZetBladVast = CLng(formules.CountLarge)

    For Each area In formules.Areas
        If Not ProbeerSchrijf(area) Then
            ' Meestal een (deel van een) matrixformule: cel voor cel afhandelen
            For Each cel In area.Cells
                If cel.HasFormula Then
                    If cel.HasArray Then
                        ProbeerSchrijf cel.CurrentArray
                    Else
                        ProbeerSchrijf cel
                    End If
                End If
            Next cel
        End If
    Next area
End Function

Private Function ProbeerSchrijf(ByVal doel As Range) As Boolean
    On Error GoTo Fout
    SchrijfWaarden doel
    ProbeerSchrijf = True
    Exit Function
Fout:
    ProbeerSchrijf = False
End Function

' Schrijft de huidige waarden terug als constanten (in 1 keer, snel)
Private Sub SchrijfWaarden(ByVal doel As Range)
    Dim v As Variant, i As Long, j As Long
    v = doel.Value2
    If IsArray(v) Then
        For i = LBound(v, 1) To UBound(v, 1)
            For j = LBound(v, 2) To UBound(v, 2)
                v(i, j) = TekstVeilig(v(i, j))
            Next j
        Next i
    Else
        v = TekstVeilig(v)
    End If
    doel.Value2 = v
End Sub

' Tekst die op een getal/datum/formule lijkt (bijv. "001") zou anders omgezet
' worden naar een getal. Met een apostrof blijft het tekst.
Private Function TekstVeilig(ByVal x As Variant) As Variant
    If VarType(x) = vbString Then
        If Len(x) > 0 Then
            If Left$(x, 1) = "=" Or Left$(x, 1) = "'" Or IsNumeric(x) Or IsDate(x) _
               Or UCase$(x) = "TRUE" Or UCase$(x) = "FALSE" Then
                TekstVeilig = "'" & x
                Exit Function
            End If
        End If
    End If
    TekstVeilig = x
End Function

'---------------------------------------------------------------------
' Hulpfuncties
'---------------------------------------------------------------------
Private Function GeselecteerdBereik() As Range
    If ActiveWorkbook Is Nothing Then Exit Function
    If TypeName(Selection) <> "Range" Then
        MsgBox "Selecteer eerst een of meer cellen.", vbInformation, TITEL
        Exit Function
    End If
    Set GeselecteerdBereik = Selection
End Function

Private Function ActiefBlad() As Worksheet
    If ActiveWorkbook Is Nothing Then Exit Function
    If TypeName(ActiveSheet) <> "Worksheet" Then
        MsgBox "Het actieve tabblad is geen werkblad.", vbInformation, TITEL
        Exit Function
    End If
    Set ActiefBlad = ActiveSheet
End Function

Private Function Bevestig(ByVal waar As String) As Boolean
    If Not VRAAG_BEVESTIGING Then Bevestig = True: Exit Function
    Bevestig = (MsgBox("Formules in " & waar & " omzetten naar waarden?" & vbNewLine & vbNewLine & _
                       "Dit kan niet ongedaan gemaakt worden met Ongedaan maken.", _
                       vbYesNo + vbQuestion + vbDefaultButton2, TITEL) = vbYes)
End Function

Private Sub StartBatch()
    mOudeBerekening = Application.Calculation
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
End Sub

Private Sub EindBatch()
    Application.Calculation = mOudeBerekening
    Application.ScreenUpdating = True
End Sub

Private Sub Klaar(ByVal n As Long, ByVal geblokkeerd As Long)
    Dim tekst As String
    tekst = n & " formulecellen omgezet naar waarden."
    If geblokkeerd > 0 Then tekst = tekst & vbNewLine & geblokkeerd & " beveiligd(e) tabblad(en) overgeslagen."
    MsgBox tekst, vbInformation, TITEL
End Sub
