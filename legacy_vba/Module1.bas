Attribute VB_Name = "Module1"
'===========================================================
' SUIVI DES NON-CONFORMITES - Norvel Composites
' cree par JLB 2009
' modif SAR 2013 : ajout pareto
' modif JLB 2016 : export csv pour reunion qualite
' !! ne pas renommer les onglets !!
'===========================================================

Dim chem

Sub LANCER_TOUT()
    Application.ScreenUpdating = False
    ImportDonnees
    CalcHebdo
    CalcPareto
    CalcAlertes
    CalcMensuel
    Colorer
    ExportSorties
    Application.ScreenUpdating = True
    MsgBox "Traitement termine"
End Sub

Sub ImportDonnees()
    chem = ThisWorkbook.Path & Application.PathSeparator & "data" & Application.PathSeparator
    ImportCSV chem & "defect_log.csv", "SAISIE"
    ImportCSV chem & "production_log.csv", "PROD"
    ImportCSV chem & "defect_types.csv", "REF_DEF"
    ImportCSV chem & "parts.csv", "REF_PCE"
    'ImportCSV chem & "parameters.csv", "PARAM"   ' pas utilise pour l'instant
End Sub

Sub ImportCSV(fic, nomF)
    Dim ws, f, lgn, t, i, j
    Set ws = FeuilleVide(nomF)
    f = FreeFile
    Open fic For Input As #f
    i = 1
    Do While Not EOF(f)
        Line Input #f, lgn
        lgn = Replace(Replace(lgn, vbCr, ""), vbLf, "")
        If Len(lgn) > 0 Then
            t = Split(lgn, ",")
            For j = 0 To UBound(t)
                ws.Cells(i, j + 1).Value = t(j)
            Next j
            i = i + 1
        End If
    Loop
    Close #f
End Sub

Function FeuilleVide(nom)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(nom)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = nom
    Else
        ws.Cells.Clear
    End If
    ws.Cells.NumberFormat = "@"
    Set FeuilleVide = ws
End Function

Sub Colorer()
    Dim i, v
    i = 2
    Do While Sheets("RAP_HEBDO").Cells(i, 1).Value <> ""
        v = Val(Sheets("RAP_HEBDO").Cells(i, 7).Value)
        If v > 3 Then
            Sheets("RAP_HEBDO").Cells(i, 8).Interior.Color = RGB(255, 0, 0)
        ElseIf v > 2 Then
            Sheets("RAP_HEBDO").Cells(i, 8).Interior.Color = RGB(255, 165, 0)
        Else
            Sheets("RAP_HEBDO").Cells(i, 8).Interior.Color = RGB(0, 176, 80)
        End If
        i = i + 1
    Loop
End Sub

Sub ExportSorties()
    Dim dos, nm, f, r, c, nc, lgn
    dos = ThisWorkbook.Path & Application.PathSeparator & "sorties_legacy"
    On Error Resume Next
    MkDir dos
    On Error GoTo 0
    For Each nm In Array("RAP_HEBDO", "PARETO", "ALERTES", "MENS_LIGNES", "MENS_GLOBAL")
        nc = 0
        Do While Sheets(nm).Cells(1, nc + 1).Value <> ""
            nc = nc + 1
        Loop
        f = FreeFile
        Open dos & Application.PathSeparator & LCase(nm) & ".csv" For Output As #f
        r = 1
        Do While Sheets(nm).Cells(r, 1).Value <> ""
            lgn = ""
            For c = 1 To nc
                If c > 1 Then lgn = lgn & ","
                lgn = lgn & Sheets(nm).Cells(r, c).Value
            Next c
            Print #f, lgn
            r = r + 1
        Loop
        Close #f
    Next nm
End Sub
