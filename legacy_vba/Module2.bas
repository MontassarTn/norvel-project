Attribute VB_Name = "Module2"
'===========================================================
' calculs indicateurs qualite
' v3.2 - 14/03/2016
' taux en %, cout en euros
'===========================================================

Function DT(s)
    DT = DateSerial(Val(Left(s, 4)), Val(Mid(s, 6, 2)), Val(Mid(s, 9, 2)))
End Function

Function NumL(l)
    NumL = Val(Mid(l, 2, 1))
End Function

Function Txt(x)
    Txt = Replace(CStr(x), ",", ".")
End Function

Function Fmt2(x)
    Fmt2 = Replace(Format(x, "0.00"), ",", ".")
End Function

Function ProdJour(d, lg)
    Dim r
    ProdJour = 0
    r = 2
    Do While Sheets("PROD").Cells(r, 1).Value <> ""
        If DT(Sheets("PROD").Cells(r, 1).Value) = d And Sheets("PROD").Cells(r, 2).Value = lg Then
            ProdJour = Val(Sheets("PROD").Cells(r, 3).Value)
            Exit Function
        End If
        r = r + 1
    Loop
End Function

Function PrixPce(ref)
    Dim r
    PrixPce = 0
    r = 2
    Do While Sheets("REF_PCE").Cells(r, 1).Value <> ""
        If Sheets("REF_PCE").Cells(r, 1).Value = ref Then PrixPce = Val(Sheets("REF_PCE").Cells(r, 3).Value)
        r = r + 1
    Loop
End Function

Function CoutNC(r)
    Dim ws, q, c, dp, x
    Set ws = Sheets("SAISIE")
    q = Val(ws.Cells(r, 6).Value)
    c = ws.Cells(r, 5).Value
    dp = ws.Cells(r, 7).Value
    If dp = "SCRAP" Or (dp = "REWORK" And c = "D06") Then
        x = q * PrixPce(ws.Cells(r, 4).Value)
    ElseIf dp = "REWORK" Then
        x = q * Val(Sheets("REF_DEF").Cells(Val(Mid(c, 2)) + 1, 4).Value) * 45
    Else
        x = 0
    End If
    CoutNC = Round(x, 0)
End Function

Sub CalcHebdo()
    Dim p(1 To 53, 1 To 4), nd(1 To 53, 1 To 4), rb(1 To 53, 1 To 4), cq(1 To 53, 1 To 4)
    Dim wsP, wsS, wr, r, d, l, s, q, i, tx, txr, st
    Set wsP = Sheets("PROD")
    Set wsS = Sheets("SAISIE")
    r = 2
    Do While wsP.Cells(r, 1).Value <> ""
        d = DT(wsP.Cells(r, 1).Value)
        s = DatePart("ww", d, vbMonday, vbFirstFourDays)
        l = NumL(wsP.Cells(r, 2).Value)
        p(s, l) = p(s, l) + Val(wsP.Cells(r, 3).Value)
        r = r + 1
    Loop
    r = 2
    Do While wsS.Cells(r, 1).Value <> ""
        d = DT(wsS.Cells(r, 2).Value)
        If ProdJour(d, wsS.Cells(r, 3).Value) = 0 Then GoTo suivant
        s = DatePart("ww", d, vbMonday, vbFirstFourDays)
        l = NumL(wsS.Cells(r, 3).Value)
        q = Val(wsS.Cells(r, 6).Value)
        ' modif JLB 2011 - ne pas toucher
        If l = 4 Then
            nd(s, l) = nd(s, l) + q / 2
        Else
            nd(s, l) = nd(s, l) + q
        End If
        If wsS.Cells(r, 7).Value = "SCRAP" Then rb(s, l) = rb(s, l) + q
        cq(s, l) = cq(s, l) + CoutNC(r)
suivant:
        r = r + 1
    Loop
    Set wr = FeuilleVide("RAP_HEBDO")
    wr.Range("A1:I1").Value = Array("semaine", "ligne", "produit", "nb_def", "tx_def", "rebut", "tx_rebut", "statut", "cnq_eur")
    i = 2
    For s = 1 To 53
        For l = 1 To 4
            If p(s, l) > 0 Then
                tx = (nd(s, l) + 0) / p(s, l) * 100
                txr = (rb(s, l) + 0) / p(s, l) * 100
                If txr >= 3 Then
                    st = "ROUGE"
                ElseIf txr > 2 Then
                    st = "ORANGE"
                Else
                    st = "VERT"
                End If
                wr.Cells(i, 1).Value = "2026-W" & Format(s, "00")
                wr.Cells(i, 2).Value = "L" & l
                wr.Cells(i, 3).Value = Txt(p(s, l))
                wr.Cells(i, 4).Value = Txt(nd(s, l) + 0)
                wr.Cells(i, 5).Value = Fmt2(tx)
                wr.Cells(i, 6).Value = Txt(rb(s, l) + 0)
                wr.Cells(i, 7).Value = Fmt2(txr)
                wr.Cells(i, 8).Value = st
                wr.Cells(i, 9).Value = Txt(cq(s, l) + 0)
                i = i + 1
            End If
        Next l
    Next s
End Sub

Sub CalcPareto()
    Dim q(1 To 60, 1 To 8), cod(1 To 8), qt(1 To 8)
    Dim ws, wr, r, d, k, c, i, j, n, tot, cum, t
    Set ws = Sheets("SAISIE")
    r = 2
    Do While ws.Cells(r, 1).Value <> ""
        d = DT(ws.Cells(r, 2).Value)
        d = d - Weekday(d, vbSunday) + 1
        k = (d - DateSerial(2025, 12, 28)) / 7 + 1
        c = Val(Mid(ws.Cells(r, 5).Value, 2))
        q(k, c) = q(k, c) + Val(ws.Cells(r, 6).Value)
        r = r + 1
    Loop
    Set wr = FeuilleVide("PARETO")
    wr.Range("A1:H1").Value = Array("semaine_du", "rang", "code", "libelle", "qte", "pct", "cumul_pct", "prioritaire")
    i = 2
    For k = 1 To 60
        tot = 0
        For c = 1 To 8
            tot = tot + q(k, c)
        Next c
        If tot > 0 Then
            n = 0
            For c = 1 To 8
                If q(k, c) > 0 Then
                    n = n + 1
                    cod(n) = c
                    qt(n) = q(k, c)
                End If
            Next c
            For j = 1 To n - 1
                For c = 1 To n - j
                    If qt(c) < qt(c + 1) Then
                        t = qt(c): qt(c) = qt(c + 1): qt(c + 1) = t
                        t = cod(c): cod(c) = cod(c + 1): cod(c + 1) = t
                    End If
                Next c
            Next j
            cum = 0
            For j = 1 To n
                wr.Cells(i, 1).Value = Format(DateSerial(2025, 12, 28) + (k - 1) * 7, "yyyy-mm-dd")
                wr.Cells(i, 2).Value = CStr(j)
                wr.Cells(i, 3).Value = "D" & Format(cod(j), "00")
                wr.Cells(i, 4).Value = Sheets("REF_DEF").Cells(cod(j) + 1, 2).Value
                wr.Cells(i, 5).Value = Txt(qt(j))
                wr.Cells(i, 6).Value = Fmt2(qt(j) / tot * 100)
                If cum < 80 Then wr.Cells(i, 8).Value = "X"
                cum = cum + qt(j) / tot * 100
                wr.Cells(i, 7).Value = Fmt2(cum)
                i = i + 1
            Next j
        End If
    Next k
End Sub

Sub CalcAlertes()
    Dim ws, wa, r, r2, i, sv, q, li, cnt, d1, d2
    Set ws = Sheets("SAISIE")
    Set wa = FeuilleVide("ALERTES")
    wa.Range("A1:H1").Value = Array("type", "id", "date", "ligne", "piece", "code", "qte", "message")
    i = 2
    ' defauts critiques
    r = 2
    Do While ws.Cells(r, 1).Value <> ""
        q = Val(ws.Cells(r, 6).Value)
        sv = Sheets("REF_DEF").Cells(Val(Mid(ws.Cells(r, 5).Value, 2)) + 1, 3).Value
        If q >= 20 Then
            If sv = "MINOR" Then
                sv = "MAJOR"
            ElseIf sv = "MAJOR" Then
                sv = "CRITICAL"
            End If
        End If
        If sv = "CRITICAL" Then
            wa.Cells(i, 1).Value = "CRITIQUE"
            wa.Cells(i, 2).Value = ws.Cells(r, 1).Value
            wa.Cells(i, 3).Value = ws.Cells(r, 2).Value
            wa.Cells(i, 4).Value = ws.Cells(r, 3).Value
            wa.Cells(i, 5).Value = ws.Cells(r, 4).Value
            wa.Cells(i, 6).Value = ws.Cells(r, 5).Value
            wa.Cells(i, 7).Value = ws.Cells(r, 6).Value
            wa.Cells(i, 8).Value = "defaut critique - prevenir resp. qualite"
            i = i + 1
        End If
        r = r + 1
    Loop
    ' recurrences (ajout SAR 2013)
    For li = 1 To 4
        r = 2
        Do While ws.Cells(r, 1).Value <> ""
            If ws.Cells(r, 3).Value = "L" & li Then
                cnt = 0
                d1 = DT(ws.Cells(r, 2).Value)
                r2 = r
                Do While ws.Cells(r2, 1).Value <> ""
                    If ws.Cells(r2, 3).Value = "L" & li And ws.Cells(r2, 4).Value = ws.Cells(r, 4).Value And ws.Cells(r2, 5).Value = ws.Cells(r, 5).Value Then
                        d2 = DT(ws.Cells(r2, 2).Value)
                        If d2 - d1 <= 6 Then cnt = cnt + 1
                    End If
                    r2 = r2 + 1
                Loop
                If cnt >= 3 Then
                    wa.Cells(i, 1).Value = "RECURRENCE"
                    wa.Cells(i, 2).Value = ws.Cells(r, 1).Value
                    wa.Cells(i, 3).Value = ws.Cells(r, 2).Value
                    wa.Cells(i, 4).Value = ws.Cells(r, 3).Value
                    wa.Cells(i, 5).Value = ws.Cells(r, 4).Value
                    wa.Cells(i, 6).Value = ws.Cells(r, 5).Value
                    wa.Cells(i, 7).Value = CStr(cnt)
                    wa.Cells(i, 8).Value = "recurrence 7j - ouvrir 8D"
                    i = i + 1
                End If
            End If
            r = r + 1
        Loop
    Next li
End Sub

Sub CalcMensuel()
    Dim p(1 To 12, 1 To 4), nd(1 To 12, 1 To 4), rb(1 To 12, 1 To 4), cq(1 To 12, 1 To 4)
    Dim tq(1 To 12, 1 To 8), nbc(1 To 12), nbr(1 To 12), cod(1 To 8), qt(1 To 8)
    Dim wsP, wsS, wsA, wm, wg, r, d, m, l, c, q, dp, i, tx, fr, st, tp, tn, pt, dt2, ct, j, n, t, top3
    Set wsP = Sheets("PROD")
    Set wsS = Sheets("SAISIE")
    Set wsA = Sheets("ALERTES")
    r = 2
    Do While wsP.Cells(r, 1).Value <> ""
        m = Month(DT(wsP.Cells(r, 1).Value))
        l = NumL(wsP.Cells(r, 2).Value)
        p(m, l) = p(m, l) + Val(wsP.Cells(r, 3).Value)
        r = r + 1
    Loop
    r = 2
    Do While wsS.Cells(r, 1).Value <> ""
        d = DT(wsS.Cells(r, 2).Value)
        m = Month(d)
        l = NumL(wsS.Cells(r, 3).Value)
        c = Val(Mid(wsS.Cells(r, 5).Value, 2))
        q = Val(wsS.Cells(r, 6).Value)
        dp = wsS.Cells(r, 7).Value
        tq(m, c) = tq(m, c) + q
        If ProdJour(d, wsS.Cells(r, 3).Value) <> 0 Then
            If dp <> "ACCEPT" Then
                If l = 4 Then
                    nd(m, l) = nd(m, l) + q / 2
                Else
                    nd(m, l) = nd(m, l) + q
                End If
            End If
            If dp = "SCRAP" Then rb(m, l) = rb(m, l) + q
            cq(m, l) = cq(m, l) + CoutNC(r)
        End If
        r = r + 1
    Loop
    r = 2
    Do While wsA.Cells(r, 1).Value <> ""
        m = Month(DT(wsA.Cells(r, 3).Value))
        If wsA.Cells(r, 1).Value = "CRITIQUE" Then
            nbc(m) = nbc(m) + 1
        Else
            nbr(m) = nbr(m) + 1
        End If
        r = r + 1
    Loop
    Set wm = FeuilleVide("MENS_LIGNES")
    wm.Range("A1:J1").Value = Array("mois", "ligne", "produit", "defauts", "tx_def", "rebut", "tx_rebut", "statut", "cnq_eur", "tendance")
    i = 2
    For m = 1 To 12
        For l = 1 To 4
            If p(m, l) > 0 Then
                tx = (nd(m, l) + 0) / p(m, l) * 100
                fr = (rb(m, l) + 0) / p(m, l)
                If fr > 0.03 Then
                    st = "ROUGE"
                ElseIf fr * 100 > 2 Then
                    st = "ORANGE"
                Else
                    st = "VERT"
                End If
                tn = "N/A"
                If m > 1 Then
                    If p(m - 1, l) > 0 Then
                        tp = (nd(m - 1, l) + 0) / p(m - 1, l) * 100
                        If tx - tp > 0.5 Then
                            tn = "HAUSSE"
                        ElseIf tx - tp < -0.5 Then
                            tn = "BAISSE"
                        Else
                            tn = "STABLE"
                        End If
                    End If
                End If
                wm.Cells(i, 1).Value = "2026-" & Format(m, "00")
                wm.Cells(i, 2).Value = "L" & l
                wm.Cells(i, 3).Value = Txt(p(m, l))
                wm.Cells(i, 4).Value = Txt(nd(m, l) + 0)
                wm.Cells(i, 5).Value = Fmt2(tx)
                wm.Cells(i, 6).Value = Txt(rb(m, l) + 0)
                wm.Cells(i, 7).Value = Fmt2(fr * 100)
                wm.Cells(i, 8).Value = st
                wm.Cells(i, 9).Value = Txt(cq(m, l) + 0)
                wm.Cells(i, 10).Value = tn
                i = i + 1
            End If
        Next l
    Next m
    Set wg = FeuilleVide("MENS_GLOBAL")
    wg.Range("A1:G1").Value = Array("mois", "produit", "defauts", "cnq_eur", "top3", "nb_critiques", "nb_recurrences")
    i = 2
    For m = 1 To 12
        pt = 0: dt2 = 0: ct = 0
        For l = 1 To 4
            pt = pt + p(m, l)
            dt2 = dt2 + nd(m, l)
            ct = ct + cq(m, l)
        Next l
        If pt > 0 Then
            n = 0
            For c = 1 To 8
                If tq(m, c) > 0 Then
                    n = n + 1
                    cod(n) = c
                    qt(n) = tq(m, c)
                End If
            Next c
            For j = 1 To n - 1
                For c = 1 To n - j
                    If qt(c) < qt(c + 1) Then
                        t = qt(c): qt(c) = qt(c + 1): qt(c + 1) = t
                        t = cod(c): cod(c) = cod(c + 1): cod(c + 1) = t
                    End If
                Next c
            Next j
            top3 = ""
            For j = 1 To 3
                If j <= n Then
                    If j > 1 Then top3 = top3 & " / "
                    top3 = top3 & "D" & Format(cod(j), "00")
                End If
            Next j
            wg.Cells(i, 1).Value = "2026-" & Format(m, "00")
            wg.Cells(i, 2).Value = Txt(pt)
            wg.Cells(i, 3).Value = Txt(dt2 + 0)
            wg.Cells(i, 4).Value = Txt(ct + 0)
            wg.Cells(i, 5).Value = top3
            wg.Cells(i, 6).Value = Txt(nbc(m) + 0)
            wg.Cells(i, 7).Value = Txt(nbr(m) + 0)
            i = i + 1
        End If
    Next m
End Sub
