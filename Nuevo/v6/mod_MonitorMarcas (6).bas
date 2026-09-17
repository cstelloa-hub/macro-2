Attribute VB_Name = "mod_MonitorMarcas"
Option Explicit
'==================================================================
' mod_MonitorMarcas  -  Monitor de marcas de alternativos
'
' Fuente : hoja "3" del libro de marcas (Basicos ya calculados)
' Eje    : columna Fecha (Dia de la marca).  Salida en pb.
' Hojas  : Config | Mapa | Eventos | Resumen | Mensual
'
' Ventanas: SEMANA | MTD | PERIODO (inicio libre en C11) | YTD
'           las cuatro cortan en el domingo de la semana de C6
'
' Uso    : 1) CrearConfig  (una sola vez)
'          2) llenar Config!C3 con la ruta del libro de marcas
'          3) ActualizarTodo
'          4) REVISAR LA HOJA MAPA antes de creerle a los numeros
'==================================================================

Private Const S_CFG As String = "Config"
Private Const S_MAP As String = "Mapa"
Private Const S_EVT As String = "Eventos"
Private Const S_RES As String = "Resumen"
Private Const S_MEN As String = "Mensual"
Private Const S_MAR As String = "Marcas"
Private Const ROJO  As Long = 789716          ' RGB(212,12,12)
Private Const MAXC  As Long = 400

' --- mapa de columnas
Private AFP(1 To 4) As String
Private cNom As Long, cFec As Long, cEEF As Long
Private cVSB As Long, cVAD As Long, cSBS As Long
Private cEst As Long, cExp As Long, cVin As Long, cMon As Long
Private colGap(1 To 3, 1 To 4) As Long
Private colBas(1 To 3, 1 To 4) As Long

' --- datos en memoria
Private nEv As Long
Private pbF As Double
Private evFec() As Double, evNom() As String, evSBS() As String
Private evEEF() As Variant, evVSB() As Double, evVAD() As Double
Private evGap() As Double, evBas() As Double, evFlg() As String
Private evEst() As String, evExp() As String, evVin() As String, evMon() As String

' --- cabeceras de la hoja fuente, cargadas una sola vez
Private hN() As String, gN() As String, g2N() As String
Private nCols As Long

' --- corte de la semana reportada, para nombrar el archivo
Private gDom As Double


'==================================================================
' PRINCIPAL
'==================================================================
Public Sub ActualizarTodo()
    Dim t0 As Single: t0 = Timer
    Dim wbS As Workbook, wsS As Worksheet

    On Error GoTo Fallo
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    AFP(1) = "PRF": AFP(2) = "INF": AFP(3) = "RIF": AFP(4) = "HAF"
    ValidarConfig
    pbF = CDbl(Cfg("C10").Value)

    Estado "abriendo libro de marcas..."
    Set wbS = AbrirFuente()
    On Error Resume Next
    Set wsS = wbS.Worksheets(CStr(Cfg("C4").Value))
    On Error GoTo Fallo
    If wsS Is Nothing Then Err.Raise 5, , "No existe la hoja '" & Cfg("C4").Value & "' en el libro fuente."

    Estado "detectando columnas..."
    DetectarMapa wsS

    Estado "cargando marcas..."
    CargarDatos wsS

    wbS.Close SaveChanges:=False
    Set wbS = Nothing

    If nEv = 0 Then Err.Raise 5, , "No hay marcas con fecha >= " & _
                                   Format(Cfg("C5").Value, "dd/mm/yyyy") & "."

    Estado "eventos...":  ConstruirEventos
    Estado "marcas...":   ConstruirMarcas
    Estado "resumen...":  ConstruirResumen
    Estado "mensual...":  ConstruirMensual
    Estado "archivando...": ArchivarResumen

    ThisWorkbook.Worksheets(S_RES).Activate
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.StatusBar = False

    MsgBox nEv & " marcas procesadas en " & Format(Timer - t0, "0.0") & " s." & vbLf & vbLf & _
           "Revisa la hoja Mapa antes de usar los numeros.", vbInformation, "Monitor de marcas"
    Exit Sub

Fallo:
    On Error Resume Next
    If Not wbS Is Nothing Then wbS.Close SaveChanges:=False
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.StatusBar = False
    MsgBox "Error: " & Err.Description, vbCritical, "mod_MonitorMarcas"
End Sub


'==================================================================
' CONFIG
'==================================================================
Private Function Cfg(addr As String) As Range
    Set Cfg = ThisWorkbook.Worksheets(S_CFG).Range(addr)
End Function

Public Sub CrearConfig()
    Dim ws As Worksheet, ruta As String
    Set ws = Hoja(S_CFG)

    On Error Resume Next
    ruta = CStr(ws.Range("C3").Value)       ' conserva la ruta si ya estaba
    On Error GoTo 0

    ws.Cells.Clear
    ws.Range("B2").Value = "MONITOR DE MARCAS - configuracion"
    ws.Range("B3").Value = "Ruta del libro de marcas (vacio = preguntar)"
    ws.Range("C3").Value = ruta
    ws.Range("B4").Value = "Hoja fuente":                ws.Range("C4").Value = "3"
    ws.Range("B5").Value = "Fecha inicio del historico": ws.Range("C5").Value = DateSerial(2025, 1, 1)
    ws.Range("B6").Value = "FECHA DE REFERENCIA (vacio = ultima marca)"
    ws.Range("B7").Value = "Top N del ranking":          ws.Range("C7").Value = 5
    ws.Range("B8").Value = "Fila de cabecera de campos": ws.Range("C8").Value = 3
    ws.Range("B9").Value = "Primera fila de datos":      ws.Range("C9").Value = 4
    ws.Range("B10").Value = "Factor a pb (10000 si viene en fraccion)": ws.Range("C10").Value = 10000
    ws.Range("B11").Value = "INICIO DEL PERIODO LIBRE (acumulado desde esa fecha)"
    ws.Range("C11").Value = DateSerial(2026, 5, 1)
    ws.Range("B12").Value = "CARPETA DE ARCHIVO (vacio = subcarpeta Resumenes junto a este libro)"
    ws.Range("B13").Value = "Formato a archivar: XLSX | PDF | AMBOS | NO"
    ws.Range("C13").Value = "XLSX"

    ws.Range("B16").Value = "Como usar C6"
    ws.Range("B17").Value = "   una semana  ->  cualquier dia de esa semana"
    ws.Range("B18").Value = "   un mes      ->  el ultimo dia del mes (ej. 31/08/2026)"
    ws.Range("B19").Value = "   al dia      ->  dejarla vacia"
    ws.Range("B20").Value = "SEMANA  = lunes a domingo de la semana de C6"
    ws.Range("B21").Value = "MTD     = del 1 del mes de C6 hasta el domingo de esa semana"
    ws.Range("B22").Value = "PERIODO = desde C11 hasta el mismo corte"
    ws.Range("B23").Value = "YTD     = del 1 de enero hasta el mismo corte"
    ws.Range("B24").Value = "El archivado nunca sobrescribe: si ya existe, agrega _v2, _v3.."

    ws.Range("B2").Font.Bold = True: ws.Range("B2").Font.Size = 11
    ws.Range("B3:B13").Font.Bold = True
    ws.Range("B6:C6").Interior.Color = RGB(255, 242, 204)
    ws.Range("B11:C11").Interior.Color = RGB(255, 242, 204)
    ws.Range("B12:C13").Interior.Color = RGB(226, 239, 218)
    ws.Range("B16").Font.Bold = True
    ws.Range("B17:B24").Font.Italic = True
    ws.Columns("B").ColumnWidth = 50
    ws.Columns("C").ColumnWidth = 24
    ws.Range("C5,C6,C11").NumberFormat = "dd/mm/yyyy"
    ws.Cells.Font.Name = "Arial": ws.Cells.Font.Size = 8
End Sub

Private Sub ValidarConfig()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(S_CFG)
    On Error GoTo 0
    If ws Is Nothing Then CrearConfig

    If Len(Trim$(CStr(Cfg("C4").Value))) = 0 Then Cfg("C4").Value = "3"
    If Not IsDate(Cfg("C5").Value) Then Cfg("C5").Value = DateSerial(2025, 1, 1)
    If Val(Cfg("C7").Value) <= 0 Then Cfg("C7").Value = 5
    If Val(Cfg("C8").Value) <= 0 Then Cfg("C8").Value = 3
    If Val(Cfg("C9").Value) <= Val(Cfg("C8").Value) Then Cfg("C9").Value = Val(Cfg("C8").Value) + 1
    If Val(Cfg("C10").Value) = 0 Then Cfg("C10").Value = 10000
    If Not IsDate(Cfg("C11").Value) Then Cfg("C11").Value = DateSerial(2026, 5, 1)
    If Len(Trim$(CStr(Cfg("C13").Value))) = 0 Then Cfg("C13").Value = "XLSX"
End Sub

Private Function AbrirFuente() As Workbook
    Dim p As String
    p = Trim$(CStr(Cfg("C3").Value))
    If Len(p) = 0 Then
        p = Application.GetOpenFilename("Excel,*.xls*", , "Selecciona el libro de marcas")
        If p = "False" Then Err.Raise 5, , "Cancelado por el usuario."
        Cfg("C3").Value = p
    End If
    If Len(Dir(p)) = 0 Then Err.Raise 5, , "No encuentro el archivo:" & vbLf & p
    Set AbrirFuente = Workbooks.Open(Filename:=p, ReadOnly:=True, UpdateLinks:=0)
End Function


'==================================================================
' DETECCION DE COLUMNAS
'==================================================================
Private Function Norm(v As Variant) As String
    Dim s As String, i As Long, ch As String, o As String
    s = UCase$(Trim$(CStr(v)))
    s = Replace(s, ChrW(193), "A"): s = Replace(s, ChrW(201), "E"): s = Replace(s, ChrW(205), "I")
    s = Replace(s, ChrW(211), "O"): s = Replace(s, ChrW(218), "U"): s = Replace(s, ChrW(209), "N")
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If (ch >= "A" And ch <= "Z") Or (ch >= "0" And ch <= "9") Then o = o & ch
    Next i
    Norm = o
End Function

Private Function Letra(col As Long) As String
    If col <= 0 Then Exit Function
    Letra = Split(ThisWorkbook.Worksheets(1).Cells(1, col).Address(True, False), "$")(0)
End Function

Private Function EsAfp(s As String, nm As String) As Boolean
    ' tolera el sufijo de fondo: PRF, PRF2, PRF3, "RIF 3" -> RIF3
    If Len(s) < 3 Or Len(s) > 4 Then Exit Function
    EsAfp = (Left$(s, 3) = nm)
End Function

Private Sub CargarCabeceras(ws As Worksheet, hr As Long, gr As Long)
    Dim v As Variant, j As Long
    nCols = MAXC
    ReDim hN(1 To nCols): ReDim gN(1 To nCols): ReDim g2N(1 To nCols)

    v = ws.Range(ws.Cells(hr, 1), ws.Cells(hr, nCols)).Value
    For j = 1 To nCols: hN(j) = Norm(v(1, j)): Next j

    If gr >= 1 Then
        v = ws.Range(ws.Cells(gr, 1), ws.Cells(gr, nCols)).Value
        For j = 1 To nCols: gN(j) = Norm(v(1, j)): Next j
    End If
    If gr - 1 >= 1 Then
        v = ws.Range(ws.Cells(gr - 1, 1), ws.Cells(gr - 1, nCols)).Value
        For j = 1 To nCols: g2N(j) = Norm(v(1, j)): Next j
    End If
End Sub

Private Sub DetectarMapa(ws As Worksheet)
    Dim hr As Long, gr As Long, j As Long, k As Long, f As Long, jj As Long
    Dim s As String, g As String, tipo As String

    hr = CLng(Cfg("C8").Value)
    gr = hr - 1
    CargarCabeceras ws, hr, gr

    cNom = 0: cFec = 0: cEEF = 0: cVSB = 0: cVAD = 0: cSBS = 0
    cEst = 0: cExp = 0: cVin = 0: cMon = 0
    For f = 1 To 3
        For k = 1 To 4
            colGap(f, k) = 0: colBas(f, k) = 0
    Next k, f

    ' --- campos base y clasificacion
    For j = 1 To nCols
        s = hN(j)
        If Len(s) > 0 Then
            If s = "NOMBRE" And cNom = 0 Then cNom = j
            If s = "FECHA" And cFec = 0 Then cFec = j
            If Left$(s, 7) = "FECHAEE" And cEEF = 0 Then cEEF = j
            If s = "VARSBS" And cVSB = 0 Then cVSB = j
            If (s = "VARADJ" Or s = "VARADJUSTADA" Or s = "VARADJUST") And cVAD = 0 Then cVAD = j
            If Left$(s, 9) = "CODIGOSBS" And cSBS = 0 Then cSBS = j
            If s = "ESTRATEGIA" And cEst = 0 Then cEst = j
            If Left$(s, 9) = "EXPOSICIO" And cExp = 0 Then cExp = j
            If s = "VINTAGE" And cVin = 0 Then cVin = j
            If s = "MONEDA" And cMon = 0 Then cMon = j
        End If
    Next j

    ' --- bloques: cada PRF abre un bloque de 4 (PRF INF RIF HAF)
    For j = 1 To nCols
        If EsAfp(hN(j), "PRF") Then
            g = EtiquetaGrupo(j)
            tipo = ""
            If InStr(g, "GAP") > 0 Then tipo = "GAP"
            If InStr(g, "BASIC") > 0 Then tipo = "BAS"
            f = FondoDe(g)
            If f = 0 Then f = FondoDe(hN(j))
            If f = 0 Then f = 1
            If tipo <> "" Then
                For k = 1 To 4
                    jj = BuscarAfp(j, AFP(k))
                    If tipo = "GAP" Then
                        If colGap(f, k) = 0 Then colGap(f, k) = jj
                    Else
                        If colBas(f, k) = 0 Then colBas(f, k) = jj
                    End If
                Next k
            End If
        End If
    Next j

    EscribirMapa
    VerificarMapa
End Sub

Private Function EtiquetaGrupo(j As Long) As String
    ' busca hacia la izquierda la etiqueta de grupo (GAPs / Basicos)
    Dim k As Long, s As String
    For k = j To 1 Step -1
        s = gN(k)
        If InStr(s, "GAP") > 0 Or InStr(s, "BASIC") > 0 Then EtiquetaGrupo = s: Exit Function
        s = g2N(k)
        If InStr(s, "GAP") > 0 Or InStr(s, "BASIC") > 0 Then EtiquetaGrupo = s: Exit Function
    Next k
End Function

Private Function FondoDe(g As String) As Long
    If InStr(g, "3") > 0 Then
        FondoDe = 3
    ElseIf InStr(g, "2") > 0 Then
        FondoDe = 2
    ElseIf InStr(g, "1") > 0 Then
        FondoDe = 1
    End If
End Function

Private Function BuscarAfp(j0 As Long, nm As String) As Long
    Dim k As Long, lim As Long
    lim = j0 + 9
    If lim > nCols Then lim = nCols
    For k = j0 To lim
        If EsAfp(hN(k), nm) Then BuscarAfp = k: Exit Function
    Next k
End Function

Private Sub EscribirMapa()
    Dim m As Worksheet, r As Long, f As Long, k As Long
    Set m = Hoja(S_MAP)
    m.Cells.Clear
    m.Range("A1:C1").Value = Array("Campo", "Columna", "Letra")
    r = 2
    PonMapa m, r, "Nombre", cNom
    PonMapa m, r, "Fecha (Dia)", cFec
    PonMapa m, r, "Fecha EEFF", cEEF
    PonMapa m, r, "Var SBS", cVSB
    PonMapa m, r, "Var Adj", cVAD
    PonMapa m, r, "Codigo SBS", cSBS
    PonMapa m, r, "Estrategia", cEst
    PonMapa m, r, "Exposicion", cExp
    PonMapa m, r, "Vintage", cVin
    PonMapa m, r, "Moneda", cMon
    For f = 1 To 3
        For k = 1 To 4
            PonMapa m, r, "GAP  F" & f & "  " & AFP(k), colGap(f, k)
        Next k
        For k = 1 To 4
            PonMapa m, r, "BAS  F" & f & "  " & AFP(k), colBas(f, k)
        Next k
    Next f
    FormatoTabla m, 1, 1, r - 1, 3
    m.Columns("A:C").AutoFit
End Sub

Private Sub PonMapa(m As Worksheet, r As Long, etq As String, col As Long)
    m.Cells(r, 1).Value = etq
    If col = 0 Then
        m.Cells(r, 2).Value = "NO ENCONTRADA"
        m.Range(m.Cells(r, 1), m.Cells(r, 3)).Interior.Color = RGB(251, 227, 227)
    Else
        m.Cells(r, 2).Value = col
        m.Cells(r, 3).Value = Letra(col)
    End If
    r = r + 1
End Sub

Private Sub VerificarMapa()
    Dim falta As String, f As Long, k As Long
    If cNom = 0 Then falta = falta & "Nombre, "
    If cFec = 0 Then falta = falta & "Fecha, "
    If cVAD = 0 Then falta = falta & "Var Adj, "
    If cSBS = 0 Then falta = falta & "Codigo SBS, "
    For f = 1 To 3
        For k = 1 To 4
            If colGap(f, k) = 0 Then falta = falta & "GAP F" & f & " " & AFP(k) & ", "
            If colBas(f, k) = 0 Then falta = falta & "BAS F" & f & " " & AFP(k) & ", "
    Next k, f
    If Len(falta) > 0 Then
        Err.Raise 5, , "No pude ubicar estas columnas:" & vbLf & vbLf & _
                       Left$(falta, Len(falta) - 2) & vbLf & vbLf & _
                       "Revisa la hoja Mapa."
    End If
End Sub


'==================================================================
' CARGA DE DATOS
'==================================================================
Private Sub CargarDatos(ws As Worksheet)
    Dim r1 As Long, r2 As Long, ultC As Long
    Dim i As Long, n As Long, f As Long, k As Long
    Dim fIni As Double, v As Variant, dat As Variant, nMax As Long

    r1 = CLng(Cfg("C9").Value)
    r2 = ws.Cells(ws.Rows.Count, cNom).End(xlUp).Row
    If r2 < r1 Then Err.Raise 5, , "La hoja fuente no tiene filas de datos."
    fIni = CDbl(CDate(Cfg("C5").Value))

    ultC = cNom
    If cFec > ultC Then ultC = cFec
    If cEEF > ultC Then ultC = cEEF
    If cVSB > ultC Then ultC = cVSB
    If cVAD > ultC Then ultC = cVAD
    If cSBS > ultC Then ultC = cSBS
    If cEst > ultC Then ultC = cEst
    If cExp > ultC Then ultC = cExp
    If cVin > ultC Then ultC = cVin
    If cMon > ultC Then ultC = cMon
    For f = 1 To 3
        For k = 1 To 4
            If colGap(f, k) > ultC Then ultC = colGap(f, k)
            If colBas(f, k) > ultC Then ultC = colBas(f, k)
    Next k, f

    dat = ws.Range(ws.Cells(r1, 1), ws.Cells(r2, ultC)).Value

    nMax = r2 - r1 + 1
    ReDim evFec(1 To nMax): ReDim evNom(1 To nMax): ReDim evSBS(1 To nMax)
    ReDim evEEF(1 To nMax): ReDim evVSB(1 To nMax): ReDim evVAD(1 To nMax)
    ReDim evFlg(1 To nMax)
    ReDim evEst(1 To nMax): ReDim evExp(1 To nMax)
    ReDim evVin(1 To nMax): ReDim evMon(1 To nMax)
    ReDim evGap(1 To nMax, 1 To 3, 1 To 4)
    ReDim evBas(1 To nMax, 1 To 3, 1 To 4)

    n = 0
    For i = 1 To nMax
        v = dat(i, cFec)
        If IsDate(v) Then
            If CDbl(CDate(v)) >= fIni Then
                n = n + 1
                evFec(n) = CDbl(CDate(v))
                evNom(n) = Trim$(CStr(dat(i, cNom)))
                evSBS(n) = Trim$(CStr(dat(i, cSBS)))
                If cEEF > 0 Then evEEF(n) = dat(i, cEEF)
                If cVSB > 0 Then evVSB(n) = Num(dat(i, cVSB))
                evVAD(n) = Num(dat(i, cVAD))
                evEst(n) = Cat(dat, i, cEst)
                evExp(n) = Cat(dat, i, cExp)
                evVin(n) = TramoVintage(Cat(dat, i, cVin))
                evMon(n) = Cat(dat, i, cMon)
                For f = 1 To 3
                    For k = 1 To 4
                        evGap(n, f, k) = Num(dat(i, colGap(f, k)))
                        evBas(n, f, k) = Num(dat(i, colBas(f, k)))
                Next k, f
            End If
        ElseIf Len(Trim$(CStr(dat(i, cNom)))) > 0 _
               And Left$(Norm(dat(i, cNom)), 5) <> "TOTAL" Then
            ' fila con nombre pero sin fecha: entra solo para que la veas
            n = n + 1
            evFec(n) = 0
            evNom(n) = Trim$(CStr(dat(i, cNom)))
            evSBS(n) = Trim$(CStr(dat(i, cSBS)))
            evVAD(n) = Num(dat(i, cVAD))
            evEst(n) = "(sin dato)": evExp(n) = "(sin dato)"
            evVin(n) = "(sin dato)": evMon(n) = "(sin dato)"
            evFlg(n) = "SIN FECHA"
        End If
    Next i

    nEv = n
    MarcarDuplicados
End Sub

Private Function Num(v As Variant) As Double
    If IsNumeric(v) Then Num = CDbl(v)
End Function

Private Function Cat(dat As Variant, i As Long, col As Long) As String
    If col = 0 Then Cat = "(sin dato)": Exit Function
    Cat = Trim$(CStr(dat(i, col)))
    If Len(Cat) = 0 Then Cat = "(sin dato)"
End Function

Private Function TramoVintage(v As String) As String
    Dim a As Long
    If v = "(sin dato)" Or Len(Trim$(v)) = 0 Then TramoVintage = "(sin dato)": Exit Function
    If Not IsNumeric(v) Then
        ' textos tipo "Ant. - 2008"
        If InStr(UCase$(v), "ANT") > 0 Then
            TramoVintage = "Ant. 2009"
        Else
            TramoVintage = "(sin dato)"
        End If
        Exit Function
    End If
    a = CLng(Val(v))
    If a <= 0 Then
        TramoVintage = "(sin dato)"
    ElseIf a < 2009 Then
        TramoVintage = "Ant. 2009"
    ElseIf a <= 2012 Then
        TramoVintage = "2009 - 2012"
    ElseIf a <= 2016 Then
        TramoVintage = "2013 - 2016"
    ElseIf a <= 2020 Then
        TramoVintage = "2017 - 2020"
    Else
        TramoVintage = "2021 en adelante"
    End If
End Function

Private Sub MarcarDuplicados()
    ' misma llave SBS + mismo Var Adj dentro de 45 dias -> marca, no anula
    Dim i As Long, j As Long
    For i = 2 To nEv
        If evFec(i) > 0 And evVAD(i) <> 0 Then
            For j = i - 1 To 1 Step -1
                If evFec(j) > 0 Then
                    If evFec(i) - evFec(j) > 45 Then Exit For
                    If evSBS(j) = evSBS(i) Then
                        If Abs(evVAD(j) - evVAD(i)) < 0.0000000001 Then
                            evFlg(i) = Trim$(evFlg(i) & " REVISAR dup")
                            Exit For
                        End If
                    End If
                End If
            Next j
        End If
    Next i
End Sub


'==================================================================
' HOJA EVENTOS
'==================================================================
Private Function TotalPRF(i As Long) As Double
    TotalPRF = (evBas(i, 1, 1) + evBas(i, 2, 1) + evBas(i, 3, 1)) * pbF
End Function

Private Function LunesDe(d As Double) As Double
    If d > 0 Then LunesDe = d - Weekday(CDate(d), vbMonday) + 1
End Function

Private Sub ConstruirEventos()
    Dim ws As Worksheet, i As Long, f As Long, k As Long, c As Long
    Dim hdr() As Variant, arr() As Variant
    Dim ultF As Long

    Set ws = Hoja(S_EVT)
    ws.Cells.Clear
    If ws.AutoFilterMode Then ws.AutoFilterMode = False

    ' ---- cabecera, de un golpe
    ReDim hdr(1 To 38)
    hdr(1) = "Fecha": hdr(2) = "Semana": hdr(3) = "Mes": hdr(4) = "Fondo alternativo"
    hdr(5) = "Codigo SBS": hdr(6) = "Fecha EEFF": hdr(7) = "Var SBS": hdr(8) = "Var Adj"
    c = 9
    For f = 1 To 3
        For k = 1 To 4
            hdr(c) = "GAP F" & f & " " & AFP(k): c = c + 1
        Next k
        For k = 1 To 4
            hdr(c) = "pb F" & f & " " & AFP(k): c = c + 1
        Next k
    Next f
    hdr(33) = "Total pb PRF"
    hdr(34) = "Estrategia": hdr(35) = "Exposicion"
    hdr(36) = "Vintage": hdr(37) = "Moneda": hdr(38) = "Flag"
    ws.Range(ws.Cells(1, 1), ws.Cells(1, 38)).Value = hdr

    ' ---- cuerpo en memoria
    ReDim arr(1 To nEv, 1 To 38)
    For i = 1 To nEv
        If evFec(i) > 0 Then
            arr(i, 1) = CDate(evFec(i))
            arr(i, 2) = CDate(LunesDe(evFec(i)))
            arr(i, 3) = Format(CDate(evFec(i)), "yyyy-mm")
        End If
        arr(i, 4) = evNom(i)
        arr(i, 5) = evSBS(i)
        arr(i, 6) = evEEF(i)
        arr(i, 7) = evVSB(i)
        arr(i, 8) = evVAD(i)
        c = 9
        For f = 1 To 3
            For k = 1 To 4
                arr(i, c) = evGap(i, f, k): c = c + 1
            Next k
            For k = 1 To 4
                arr(i, c) = evBas(i, f, k) * pbF: c = c + 1
            Next k
        Next f
        arr(i, 33) = TotalPRF(i)
        arr(i, 34) = evEst(i)
        arr(i, 35) = evExp(i)
        arr(i, 36) = evVin(i)
        arr(i, 37) = evMon(i)
        arr(i, 38) = evFlg(i)
    Next i

    ' ---- un solo volcado
    ultF = 1 + nEv
    ws.Range(ws.Cells(2, 1), ws.Cells(ultF, 38)).Value = arr

    ' ---- formatos por bloque
    ws.Range(ws.Cells(2, 1), ws.Cells(ultF, 2)).NumberFormat = "dd/mm/yyyy"
    ws.Range(ws.Cells(2, 7), ws.Cells(ultF, 8)).NumberFormat = "0.00%"
    For f = 1 To 3
        c = 9 + (f - 1) * 8
        ws.Range(ws.Cells(2, c), ws.Cells(ultF, c + 3)).NumberFormat = "0.00%"
        ws.Range(ws.Cells(2, c + 4), ws.Cells(ultF, c + 7)).NumberFormat = "#,##0.00"
    Next f
    ws.Range(ws.Cells(2, 33), ws.Cells(ultF, 33)).NumberFormat = "#,##0.00"

    FormatoTabla ws, 1, 1, ultF, 38
    ws.Range(ws.Cells(1, 1), ws.Cells(1, 38)).AutoFilter

    ' anchos fijos: el AutoFit sobre 38 columnas y miles de filas es lo que colgaba
    ws.Columns("A:C").ColumnWidth = 10
    ws.Columns("D").ColumnWidth = 32
    ws.Columns("E").ColumnWidth = 14
    ws.Columns("F:H").ColumnWidth = 10
    ws.Range(ws.Columns(9), ws.Columns(33)).ColumnWidth = 9
    ws.Range(ws.Columns(34), ws.Columns(37)).ColumnWidth = 16
    ws.Columns(38).ColumnWidth = 14
End Sub


'==================================================================
' HOJA RESUMEN
'   1. Cuadro por fondo, partido en LOCALES y EXTRANJEROS
'   2. Por estrategia, dentro de cada mitad
'   3. Por moneda      4. Por vintage
'   5. Top winners y top losers
'   Los conteos de marcas van en tabla aparte, al costado.
'==================================================================
Private Function EsLocal(i As Long) As Boolean
    EsLocal = (Norm(evExp(i)) = "LOCAL")
End Function

Private Function EnAmbito(i As Long, ambito As Long) As Boolean
    ' 0 = todos | 1 = solo local | 2 = solo extranjero
    Select Case ambito
        Case 1: EnAmbito = EsLocal(i)
        Case 2: EnAmbito = Not EsLocal(i)
        Case Else: EnAmbito = True
    End Select
End Function

Private Function TieneImpacto(i As Long, f As Long) As Boolean
    TieneImpacto = (evBas(i, f, 1) <> 0)
End Function

Private Function MaxFecha() As Double
    Dim i As Long, m As Double
    For i = 1 To nEv
        If evFec(i) > m Then m = evFec(i)
    Next i
    MaxFecha = m
End Function

Private Function CatDe(i As Long, tipo As Long) As String
    Select Case tipo
        Case 1: CatDe = evEst(i)
        Case 2: CatDe = evExp(i)
        Case 3: CatDe = evVin(i)
        Case 4: CatDe = evMon(i)
    End Select
End Function

Private Sub ConstruirResumen()
    Dim ws As Worksheet
    Dim dRef As Double, lun As Double, dom As Double
    Dim mIni As Double, aIni As Double, pIni As Double
    Dim r As Long, topN As Long

    Set ws = Hoja(S_RES)
    ws.Cells.Clear

    If IsDate(Cfg("C6").Value) Then dRef = CDbl(CDate(Cfg("C6").Value)) Else dRef = MaxFecha()
    If dRef = 0 Then Err.Raise 5, , "No hay fechas validas para armar el Resumen."
    lun = LunesDe(dRef): dom = lun + 6
    gDom = dom
    mIni = CDbl(DateSerial(Year(CDate(dRef)), Month(CDate(dRef)), 1))
    aIni = CDbl(DateSerial(Year(CDate(dRef)), 1, 1))
    If IsDate(Cfg("C11").Value) Then pIni = CDbl(CDate(Cfg("C11").Value)) Else pIni = aIni

    ws.Range("B2").Value = "MONITOR DE MARCAS DE ALTERNATIVOS"
    ws.Range("B2").Font.Bold = True: ws.Range("B2").Font.Size = 11
    ws.Range("B3").Value = "Semana del " & Format(CDate(lun), "dd/mm/yyyy") & _
                           " al " & Format(CDate(dom), "dd/mm/yyyy") & _
                           "   |   MTD desde " & Format(CDate(mIni), "dd/mm/yyyy") & _
                           "   |   Periodo desde " & Format(CDate(pIni), "dd/mm/yyyy") & _
                           "   |   YTD " & Year(CDate(dRef)) & _
                           "   |   contribuciones en pb   |   eje: Dia de la marca"

    r = CuadroFondos(ws, 5, "1.  RESUMEN POR FONDO - LOCALES", 1, _
                     lun, dom, mIni, pIni, aIni)
    r = CuadroFondos(ws, r + 3, "     RESUMEN POR FONDO - EXTRANJEROS", 2, _
                     lun, dom, mIni, pIni, aIni)

    r = Desglose(ws, r + 3, "2.  LOCALES - POR ESTRATEGIA", 1, 1, _
                 lun, dom, mIni, pIni, aIni)
    r = Desglose(ws, r + 3, "     EXTRANJEROS - POR ESTRATEGIA", 1, 2, _
                 lun, dom, mIni, pIni, aIni)

    r = Desglose(ws, r + 3, "3.  POR MONEDA", 4, 0, lun, dom, mIni, pIni, aIni)
    r = Desglose(ws, r + 3, "4.  POR VINTAGE", 3, 0, lun, dom, mIni, pIni, aIni)

    topN = CLng(Cfg("C7").Value)
    Ranking ws, r + 3, lun, dom, True
    Ranking ws, r + 3 + topN + 4, lun, dom, False

    ws.Columns("B:X").AutoFit
    ws.Cells.Font.Name = "Arial": ws.Cells.Font.Size = 8
End Sub

'------------------------------------------------- 1. cuadro por fondo
Private Function CuadroFondos(ws As Worksheet, r0 As Long, titulo As String, ambito As Long, _
                              lun As Double, dom As Double, mIni As Double, _
                              pIni As Double, aIni As Double) As Long
    Dim f As Long, k As Long, i As Long, r As Long, d As Double
    Dim nS As Long, nM As Long, nP As Long, nY As Long
    Dim sS(1 To 4) As Double, sM(1 To 4) As Double
    Dim sP(1 To 4) As Double, sY(1 To 4) As Double

    ws.Cells(r0, 2).Value = titulo
    ws.Cells(r0, 2).Font.Bold = True

    ' --- cabeceras de grupo del cuadro de contribucion
    ws.Cells(r0 + 1, 3).Value = "SEMANA"
    ws.Cells(r0 + 1, 7).Value = "MTD"
    ws.Cells(r0 + 1, 11).Value = "PERIODO"
    ws.Cells(r0 + 1, 15).Value = "YTD"
    ws.Range(ws.Cells(r0 + 1, 3), ws.Cells(r0 + 1, 15)).Font.Bold = True
    ws.Range(ws.Cells(r0 + 1, 3), ws.Cells(r0 + 1, 15)).HorizontalAlignment = xlCenter

    ws.Range(ws.Cells(r0 + 2, 2), ws.Cells(r0 + 2, 18)).Value = Array("Fondo", _
        "PRF", "vs INF", "vs RIF", "vs HAF", _
        "PRF", "vs INF", "vs RIF", "vs HAF", _
        "PRF", "vs INF", "vs RIF", "vs HAF", _
        "PRF", "vs INF", "vs RIF", "vs HAF")

    ' --- tabla de conteos, al costado
    ws.Range(ws.Cells(r0 + 2, 20), ws.Cells(r0 + 2, 24)).Value = _
        Array("Fondo", "Semana", "MTD", "Periodo", "YTD")

    r = r0 + 3
    For f = 1 To 3
        nS = 0: nM = 0: nP = 0: nY = 0
        For k = 1 To 4: sS(k) = 0: sM(k) = 0: sP(k) = 0: sY(k) = 0: Next k
        For i = 1 To nEv
            d = evFec(i)
            If d > 0 And d <= dom And EnAmbito(i, ambito) Then
                If d >= aIni Then
                    If TieneImpacto(i, f) Then nY = nY + 1
                    For k = 1 To 4: sY(k) = sY(k) + evBas(i, f, k) * pbF: Next k
                End If
                If d >= pIni Then
                    If TieneImpacto(i, f) Then nP = nP + 1
                    For k = 1 To 4: sP(k) = sP(k) + evBas(i, f, k) * pbF: Next k
                End If
                If d >= mIni Then
                    If TieneImpacto(i, f) Then nM = nM + 1
                    For k = 1 To 4: sM(k) = sM(k) + evBas(i, f, k) * pbF: Next k
                End If
                If d >= lun Then
                    If TieneImpacto(i, f) Then nS = nS + 1
                    For k = 1 To 4: sS(k) = sS(k) + evBas(i, f, k) * pbF: Next k
                End If
            End If
        Next i

        ws.Cells(r, 2).Value = "Fondo " & f
        For k = 1 To 4: ws.Cells(r, 2 + k).Value = sS(k): Next k
        For k = 1 To 4: ws.Cells(r, 6 + k).Value = sM(k): Next k
        For k = 1 To 4: ws.Cells(r, 10 + k).Value = sP(k): Next k
        For k = 1 To 4: ws.Cells(r, 14 + k).Value = sY(k): Next k

        ws.Cells(r, 20).Value = "Fondo " & f
        ws.Cells(r, 21).Value = nS
        ws.Cells(r, 22).Value = nM
        ws.Cells(r, 23).Value = nP
        ws.Cells(r, 24).Value = nY
        r = r + 1
    Next f

    ws.Range(ws.Cells(r0 + 3, 3), ws.Cells(r - 1, 18)).NumberFormat = "#,##0.00"
    FormatoTabla ws, r0 + 2, 2, r - 1, 18
    FormatoTabla ws, r0 + 2, 20, r - 1, 24
    CuadroFondos = r - 1
End Function

'------------------------------------------------- 2/3/4. desgloses
Private Function Desglose(ws As Worksheet, r0 As Long, titulo As String, tipo As Long, _
                          ambito As Long, lun As Double, dom As Double, mIni As Double, _
                          pIni As Double, aIni As Double) As Long
    ' contribucion PRF (F1+F2+F3) agrupada por categoria, en pb
    Dim dc As Object, i As Long, r As Long, ix As Long, d As Double, fLim As Double
    Dim ky As String, arr As Variant
    Dim nS() As Long, vS() As Double, vM() As Double, vP() As Double, vY() As Double
    Dim tS As Double, tM As Double, tP As Double, tY As Double, tN As Long

    fLim = aIni
    If pIni < fLim Then fLim = pIni

    Set dc = CreateObject("Scripting.Dictionary")
    For i = 1 To nEv
        If evFec(i) > 0 And evFec(i) <= dom And evFec(i) >= fLim And EnAmbito(i, ambito) Then
            ky = CatDe(i, tipo)
            If Not dc.Exists(ky) Then dc.Add ky, dc.Count + 1
        End If
    Next i

    ws.Cells(r0, 2).Value = titulo & "   (contribucion a Profuturo, pb)"
    ws.Cells(r0, 2).Font.Bold = True
    ws.Range(ws.Cells(r0 + 1, 2), ws.Cells(r0 + 1, 6)).Value = _
        Array("Categoria", "Semana", "MTD", "Periodo", "YTD")
    ws.Range(ws.Cells(r0 + 1, 8), ws.Cells(r0 + 1, 9)).Value = _
        Array("Categoria", "Marcas sem")

    If dc.Count = 0 Then
        ws.Cells(r0 + 2, 2).Value = "Sin datos"
        FormatoTabla ws, r0 + 1, 2, r0 + 2, 6
        FormatoTabla ws, r0 + 1, 8, r0 + 1, 9
        Desglose = r0 + 2
        Exit Function
    End If

    ReDim nS(1 To dc.Count): ReDim vS(1 To dc.Count): ReDim vM(1 To dc.Count)
    ReDim vP(1 To dc.Count): ReDim vY(1 To dc.Count)

    For i = 1 To nEv
        d = evFec(i)
        If d > 0 And d <= dom And d >= fLim And EnAmbito(i, ambito) Then
            ix = dc(CatDe(i, tipo))
            If d >= aIni Then vY(ix) = vY(ix) + TotalPRF(i)
            If d >= pIni Then vP(ix) = vP(ix) + TotalPRF(i)
            If d >= mIni Then vM(ix) = vM(ix) + TotalPRF(i)
            If d >= lun Then
                vS(ix) = vS(ix) + TotalPRF(i)
                nS(ix) = nS(ix) + 1
            End If
        End If
    Next i

    arr = dc.Keys
    r = r0 + 2
    For i = 1 To dc.Count
        ws.Cells(r, 2).Value = arr(i - 1)
        ws.Cells(r, 3).Value = vS(i)
        ws.Cells(r, 4).Value = vM(i)
        ws.Cells(r, 5).Value = vP(i)
        ws.Cells(r, 6).Value = vY(i)
        ws.Cells(r, 8).Value = arr(i - 1)
        ws.Cells(r, 9).Value = nS(i)
        tN = tN + nS(i): tS = tS + vS(i)
        tM = tM + vM(i): tP = tP + vP(i): tY = tY + vY(i)
        r = r + 1
    Next i

    ws.Cells(r, 2).Value = "Total"
    ws.Cells(r, 3).Value = tS
    ws.Cells(r, 4).Value = tM
    ws.Cells(r, 5).Value = tP
    ws.Cells(r, 6).Value = tY
    ws.Cells(r, 8).Value = "Total"
    ws.Cells(r, 9).Value = tN
    ws.Range(ws.Cells(r, 2), ws.Cells(r, 6)).Font.Bold = True
    ws.Range(ws.Cells(r, 8), ws.Cells(r, 9)).Font.Bold = True

    ws.Range(ws.Cells(r0 + 2, 3), ws.Cells(r, 6)).NumberFormat = "#,##0.00"
    FormatoTabla ws, r0 + 1, 2, r, 6
    FormatoTabla ws, r0 + 1, 8, r, 9
    Desglose = r
End Function

'------------------------------------------------- 5. rankings
Private Sub Ranking(ws As Worksheet, r0 As Long, lun As Double, dom As Double, winners As Boolean)
    Dim idx() As Long, vv() As Double
    Dim n As Long, i As Long, j As Long, r As Long, topN As Long
    Dim ti As Long, tv As Double

    ReDim idx(1 To nEv): ReDim vv(1 To nEv)
    For i = 1 To nEv
        If evFec(i) >= lun And evFec(i) <= dom Then
            n = n + 1: idx(n) = i: vv(n) = TotalPRF(i)
        End If
    Next i

    ws.Cells(r0, 2).Value = IIf(winners, "5.  TOP WINNERS de la semana", "     TOP LOSERS de la semana")
    ws.Cells(r0, 2).Font.Bold = True
    ws.Range(ws.Cells(r0 + 1, 2), ws.Cells(r0 + 1, 11)).Value = _
        Array("#", "Fondo alternativo", "Dia", "Var Adj", "pb F1", "pb F2", "pb F3", _
              "Total pb", "Exposicion", "Estrategia")

    If n = 0 Then
        ws.Cells(r0 + 2, 2).Value = "Sin marcas en la semana"
        FormatoTabla ws, r0 + 1, 2, r0 + 2, 11
        Exit Sub
    End If

    For i = 1 To n - 1
        For j = 1 To n - i
            If (winners And vv(j) < vv(j + 1)) Or ((Not winners) And vv(j) > vv(j + 1)) Then
                tv = vv(j): vv(j) = vv(j + 1): vv(j + 1) = tv
                ti = idx(j): idx(j) = idx(j + 1): idx(j + 1) = ti
            End If
        Next j
    Next i

    topN = CLng(Cfg("C7").Value)
    If topN > n Then topN = n

    r = r0 + 2
    For i = 1 To topN
        ws.Cells(r, 2).Value = i
        ws.Cells(r, 3).Value = evNom(idx(i))
        ws.Cells(r, 4).Value = CDate(evFec(idx(i)))
        ws.Cells(r, 5).Value = evVAD(idx(i))
        ws.Cells(r, 6).Value = evBas(idx(i), 1, 1) * pbF
        ws.Cells(r, 7).Value = evBas(idx(i), 2, 1) * pbF
        ws.Cells(r, 8).Value = evBas(idx(i), 3, 1) * pbF
        ws.Cells(r, 9).Value = vv(i)
        ws.Cells(r, 10).Value = evExp(idx(i))
        ws.Cells(r, 11).Value = evEst(idx(i))
        r = r + 1
    Next i

    ws.Range(ws.Cells(r0 + 2, 4), ws.Cells(r - 1, 4)).NumberFormat = "dd/mm/yyyy"
    ws.Range(ws.Cells(r0 + 2, 5), ws.Cells(r - 1, 5)).NumberFormat = "0.00%"
    ws.Range(ws.Cells(r0 + 2, 6), ws.Cells(r - 1, 9)).NumberFormat = "#,##0.00"
    FormatoTabla ws, r0 + 1, 2, r - 1, 11
End Sub


'==================================================================
' HOJA MARCAS  -  detalle resumido de las marcas de la semana
'==================================================================
Private Sub ConstruirMarcas()
    Dim ws As Worksheet, i As Long, r As Long, n As Long
    Dim dRef As Double, lun As Double, dom As Double
    Dim t1 As Double, t2 As Double, t3 As Double, tt As Double

    Set ws = Hoja(S_MAR)
    ws.Cells.Clear

    If IsDate(Cfg("C6").Value) Then dRef = CDbl(CDate(Cfg("C6").Value)) Else dRef = MaxFecha()
    If dRef = 0 Then Exit Sub
    lun = LunesDe(dRef): dom = lun + 6

    ws.Range("B2").Value = "MARCAS DE LA SEMANA   (contribucion en pb)"
    ws.Range("B2").Font.Bold = True: ws.Range("B2").Font.Size = 11
    ws.Range("B3").Value = "Semana del " & Format(CDate(lun), "dd/mm/yyyy") & _
                           " al " & Format(CDate(dom), "dd/mm/yyyy")

    ws.Range(ws.Cells(5, 2), ws.Cells(5, 10)).Value = _
        Array("Dia", "Fondo alternativo", "Var Adj", "pb F1", "pb F2", "pb F3", _
              "Total", "Exposicion", "Estrategia")

    r = 6
    For i = 1 To nEv
        If evFec(i) >= lun And evFec(i) <= dom Then
            n = n + 1
            ws.Cells(r, 2).Value = CDate(evFec(i))
            ws.Cells(r, 3).Value = evNom(i)
            ws.Cells(r, 4).Value = evVAD(i)
            ws.Cells(r, 5).Value = evBas(i, 1, 1) * pbF
            ws.Cells(r, 6).Value = evBas(i, 2, 1) * pbF
            ws.Cells(r, 7).Value = evBas(i, 3, 1) * pbF
            ws.Cells(r, 8).Value = TotalPRF(i)
            ws.Cells(r, 9).Value = evExp(i)
            ws.Cells(r, 10).Value = evEst(i)
            t1 = t1 + evBas(i, 1, 1) * pbF
            t2 = t2 + evBas(i, 2, 1) * pbF
            t3 = t3 + evBas(i, 3, 1) * pbF
            tt = tt + TotalPRF(i)
            r = r + 1
        End If
    Next i

    If n = 0 Then
        ws.Cells(r, 2).Value = "Sin marcas en la semana"
        FormatoTabla ws, 5, 2, r, 10
        ws.Cells.Font.Name = "Arial": ws.Cells.Font.Size = 8
        Exit Sub
    End If

    ws.Cells(r, 3).Value = n & " marcas"
    ws.Cells(r, 5).Value = t1
    ws.Cells(r, 6).Value = t2
    ws.Cells(r, 7).Value = t3
    ws.Cells(r, 8).Value = tt
    ws.Range(ws.Cells(r, 2), ws.Cells(r, 10)).Font.Bold = True

    ws.Range(ws.Cells(6, 2), ws.Cells(r - 1, 2)).NumberFormat = "dd/mm/yyyy"
    ws.Range(ws.Cells(6, 4), ws.Cells(r, 4)).NumberFormat = "0.00%"
    ws.Range(ws.Cells(6, 5), ws.Cells(r, 8)).NumberFormat = "#,##0.00"
    FormatoTabla ws, 5, 2, r, 10
    ws.Columns("B:J").AutoFit
    ws.Cells.Font.Name = "Arial": ws.Cells.Font.Size = 8
End Sub


'==================================================================
' HOJA MENSUAL
'==================================================================
Private Sub ConstruirMensual()
    Dim ws As Worksheet, meses As Object, filas As Object
    Dim i As Long, f As Long, j As Long, nM As Long
    Dim ky As String, ms As String
    Dim arrM As Variant, arrF As Variant, dt() As Double

    Set ws = Hoja(S_MEN)
    ws.Cells.Clear
    Set meses = CreateObject("Scripting.Dictionary")
    Set filas = CreateObject("Scripting.Dictionary")

    For i = 1 To nEv
        If evFec(i) > 0 Then
            ms = Format(CDate(evFec(i)), "yyyy-mm")
            If Not meses.Exists(ms) Then meses.Add ms, meses.Count + 1
            For f = 1 To 3
                If evBas(i, f, 1) <> 0 Then
                    ky = "F" & f & "|" & evNom(i)
                    If Not filas.Exists(ky) Then filas.Add ky, filas.Count + 1
                End If
            Next f
        End If
    Next i
    If meses.Count = 0 Or filas.Count = 0 Then Exit Sub

    nM = meses.Count
    arrM = meses.Keys
    arrF = filas.Keys
    ReDim dt(1 To filas.Count, 1 To nM)

    For i = 1 To nEv
        If evFec(i) > 0 Then
            ms = Format(CDate(evFec(i)), "yyyy-mm")
            For f = 1 To 3
                If evBas(i, f, 1) <> 0 Then
                    ky = "F" & f & "|" & evNom(i)
                    dt(filas(ky), meses(ms)) = dt(filas(ky), meses(ms)) + evBas(i, f, 1) * pbF
                End If
            Next f
        End If
    Next i

    ws.Range("A1:B1").Value = Array("Fondo", "Fondo alternativo")
    For j = 1 To nM
        ws.Cells(1, 2 + j).Value = arrM(j - 1)
    Next j
    For i = 1 To filas.Count
        ws.Cells(1 + i, 1).Value = Split(arrF(i - 1), "|")(0)
        ws.Cells(1 + i, 2).Value = Split(arrF(i - 1), "|")(1)
        For j = 1 To nM
            ws.Cells(1 + i, 2 + j).Value = dt(i, j)
        Next j
    Next i

    ws.Range(ws.Cells(2, 3), ws.Cells(1 + filas.Count, 2 + nM)).NumberFormat = "#,##0.00"
    FormatoTabla ws, 1, 1, 1 + filas.Count, 2 + nM
    ws.Columns("A").ColumnWidth = 8
    ws.Columns("B").ColumnWidth = 32
    ws.Range(ws.Columns(3), ws.Columns(2 + nM)).ColumnWidth = 10
End Sub


'==================================================================
' ARCHIVADO
'   Guarda una copia de las hojas Resumen y Marcas en una carpeta.
'   Nombre: Monitor_Marcas_AAAAMMDD  (domingo de la semana reportada)
'   Nunca sobrescribe: si existe, agrega _v2, _v3...
'==================================================================
Public Sub ArchivarAhora()
    ' para archivar sin volver a leer el libro de marcas
    On Error GoTo Fallo
    ValidarConfig
    If gDom = 0 Then Err.Raise 5, , "Corre ActualizarTodo primero: no hay un corte cargado."
    ArchivarResumen
    MsgBox "Archivado.", vbInformation, "Monitor de marcas"
    Exit Sub
Fallo:
    MsgBox "Error al archivar: " & Err.Description, vbCritical, "Monitor de marcas"
End Sub

Private Sub ArchivarResumen()
    Dim modo As String, carp As String, base As String, ruta As String
    Dim wbN As Workbook, v As Long, sufijo As String

    modo = UCase$(Trim$(CStr(Cfg("C13").Value)))
    If modo = "NO" Or Len(modo) = 0 Then Exit Sub
    If gDom = 0 Then Exit Sub

    carp = Trim$(CStr(Cfg("C12").Value))
    If Len(carp) = 0 Then carp = ThisWorkbook.Path & Application.PathSeparator & "Resumenes"
    If Right$(carp, 1) = Application.PathSeparator Then carp = Left$(carp, Len(carp) - 1)
    If Not CrearCarpeta(carp) Then
        MsgBox "No pude crear o acceder a la carpeta de archivo:" & vbLf & carp & vbLf & vbLf & _
               "El resumen quedo armado igual.", vbExclamation, "Monitor de marcas"
        Exit Sub
    End If

    base = "Monitor_Marcas_" & Format(CDate(gDom), "yyyymmdd")

    ' --- copia de las dos hojas a un libro nuevo
    On Error GoTo Fallo
    ThisWorkbook.Worksheets(Array(S_RES, S_MAR)).Copy
    Set wbN = ActiveWorkbook

    ' --- nombre libre
    v = 1: sufijo = ""
    Do
        ruta = carp & Application.PathSeparator & base & sufijo
        If Len(Dir(ruta & ".xlsx")) = 0 And Len(Dir(ruta & ".pdf")) = 0 Then Exit Do
        v = v + 1: sufijo = "_v" & v
        If v > 50 Then Err.Raise 5, , "Demasiadas versiones del mismo dia en la carpeta."
    Loop

    If modo = "XLSX" Or modo = "AMBOS" Then
        wbN.SaveAs Filename:=ruta & ".xlsx", FileFormat:=xlOpenXMLWorkbook
    End If
    If modo = "PDF" Or modo = "AMBOS" Then
        wbN.ExportAsFixedFormat Type:=xlTypePDF, Filename:=ruta & ".pdf", _
                                Quality:=xlQualityStandard, OpenAfterPublish:=False
    End If

    wbN.Close SaveChanges:=False
    Set wbN = Nothing
    Exit Sub

Fallo:
    On Error Resume Next
    If Not wbN Is Nothing Then wbN.Close SaveChanges:=False
    On Error GoTo 0
    MsgBox "No pude archivar la copia: " & Err.Description & vbLf & vbLf & _
           "El resumen quedo armado igual.", vbExclamation, "Monitor de marcas"
End Sub

Private Function CrearCarpeta(ruta As String) As Boolean
    ' crea la ruta nivel por nivel; devuelve True si al final existe
    Dim partes As Variant, acum As String, i As Long, sep As String
    sep = Application.PathSeparator

    On Error Resume Next
    If Len(Dir(ruta, vbDirectory)) > 0 Then CrearCarpeta = True: Exit Function
    On Error GoTo 0

    partes = Split(ruta, sep)
    If UBound(partes) < 0 Then Exit Function

    If Left$(ruta, 2) = sep & sep Then
        ' ruta de red \\servidor\recurso
        If UBound(partes) < 4 Then Exit Function
        acum = sep & sep & partes(2) & sep & partes(3)
        For i = 4 To UBound(partes)
            If Len(partes(i)) > 0 Then
                acum = acum & sep & partes(i)
                If Len(Dir(acum, vbDirectory)) = 0 Then
                    On Error Resume Next
                    MkDir acum
                    On Error GoTo 0
                End If
            End If
        Next i
    Else
        acum = partes(0)
        For i = 1 To UBound(partes)
            If Len(partes(i)) > 0 Then
                acum = acum & sep & partes(i)
                If Len(Dir(acum, vbDirectory)) = 0 Then
                    On Error Resume Next
                    MkDir acum
                    On Error GoTo 0
                End If
            End If
        Next i
    End If

    On Error Resume Next
    CrearCarpeta = (Len(Dir(ruta, vbDirectory)) > 0)
    On Error GoTo 0
End Function


'==================================================================
' UTILIDADES
'==================================================================
Private Function Hoja(nm As String) As Worksheet
    Dim w As Worksheet
    On Error Resume Next
    Set w = ThisWorkbook.Worksheets(nm)
    On Error GoTo 0
    If w Is Nothing Then
        Set w = ThisWorkbook.Worksheets.Add( _
                After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        w.Name = nm
    End If
    Set Hoja = w
End Function

Private Sub FormatoTabla(ws As Worksheet, r1 As Long, c1 As Long, r2 As Long, c2 As Long)
    With ws.Range(ws.Cells(r1, c1), ws.Cells(r2, c2))
        .Font.Name = "Arial": .Font.Size = 8
        .Borders(xlInsideHorizontal).LineStyle = xlNone
        .Borders(xlInsideVertical).LineStyle = xlNone
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
    End With
    With ws.Range(ws.Cells(r1, c1), ws.Cells(r1, c2))
        .Interior.Color = ROJO
        .Font.Color = vbWhite
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
    End With
End Sub

Private Sub Estado(s As String)
    Application.StatusBar = "Monitor de marcas: " & s
End Sub
