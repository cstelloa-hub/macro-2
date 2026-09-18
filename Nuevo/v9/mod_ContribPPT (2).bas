Attribute VB_Name = "mod_ContribPPT"
Option Explicit
'==============================================================================
' mod_ContribPPT — Contribucion de Profuturo, formato PPT
'
' Un solo modulo, el mismo en los dos libros. Fuente unica: hoja 3 de
' "03. Marcas de Fondos v5 vb222.xlsx" (se abre en solo lectura, no se escribe).
'
'   Config!C2 = TRAD  -> solo estrategia "F. Trad", cuadro partido por moneda
'   Config!C2 = ALT   -> todo lo demas, LOCALES / INTERNACIONALES y luego estrategia
'
' Hojas: Config - Clasif - Mapa - BD - AUM - PPT
' Macros publicas: CrearConfig / CrearClasif / ActualizarTodo / Recalcular / ArchivarAhora
'
' Reglas que no se cambian:
'   Eje = Dia de la marca (col B). Contribuciones se suman, no se componen.
'   No se recalcula el Basico: solo se multiplica por 10,000 (Config C9).
'   El ruteo TRAD / ALT y el bloque salen de la hoja Clasif, no de Exposicion.
'   Los duplicados se marcan REVISAR, no se anulan solos.
'   El archivado nunca sobrescribe.
'
' Requiere Excel 2019 o 365 (usa MAXIFS).
'==============================================================================

Private Const SH_CFG As String = "Config"
Private Const SH_CLA As String = "Clasif"
Private Const SH_MAP As String = "Mapa"
Private Const SH_BD  As String = "BD"
Private Const SH_AUM As String = "AUM"
Private Const SH_PPT As String = "PPT"

Private Const F_FIN As Long = 4        ' fila fecha fin
Private Const F_BAS As Long = 5        ' fila fecha base
Private Const F_HDR As Long = 6        ' fila cabecera
Private Const F_DAT As Long = 7        ' primera fila de datos

Private Const C_ETQ As Long = 2        ' columna etiqueta
Private Const C_INI As Long = 3        ' primera columna de datos (Ene)
Private Const C_FIN As Long = 18       ' ultima columna de datos (FY)
Private Const C_SUB As Long = 19       ' columna oculta con el subgrupo

Private Const SINCLAS As String = "SIN CLASIFICAR"

Private Type tReg
    Dia As Double
    Nombre As String
    Codigo As String
    Bloque As String
    Subgrupo As String
    EstCod As String
    Exposicion As String
    Moneda As String
    Vintage As String
    VarAdj As Double
    P1 As Double: P2 As Double: P3 As Double
    B1 As Double: B2 As Double: B3 As Double
    Flag As String
End Type

Private R() As tReg
Private nR As Long

' tabla Clasif en memoria
Private cCod() As String, cNom() As String, cArch() As String, cBloq() As String
Private nC As Long

' fondo del selector y conteo de fondos excluidos por no tener posicion
Private gSel As Long
Private gExcl As Long

'==============================================================================
' UTILITARIOS
'==============================================================================
Private Function Cfg(celda As String) As Variant
    Cfg = ThisWorkbook.Worksheets(SH_CFG).Range(celda).Value
End Function

Private Function EsALT() As Boolean
    EsALT = (UCase$(Trim$(CStr(Cfg("C2")))) <> "TRAD")
End Function

Private Function Existe(nombre As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(nombre)
    On Error GoTo 0
    Existe = Not ws Is Nothing
End Function

Private Function Hoja(nombre As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(nombre)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = nombre
    End If
    Set Hoja = ws
End Function

Private Function Letra(col As Long) As String
    If col < 1 Then Letra = "": Exit Function
    Letra = Split(ThisWorkbook.Sheets(1).Cells(1, col).Address(True, False), "$")(0)
End Function

Private Function Num(x As Variant) As Double
    If IsNumeric(x) Then Num = CDbl(x) Else Num = 0
End Function

' fondo del selector: PPT!C2 manda; AUM lo espeja
Private Function FondoSel() As Long
    Dim v As Variant
    FondoSel = 2
    If Not Existe(SH_PPT) Then Exit Function
    v = ThisWorkbook.Worksheets(SH_PPT).Range("C2").Value
    If IsNumeric(v) Then
        If v >= 1 And v <= 3 Then FondoSel = CLng(v)
    End If
End Function

' peso y contribucion del fondo seleccionado para una marca
Private Function PesoSel(i As Long) As Double
    Select Case gSel
        Case 1: PesoSel = R(i).P1
        Case 3: PesoSel = R(i).P3
        Case Else: PesoSel = R(i).P2
    End Select
End Function

Private Function PbSel(i As Long) As Double
    Select Case gSel
        Case 1: PbSel = R(i).B1
        Case 3: PbSel = R(i).B3
        Case Else: PbSel = R(i).B2
    End Select
End Function

' normaliza para comparar codigos: mayusculas, sin espacios dobles ni no-break space
Private Function Norm(s As String) As String
    Norm = UCase$(Application.WorksheetFunction.Trim(Replace(s, Chr(160), " ")))
End Function

' devuelve array 0-based, ordenado, con los sufijos de las claves que empiezan con prefijo
Private Function Ordenar(d As Object, prefijo As String) As Variant
    Dim k As Variant, tmp() As String, n As Long, i As Long, j As Long, sw As String
    If d.Count = 0 Then Ordenar = Array(): Exit Function
    ReDim tmp(0 To d.Count - 1)
    n = 0
    For Each k In d.Keys
        If prefijo = "" Then
            tmp(n) = CStr(k): n = n + 1
        ElseIf Left$(CStr(k), Len(prefijo)) = prefijo Then
            tmp(n) = Mid$(CStr(k), Len(prefijo) + 1): n = n + 1
        End If
    Next k
    If n = 0 Then Ordenar = Array(): Exit Function
    ReDim Preserve tmp(0 To n - 1)
    For i = 0 To n - 2
        For j = i + 1 To n - 1
            If tmp(j) < tmp(i) Then sw = tmp(i): tmp(i) = tmp(j): tmp(j) = sw
        Next j
    Next i
    Ordenar = tmp
End Function

'==============================================================================
' 1) CONFIG
'==============================================================================
Public Sub CrearConfig()
    Dim ws As Worksheet, ruta As String, modo As String, i As Long
    Set ws = Hoja(SH_CFG)
    ruta = CStr(ws.Range("C3").Value)
    modo = UCase$(Trim$(CStr(ws.Range("C2").Value)))
    If modo <> "TRAD" Then modo = "ALT"

    ws.Cells.Clear
    ws.Range("B1").Value = "CONTRIBUCION PROFUTURO — parametros"
    ws.Range("B1").Font.Bold = True

    ws.Range("B2").Value = "Libro: TRAD (tradicionales) / ALT (alternativos)": ws.Range("C2").Value = modo
    ws.Range("B3").Value = "Ruta de 03. Marcas de Fondos v5 vb222.xlsx":       ws.Range("C3").Value = ruta
    ws.Range("B4").Value = "Hoja fuente":                                      ws.Range("C4").Value = "3"
    ws.Range("B5").Value = "Fecha inicio del historico":                       ws.Range("C5").Value = DateSerial(2025, 1, 1)
    ws.Range("B6").Value = "Fecha de corte (vacio = ultima marca)":            ws.Range("C6").Value = ""
    ws.Range("B7").Value = "Fila de cabecera en la fuente":                    ws.Range("C7").Value = 3
    ws.Range("B8").Value = "Primera fila de datos en la fuente":               ws.Range("C8").Value = 4
    ws.Range("B9").Value = "Factor a pb":                                      ws.Range("C9").Value = 10000
    ws.Range("B10").Value = "Anio del cuadro (vacio = anio del corte)":        ws.Range("C10").Value = ""
    ws.Range("B11").Value = "Inicio del periodo libre (MayTD)":                ws.Range("C11").Value = DateSerial(2026, 5, 1)
    ws.Range("B12").Value = "Carpeta de archivo (vacio = subcarpeta Resumenes)": ws.Range("C12").Value = ""
    ws.Range("B13").Value = "Formato archivo: XLSX / PDF / AMBOS / NO":        ws.Range("C13").Value = "XLSX"

    ws.Range("B16").Value = "Mapa de campos — LETRA de columna en la hoja 3"
    ws.Range("B16").Font.Bold = True

    Dim et As Variant, df As Variant
    et = Array("Nombre del fondo", "Dia (fecha de la marca)", "Fecha EEFF", "Var Adj", "Codigo SBS", _
               "Peso F1 (GAP PRF F1)", "Peso F2 (GAP PRF2)", "Peso F3 (GAP PRF3)", _
               "Contribucion F1 (Basico PRF)", "Contribucion F2 (Basico PRF2)", "Contribucion F3 (Basico PRF3)", _
               "Estrategia", "Exposicion", "Vintage", "Moneda")
    df = Array("A", "B", "C", "F", "G", "I", "S", "AC", "N", "X", "AH", "CF", "CG", "CH", "CI")
    For i = 0 To 14
        ws.Cells(17 + i, 2).Value = et(i)
        ws.Cells(17 + i, 3).Value = df(i)
    Next i

    ws.Columns("B").ColumnWidth = 48
    ws.Columns("C").ColumnWidth = 32
    ws.Range("C5,C6,C11").NumberFormat = "dd/mm/yyyy"
    With ws.Range("C2").Validation
        .Delete
        .Add Type:=xlValidateList, Formula1:="TRAD,ALT"
    End With
    ws.Range("C2:C13").Interior.Color = RGB(255, 242, 204)

    MsgBox "Config creada." & vbCrLf & _
           "1) Elige el libro en C2 (TRAD o ALT)." & vbCrLf & _
           "2) Pon la ruta del archivo de marcas en C3." & vbCrLf & _
           "3) Corre CrearClasif y luego ActualizarTodo.", vbInformation
End Sub

'==============================================================================
' 2) CLASIF — ruteo por estrategia
'    El orden de esta tabla es el orden de las estrategias en el cuadro.
'==============================================================================
Public Sub CrearClasif()
    Dim ws As Worksheet, t As Variant, i As Long
    If Existe(SH_CLA) Then
        If MsgBox("La hoja Clasif ya existe. Se va a reescribir con la tabla base." & vbCrLf & _
                  "Perderias los codigos que hayas agregado a mano. Continuar?", _
                  vbQuestion + vbYesNo) = vbNo Then Exit Sub
    End If
    Set ws = Hoja(SH_CLA)
    ws.Cells.Clear
    ws.Range("A1:D1").Value = Array("Codigo (col CF)", "Estrategia", "Archivo", "Bloque")
    With ws.Range("A1:D1")
        .Font.Bold = True: .Interior.Color = RGB(139, 26, 26): .Font.Color = vbWhite
    End With

    t = Array( _
        Array("F. Trad", "Fondo Tradicional", "TRAD", ""), _
        Array("PE Int. Direc.", "Private Equity Internacional Directo", "ALT", "INTERNACIONALES"), _
        Array("PE Int. FoF", "Private Equity Internacional FoF", "ALT", "INTERNACIONALES"), _
        Array("PE Int. Sec.", "Private Equity Secundarios", "ALT", "INTERNACIONALES"), _
        Array("PE Int. Coinv.", "Private Equity Co-Inversiones", "ALT", "INTERNACIONALES"), _
        Array("PD Int.", "Private Debt", "ALT", "INTERNACIONALES"), _
        Array("Infra Int. Direc.", "Infraestructura Internacional Directo", "ALT", "INTERNACIONALES"), _
        Array("RE Int. Direc.", "Real Estate Directo", "ALT", "INTERNACIONALES"), _
        Array("RE Int. Sec.", "Real Estate Secundarios", "ALT", "INTERNACIONALES"), _
        Array("PE Loc. Direc.", "Private Equity Local Directo", "ALT", "LOCALES"), _
        Array("PD Loc.", "Private Debt Local", "ALT", "LOCALES"), _
        Array("Infra Loc. Direc.", "Infraestructura Directo", "ALT", "LOCALES"), _
        Array("RE Loc. Direc.", "Real Estate Local Directo", "ALT", "LOCALES"))

    For i = 0 To UBound(t)
        ws.Cells(2 + i, 1).Resize(1, 4).Value = t(i)
    Next i

    ws.Range("F2").Value = "El orden de las filas es el orden de las estrategias dentro de cada bloque."
    ws.Range("F3").Value = "Un codigo de la hoja 3 que no este aca cae en " & SINCLAS & " y se marca REVISAR en BD."
    ws.Range("F4").Value = "Archivo: TRAD o ALT.  Bloque: LOCALES o INTERNACIONALES (vacio para TRAD)."
    ws.Range("F2:F4").Font.Italic = True
    ws.Columns("A:D").AutoFit
End Sub

Private Sub CargarClasif()
    Dim ws As Worksheet, u As Long, i As Long
    If Not Existe(SH_CLA) Then CrearClasif
    Set ws = ThisWorkbook.Worksheets(SH_CLA)
    u = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If u < 2 Then Err.Raise 1010, , "La hoja Clasif esta vacia. Corre CrearClasif."
    nC = u - 1
    ReDim cCod(1 To nC): ReDim cNom(1 To nC): ReDim cArch(1 To nC): ReDim cBloq(1 To nC)
    For i = 1 To nC
        cCod(i) = Norm(CStr(ws.Cells(i + 1, 1).Value))
        cNom(i) = Trim$(CStr(ws.Cells(i + 1, 2).Value))
        cArch(i) = UCase$(Trim$(CStr(ws.Cells(i + 1, 3).Value)))
        cBloq(i) = UCase$(Trim$(CStr(ws.Cells(i + 1, 4).Value)))
        If cNom(i) = "" Then cNom(i) = Trim$(CStr(ws.Cells(i + 1, 1).Value))
        If cArch(i) <> "TRAD" Then cArch(i) = "ALT"
    Next i
End Sub

Private Function BuscarClasif(cod As String) As Long
    Dim i As Long, c As String
    c = Norm(cod)
    BuscarClasif = 0
    If c = "" Then Exit Function
    For i = 1 To nC
        If cCod(i) = c Then BuscarClasif = i: Exit Function
    Next i
End Function

'==============================================================================
' 3) CARGA DESDE LA HOJA 3
'==============================================================================
Private Sub CargarFuente()
    Dim rutaS As String, wbS As Workbook, wsS As Worksheet, ws As Worksheet, wsM As Worksheet
    Dim col(1 To 15) As Long, i As Long, filaHdr As Long, filaDat As Long, falta As Boolean

    CargarClasif
    Set ws = ThisWorkbook.Worksheets(SH_CFG)
    rutaS = Trim$(CStr(ws.Range("C3").Value))
    If rutaS = "" Then
        rutaS = Application.GetOpenFilename("Excel,*.xls*", , "Selecciona el archivo de marcas")
        If rutaS = "False" Then Err.Raise 1001, , "Sin libro fuente."
        ws.Range("C3").Value = rutaS
    End If

    Set wbS = Workbooks.Open(rutaS, ReadOnly:=True, UpdateLinks:=0)
    On Error GoTo Cerrar
    Set wsS = wbS.Worksheets(CStr(ws.Range("C4").Value))
    filaHdr = CLng(ws.Range("C7").Value)
    filaDat = CLng(ws.Range("C8").Value)

    For i = 1 To 15
        col(i) = 0
        If Trim$(CStr(ws.Cells(16 + i, 3).Value)) <> "" Then
            On Error Resume Next
            col(i) = wsS.Range(CStr(ws.Cells(16 + i, 3).Value) & "1").Column
            On Error GoTo Cerrar
        End If
    Next i

    '---- Mapa: primer control ----
    Set wsM = Hoja(SH_MAP)
    wsM.Cells.Clear
    wsM.Range("A1:E1").Value = Array("Campo", "Config", "Columna", "Cabecera real", "Estado")
    With wsM.Range("A1:E1")
        .Font.Bold = True: .Interior.Color = RGB(139, 26, 26): .Font.Color = vbWhite
    End With
    For i = 1 To 15
        wsM.Cells(1 + i, 1).Value = ws.Cells(16 + i, 2).Value
        wsM.Cells(1 + i, 2).Value = ws.Cells(16 + i, 3).Value
        wsM.Cells(1 + i, 3).Value = Letra(col(i))
        If col(i) > 0 Then wsM.Cells(1 + i, 4).Value = wsS.Cells(filaHdr, col(i)).Value
        If col(i) = 0 And Trim$(CStr(ws.Cells(16 + i, 3).Value)) <> "" Then
            wsM.Cells(1 + i, 5).Value = "NO ENCONTRADO"
            wsM.Cells(1 + i, 1).Resize(1, 5).Interior.Color = RGB(255, 190, 190)
            falta = True
        ElseIf col(i) = 0 Then
            wsM.Cells(1 + i, 5).Value = "no usado"
        Else
            wsM.Cells(1 + i, 5).Value = "OK"
        End If
    Next i
    wsM.Columns("A:E").AutoFit
    If falta Then Err.Raise 1002, , "Hay campos sin resolver. Revisa la hoja Mapa."

    '---- Lectura a memoria ----
    Dim lastR As Long, lastC As Long, v As Variant, k As Long, dia As Double
    Dim fIni As Double, fac As Double, d As Object, dSin As Object, key As String
    Dim ix As Long, estCod As String, arch As String, bloq As String, nom As String
    Dim nTrad As Long, nAlt As Long, nSin As Long

    lastR = wsS.Cells(wsS.Rows.Count, col(2)).End(xlUp).Row
    For i = 1 To 15
        If col(i) > lastC Then lastC = col(i)
    Next i
    If lastR < filaDat Then Err.Raise 1003, , "La hoja fuente no tiene datos."
    v = wsS.Range(wsS.Cells(filaDat, 1), wsS.Cells(lastR, lastC)).Value

    fIni = 0
    If IsDate(ws.Range("C5").Value) Then fIni = CDbl(CDate(ws.Range("C5").Value))
    fac = CDbl(ws.Range("C9").Value)
    Set d = CreateObject("Scripting.Dictionary")
    Set dSin = CreateObject("Scripting.Dictionary")

    ReDim R(1 To UBound(v, 1))
    nR = 0
    For k = 1 To UBound(v, 1)
        dia = 0
        If IsDate(v(k, col(2))) Then dia = CDbl(CDate(v(k, col(2))))
        If dia >= fIni And dia > 0 Then

            estCod = ""
            If col(12) > 0 Then estCod = Trim$(CStr(v(k, col(12))))
            ix = BuscarClasif(estCod)
            If ix = 0 Then
                arch = "ALT": bloq = SINCLAS: nom = SINCLAS
                nSin = nSin + 1
                If Not dSin.Exists(Norm(estCod)) Then dSin.Add Norm(estCod), IIf(estCod = "", "(vacio)", estCod)
            Else
                arch = cArch(ix): bloq = cBloq(ix): nom = cNom(ix)
            End If
            If arch = "TRAD" Then nTrad = nTrad + 1 Else nAlt = nAlt + 1

            ' filtro del libro: TRAD entra si C2=TRAD, ALT entra si C2=ALT
            If (arch = "ALT") = EsALT() Then
                nR = nR + 1
                With R(nR)
                    .Dia = dia
                    .Nombre = Trim$(CStr(v(k, col(1))))
                    If col(5) > 0 Then .Codigo = Trim$(CStr(v(k, col(5))))
                    If col(4) > 0 Then .VarAdj = Num(v(k, col(4)))
                    If col(6) > 0 Then .P1 = Num(v(k, col(6)))
                    If col(7) > 0 Then .P2 = Num(v(k, col(7)))
                    If col(8) > 0 Then .P3 = Num(v(k, col(8)))
                    If col(9) > 0 Then .B1 = Num(v(k, col(9))) * fac
                    If col(10) > 0 Then .B2 = Num(v(k, col(10))) * fac
                    If col(11) > 0 Then .B3 = Num(v(k, col(11))) * fac
                    If col(13) > 0 Then .Exposicion = Trim$(CStr(v(k, col(13))))
                    If col(14) > 0 Then .Vintage = Trim$(CStr(v(k, col(14))))
                    If col(15) > 0 Then .Moneda = UCase$(Trim$(CStr(v(k, col(15)))))
                    .EstCod = estCod

                    If EsALT() Then
                        .Bloque = bloq
                        If .Bloque = "" Then .Bloque = SINCLAS
                        .Subgrupo = nom
                    Else
                        .Bloque = .Moneda
                        If .Bloque = "" Then .Bloque = "SIN MONEDA"
                        .Subgrupo = ""
                    End If

                    If ix = 0 Then .Flag = "REVISAR estrategia sin clasificar"
                    key = .Codigo & "|" & .Nombre & "|" & CStr(.Dia)
                    If d.Exists(key) Then
                        .Flag = Trim$(.Flag & " REVISAR duplicado")
                    Else
                        d.Add key, 1
                    End If
                End With
            End If
        End If
    Next k

    '---- resumen del ruteo, debajo del Mapa ----
    wsM.Range("A19").Value = "Ruteo por estrategia"
    wsM.Range("A19").Font.Bold = True
    wsM.Range("A20").Value = "Marcas a TRAD (Fondo Tradicional)": wsM.Range("B20").Value = nTrad
    wsM.Range("A21").Value = "Marcas a ALT (alternativos)":        wsM.Range("B21").Value = nAlt
    wsM.Range("A22").Value = "De las anteriores, sin clasificar":  wsM.Range("B22").Value = nSin
    wsM.Range("A23").Value = "Cargadas en ESTE libro":             wsM.Range("B23").Value = nR
    If nSin > 0 Then
        Dim kk As Variant, rr As Long
        wsM.Range("A22:B22").Interior.Color = RGB(255, 235, 156)
        wsM.Range("A25").Value = "Codigos que no estan en Clasif:"
        wsM.Range("A25").Font.Bold = True
        rr = 26
        For Each kk In dSin.Keys
            wsM.Cells(rr, 1).Value = dSin(kk): rr = rr + 1
        Next kk
    End If

    If nR = 0 Then Err.Raise 1004, , "No hay marcas para este libro (Config C2 = " & Cfg("C2") & ")."

Cerrar:
    Dim e As Long, m As String
    e = Err.Number: m = Err.Description
    On Error Resume Next
    wbS.Close SaveChanges:=False
    On Error GoTo 0
    If e <> 0 Then Err.Raise e, , m
End Sub

'==============================================================================
' 4) HOJA BD
'==============================================================================
Private Sub EscribirBD()
    Dim ws As Worksheet, a() As Variant, i As Long
    Set ws = Hoja(SH_BD)
    ws.Cells.Clear
    ws.Range("A1:S1").Value = Array("Dia", "Anio", "Mes", "Nombre", "Codigo SBS", "Bloque", "Subgrupo", _
                                    "Estrategia cod", "Exposicion", "Moneda", "Vintage", "Var Adj", _
                                    "Peso F1", "Peso F2", "Peso F3", "pb F1", "pb F2", "pb F3", "Flag")
    With ws.Range("A1:S1")
        .Font.Bold = True: .Interior.Color = RGB(139, 26, 26): .Font.Color = vbWhite
    End With

    ReDim a(1 To nR, 1 To 19)
    For i = 1 To nR
        a(i, 1) = R(i).Dia
        a(i, 2) = Year(CDate(R(i).Dia))
        a(i, 3) = Month(CDate(R(i).Dia))
        a(i, 4) = R(i).Nombre
        a(i, 5) = R(i).Codigo
        a(i, 6) = R(i).Bloque
        a(i, 7) = R(i).Subgrupo
        a(i, 8) = R(i).EstCod
        a(i, 9) = R(i).Exposicion
        a(i, 10) = R(i).Moneda
        a(i, 11) = R(i).Vintage
        a(i, 12) = R(i).VarAdj
        a(i, 13) = R(i).P1: a(i, 14) = R(i).P2: a(i, 15) = R(i).P3
        a(i, 16) = R(i).B1: a(i, 17) = R(i).B2: a(i, 18) = R(i).B3
        a(i, 19) = R(i).Flag
    Next i
    ws.Range("A2").Resize(nR, 19).Value = a

    ws.Columns("A").NumberFormat = "dd/mm/yyyy"
    ws.Columns("L:O").NumberFormat = "0.00%"
    ws.Columns("P:R").NumberFormat = "#,##0.0"
    If ws.AutoFilterMode Then ws.AutoFilterMode = False
    ws.Rows(1).AutoFilter
    ws.Range("A1:S1").EntireColumn.AutoFit
    DefinirNombres nR
End Sub

Private Sub PonerNombre(nombre As String, ref As String)
    On Error Resume Next
    ThisWorkbook.Names(nombre).Delete
    On Error GoTo 0
    ThisWorkbook.Names.Add nombre, ref
End Sub

Private Sub DefinirNombres(n As Long)
    Dim u As Long
    u = IIf(n < 1, 2, n + 1)
    PonerNombre "bdDia", "=" & SH_BD & "!$A$2:$A$" & u
    PonerNombre "bdNom", "=" & SH_BD & "!$D$2:$D$" & u
    PonerNombre "bdBloque", "=" & SH_BD & "!$F$2:$F$" & u
    PonerNombre "bdSub", "=" & SH_BD & "!$G$2:$G$" & u
    PonerNombre "bdPeso", "=" & SH_BD & "!$M$2:$O$" & u
    PonerNombre "bdPb", "=" & SH_BD & "!$P$2:$R$" & u
End Sub

Private Sub LeerBDdeHoja()
    Dim ws As Worksheet, u As Long, v As Variant, i As Long
    CargarClasif
    If Not Existe(SH_BD) Then Err.Raise 1005, , "No hay hoja BD. Corre ActualizarTodo."
    Set ws = ThisWorkbook.Worksheets(SH_BD)
    u = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If u < 2 Then Err.Raise 1005, , "BD vacia. Corre ActualizarTodo."
    v = ws.Range("A2:S" & u).Value
    nR = UBound(v, 1)
    ReDim R(1 To nR)
    For i = 1 To nR
        With R(i)
            .Dia = Num(v(i, 1)): .Nombre = CStr(v(i, 4)): .Codigo = CStr(v(i, 5))
            .Bloque = CStr(v(i, 6)): .Subgrupo = CStr(v(i, 7)): .EstCod = CStr(v(i, 8))
            .Exposicion = CStr(v(i, 9)): .Moneda = CStr(v(i, 10)): .Vintage = CStr(v(i, 11))
            .VarAdj = Num(v(i, 12))
            .P1 = Num(v(i, 13)): .P2 = Num(v(i, 14)): .P3 = Num(v(i, 15))
            .B1 = Num(v(i, 16)): .B2 = Num(v(i, 17)): .B3 = Num(v(i, 18))
        End With
    Next i
    DefinirNombres nR
End Sub

'==============================================================================
' 5) CUADROS PPT Y AUM
'==============================================================================
Private Sub ConstruirCuadros()
    Dim i As Long, m As Long, corte As Double, mx As Double, anio As Long, fyBase As Double
    Dim finM(1 To 12) As Double, basM(1 To 12) As Double, finMes As Double

    For i = 1 To nR
        If R(i).Dia > mx Then mx = R(i).Dia
    Next i
    If IsDate(Cfg("C6")) Then corte = CDbl(CDate(Cfg("C6"))) Else corte = mx
    If CStr(Cfg("C10")) <> "" And IsNumeric(Cfg("C10")) Then
        anio = CLng(Cfg("C10"))
    Else
        anio = Year(CDate(corte))
    End If

    ' fecha fin de cada mes = ultima marca del mes; si el mes ya paso sin marcas, fin de mes
    For m = 1 To 12
        For i = 1 To nR
            If Year(CDate(R(i).Dia)) = anio And Month(CDate(R(i).Dia)) = m And R(i).Dia <= corte Then
                If R(i).Dia > finM(m) Then finM(m) = R(i).Dia
            End If
        Next i
        finMes = CDbl(DateSerial(anio, m + 1, 0))
        If finM(m) = 0 And CDbl(DateSerial(anio, m, 1)) <= corte Then
            finM(m) = IIf(finMes < corte, finMes, corte)
        End If
    Next m
    For m = 1 To 12
        If m = 1 Then
            basM(m) = CDbl(DateSerial(anio - 1, 12, 31))
        ElseIf finM(m - 1) > 0 Then
            basM(m) = finM(m - 1)
        Else
            basM(m) = CDbl(DateSerial(anio, m, 0))
        End If
    Next m
    fyBase = CDbl(DateSerial(IIf(Month(CDate(corte)) >= 11, Year(CDate(corte)), Year(CDate(corte)) - 1), 10, 31))

    '---- que filas entran ----
    ' Solo los fondos en los que tenemos posicion en el fondo del selector: peso <> 0
    ' en alguna marca de la ventana mostrada. Un fondo con peso 0 contribuye 0 pb, asi
    ' que excluirlo no mueve el Total. Se mira toda la ventana, no solo el corte, para
    ' no perder un fondo que se tuvo y se vendio en el camino.
    Dim db As Object, ds As Object, dn As Object, dTodos As Object, kFondo As String
    Set db = CreateObject("Scripting.Dictionary")
    Set ds = CreateObject("Scripting.Dictionary")
    Set dn = CreateObject("Scripting.Dictionary")
    Set dTodos = CreateObject("Scripting.Dictionary")

    gSel = FondoSel()
    For i = 1 To nR
        If R(i).Dia <= corte And (Year(CDate(R(i).Dia)) = anio Or R(i).Dia > fyBase) Then
            kFondo = R(i).Bloque & "|" & R(i).Subgrupo & "|" & R(i).Nombre
            If Not dTodos.Exists(kFondo) Then dTodos.Add kFondo, 1
            If PesoSel(i) <> 0 Or PbSel(i) <> 0 Then
                If Not db.Exists(R(i).Bloque) Then db.Add R(i).Bloque, 1
                If Not ds.Exists(R(i).Bloque & "|" & R(i).Subgrupo) Then ds.Add R(i).Bloque & "|" & R(i).Subgrupo, 1
                If Not dn.Exists(kFondo) Then dn.Add kFondo, 1
            End If
        End If
    Next i
    gExcl = dTodos.Count - dn.Count
    If dn.Count = 0 Then Err.Raise 1006, , "No hay fondos con posicion en el Fondo " & gSel & " en la ventana mostrada."

    Dim ordenBl As Variant
    If EsALT() Then
        ordenBl = Array("INTERNACIONALES", "LOCALES", SINCLAS)
    Else
        ordenBl = Array("PEN", "USD")
    End If

    Dim filas As Collection, ya As Object, o As Variant, b As Variant
    Set filas = New Collection
    Set ya = CreateObject("Scripting.Dictionary")
    For Each o In ordenBl
        If db.Exists(CStr(o)) Then
            AgregarBloque filas, CStr(o), ds, dn
            ya.Add CStr(o), 1
        End If
    Next o
    For Each b In Ordenar(db, "")
        If Not ya.Exists(CStr(b)) Then AgregarBloque filas, CStr(b), ds, dn
    Next b
    filas.Add Array("", "", "Total", 0)

    PintarHoja Hoja(SH_PPT), filas, finM, basM, corte, anio, True
    PintarHoja Hoja(SH_AUM), filas, finM, basM, corte, anio, False
End Sub

' Agrega al listado: fila de bloque, y dentro las estrategias (solo ALT) con sus fondos.
' Cada fila es Array(bloque, subgrupo, etiqueta, nivel)  nivel: 1 bloque, 2 estrategia, 3 fondo, 0 total
Private Sub AgregarBloque(filas As Collection, bloque As String, ds As Object, dn As Object)
    Dim s As Variant, t As Variant, i As Long, ya As Object, subs As Variant

    filas.Add Array(bloque, "", bloque, 1)

    If Not EsALT() Then                       ' TRAD: fondos directo bajo la moneda
        For Each t In Ordenar(dn, bloque & "||")
            filas.Add Array(bloque, "", CStr(t), 3)
        Next t
        Exit Sub
    End If

    ' ALT: estrategias en el orden de la hoja Clasif, y lo que no este, alfabetico al final
    Set ya = CreateObject("Scripting.Dictionary")
    For i = 1 To nC
        If cArch(i) = "ALT" And cBloq(i) = bloque Then
            If ds.Exists(bloque & "|" & cNom(i)) And Not ya.Exists(cNom(i)) Then
                filas.Add Array(bloque, cNom(i), cNom(i), 2)
                For Each t In Ordenar(dn, bloque & "|" & cNom(i) & "|")
                    filas.Add Array(bloque, cNom(i), CStr(t), 3)
                Next t
                ya.Add cNom(i), 1
            End If
        End If
    Next i

    subs = Ordenar(ds, bloque & "|")
    If IsArray(subs) Then
        For Each s In subs
            If Not ya.Exists(CStr(s)) Then
                filas.Add Array(bloque, CStr(s), CStr(s), 2)
                For Each t In Ordenar(dn, bloque & "|" & CStr(s) & "|")
                    filas.Add Array(bloque, CStr(s), CStr(t), 3)
                Next t
                ya.Add CStr(s), 1
            End If
        Next s
    End If
End Sub

Private Sub PintarHoja(ws As Worksheet, filas As Collection, finM() As Double, basM() As Double, _
                       corte As Double, anio As Long, esPPT As Boolean)
    Dim m As Long, i As Long, j As Long, c As Long, r As Long, cl As String
    Dim f As Variant, niv As Long, u As Range
    Dim frm(1 To 1, 1 To 16) As Variant
    Dim hijos As String

    Application.EnableEvents = False
    ws.Cells.Clear
    ws.Cells.FormatConditions.Delete
    With ws.Cells.Font
        .Name = "Arial": .Size = 8
    End With

    '---- cabecera ----
    ws.Range("B2").Formula = "=""Fondo ""&$C$2"
    ws.Range("B2").Font.Bold = True
    If esPPT Then
        ws.Range("C2").Value = gSel
        With ws.Range("C2").Validation
            .Delete
            .Add Type:=xlValidateList, Formula1:="1,2,3"
        End With
        ws.Range("C2").Interior.Color = RGB(255, 242, 204)
    Else
        ws.Range("C2").Formula = "=" & SH_PPT & "!$C$2"   ' AUM espeja el selector de PPT
        ws.Range("C2").Interior.Color = RGB(242, 242, 242)
    End If
    ws.Range("D2").Value = IIf(esPPT, "Contribucion Profuturo, en pb", "Posicionamiento del fondo, %")
    ws.Range("D2").Font.Italic = True
    If gExcl > 0 Then
        ws.Range("F2").Value = "Solo fondos con posicion en el Fondo " & gSel & ". " & _
                               gExcl & " fondo(s) de la hoja 3 quedaron fuera por peso 0."
        ws.Range("F2").Font.Italic = True
        ws.Range("F2").Font.Color = RGB(120, 120, 120)
    End If
    ws.Cells(F_FIN, 1).Value = "Fecha fin"
    ws.Cells(F_BAS, 1).Value = "Fecha base"
    ws.Cells(F_HDR, C_ETQ).Value = "Instrumento"

    Dim cab As Variant
    cab = Array("Ene", "Feb", "Mar", "Abr", "May", "Jun", "Jul", "Ago", "Sep", "Oct", "Nov", "Dic", _
                "MTD", "MayTD", "YTD", "FY")
    For m = 1 To 12
        c = C_INI + m - 1
        ws.Cells(F_HDR, c).Value = cab(m - 1)
        If finM(m) > 0 Then
            ws.Cells(F_FIN, c).Value = CDate(finM(m))
            ws.Cells(F_BAS, c).Value = CDate(basM(m))
        End If
    Next m

    Dim basX(1 To 4) As Double
    basX(1) = CDbl(DateSerial(Year(CDate(corte)), Month(CDate(corte)), 0))     ' MTD
    basX(2) = CDbl(CDate(Cfg("C11"))) - 1                                      ' MayTD
    basX(3) = CDbl(DateSerial(anio - 1, 12, 31))                               ' YTD
    basX(4) = CDbl(DateSerial(IIf(Month(CDate(corte)) >= 11, Year(CDate(corte)), _
                                  Year(CDate(corte)) - 1), 10, 31))            ' FY
    For i = 1 To 4
        c = C_INI + 11 + i
        ws.Cells(F_HDR, c).Value = cab(11 + i)
        ws.Cells(F_FIN, c).Value = CDate(corte)
        ws.Cells(F_BAS, c).Value = CDate(basX(i))
    Next i

    With ws.Range(ws.Cells(F_HDR, C_ETQ), ws.Cells(F_HDR, C_FIN))
        .Interior.Color = RGB(139, 26, 26)
        .Font.Color = vbWhite
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
    End With
    With ws.Range(ws.Cells(F_FIN, C_INI), ws.Cells(F_BAS, C_FIN))
        .NumberFormat = "dd/mm/yyyy"
        .Font.Size = 7
        .Font.Italic = True
        .HorizontalAlignment = xlCenter
    End With

    '---- filas ----
    For i = 1 To filas.Count
        f = filas(i)
        niv = CLng(f(3))
        r = F_DAT + i - 1

        ws.Cells(r, 1).Value = f(0)          ' bloque, columna oculta A
        ws.Cells(r, C_ETQ).Value = f(2)      ' etiqueta
        ws.Cells(r, C_SUB).Value = f(1)      ' subgrupo, columna oculta S

        If esPPT Then
            For j = 1 To 16
                cl = Letra(C_INI + j - 1)
                frm(1, j) = FrmPb(cl, r, niv)
            Next j
        Else
            hijos = ""
            If niv <> 3 Then hijos = FilasHijas(filas, i, niv)
            For j = 1 To 16
                cl = Letra(C_INI + j - 1)
                If niv = 3 Then
                    frm(1, j) = FrmPeso(cl, r)
                Else
                    frm(1, j) = FrmPesoAgg(cl, hijos)
                End If
            Next j
        End If
        ws.Range(ws.Cells(r, C_INI), ws.Cells(r, C_FIN)).Formula = frm

        Select Case niv
            Case 1
                With ws.Range(ws.Cells(r, C_ETQ), ws.Cells(r, C_FIN))
                    .Interior.Color = RGB(60, 60, 60): .Font.Color = vbWhite: .Font.Bold = True
                End With
            Case 2
                With ws.Range(ws.Cells(r, C_ETQ), ws.Cells(r, C_FIN))
                    .Interior.Color = RGB(217, 217, 217): .Font.Bold = True
                End With
                ws.Cells(r, C_ETQ).IndentLevel = 1
            Case 0
                With ws.Range(ws.Cells(r, C_ETQ), ws.Cells(r, C_FIN))
                    .Interior.Color = RGB(191, 191, 191): .Font.Bold = True
                End With
            Case Else
                ws.Cells(r, C_ETQ).IndentLevel = IIf(EsALT(), 2, 1)
                ws.Range(ws.Cells(r, C_FIN - 3), ws.Cells(r, C_FIN)).Interior.Color = RGB(251, 248, 230)
                If esPPT Then
                    If u Is Nothing Then
                        Set u = ws.Range(ws.Cells(r, C_INI), ws.Cells(r, C_INI + 11))
                    Else
                        Set u = Union(u, ws.Range(ws.Cells(r, C_INI), ws.Cells(r, C_INI + 11)))
                    End If
                End If
        End Select
    Next i
    r = F_DAT + filas.Count - 1

    With ws.Range(ws.Cells(F_DAT, C_INI), ws.Cells(r, C_FIN))
        .NumberFormat = IIf(esPPT, "#,##0.0;[Red]-#,##0.0", "0.00%")
        .HorizontalAlignment = xlCenter
    End With
    ws.Range(ws.Cells(F_DAT, C_ETQ), ws.Cells(r, C_FIN)).Borders.Color = RGB(160, 160, 160)

    If esPPT And Not u Is Nothing Then
        With u.FormatConditions.AddColorScale(3)
            .ColorScaleCriteria(1).Type = xlConditionValueLowestValue
            .ColorScaleCriteria(1).FormatColor.Color = RGB(248, 105, 107)
            .ColorScaleCriteria(2).Type = xlConditionValueNumber
            .ColorScaleCriteria(2).Value = 0
            .ColorScaleCriteria(2).FormatColor.Color = vbWhite
            .ColorScaleCriteria(3).Type = xlConditionValueHighestValue
            .ColorScaleCriteria(3).FormatColor.Color = RGB(99, 190, 123)
        End With
    End If

    ws.Columns(1).ColumnWidth = 2
    ws.Columns(C_ETQ).ColumnWidth = 34
    ws.Range(ws.Columns(C_INI), ws.Columns(C_FIN)).ColumnWidth = 9
    ws.Columns(1).Hidden = True
    ws.Columns(C_SUB).Hidden = True

    ' congelar paneles sin Select: Select falla si la hoja no esta activa
    On Error Resume Next
    If ThisWorkbook.Windows.Count > 0 Then
        ws.Activate
        With ActiveWindow
            .FreezePanes = False
            .ScrollRow = 1
            .ScrollColumn = 1
            .SplitRow = F_DAT - 1
            .SplitColumn = C_ETQ
            .FreezePanes = True
        End With
    End If
    On Error GoTo 0
    Application.EnableEvents = True
End Sub

' filas de detalle (nivel 3) que cuelgan de la fila idx; devuelve "10,11,14"
Private Function FilasHijas(filas As Collection, idx As Long, niv As Long) As String
    Dim j As Long, f As Variant, nj As Long, s As String
    If niv = 0 Then                       ' Total: todos los detalles del cuadro
        For j = 1 To filas.Count
            f = filas(j)
            If CLng(f(3)) = 3 Then s = s & IIf(s = "", "", ",") & CStr(F_DAT + j - 1)
        Next j
    Else
        For j = idx + 1 To filas.Count
            f = filas(j)
            nj = CLng(f(3))
            If nj <> 3 And nj <= niv Then Exit For
            If nj = 3 Then s = s & IIf(s = "", "", ",") & CStr(F_DAT + j - 1)
        Next j
    End If
    FilasHijas = s
End Function

' criterios por nivel: 1 bloque | 2 bloque+subgrupo | 3 bloque+nombre | 0 todo
Private Function Crit(r As Long, niv As Long) As String
    Select Case niv
        Case 1: Crit = ",bdBloque,$A" & r
        Case 2: Crit = ",bdBloque,$A" & r & ",bdSub,$S" & r
        Case 3: Crit = ",bdBloque,$A" & r & ",bdNom,$B" & r
        Case Else: Crit = ""
    End Select
End Function

' contribucion: suma de los pb de las marcas del periodo (aditivo, no se compone)
Private Function FrmPb(cl As String, r As Long, niv As Long) As String
    FrmPb = "=IF(" & cl & "$" & F_FIN & "="""",NA()," & _
            "SUMIFS(INDEX(bdPb,0,$C$2),bdDia,"">""&" & cl & "$" & F_BAS & _
            ",bdDia,""<=""&" & cl & "$" & F_FIN & Crit(r, niv) & "))"
End Function

' peso de un fondo: el de su ultima marca con fecha <= fecha fin
Private Function FrmPeso(cl As String, r As Long) As String
    Dim ult As String
    ult = "MAXIFS(bdDia,bdDia,""<=""&" & cl & "$" & F_FIN & Crit(r, 3) & ")"
    FrmPeso = "=IF(" & cl & "$" & F_FIN & "="""",NA(),IF(" & ult & "=0,0," & _
              "SUMIFS(INDEX(bdPeso,0,$C$2),bdDia," & ult & Crit(r, 3) & ")))"
End Function

' peso agregado: suma de los pesos de los fondos que cuelgan, no un MAXIFS del grupo
Private Function FrmPesoAgg(cl As String, hijos As String) As String
    Dim p() As String, i As Long, s As String
    If hijos = "" Then FrmPesoAgg = "=IF(" & cl & "$" & F_FIN & "="""",NA(),0)": Exit Function
    p = Split(hijos, ",")
    For i = LBound(p) To UBound(p)
        s = s & IIf(s = "", "", ",") & cl & p(i)
    Next i
    FrmPesoAgg = "=IF(" & cl & "$" & F_FIN & "="""",NA(),SUM(" & s & "))"
End Function

'==============================================================================
' 6) ARCHIVADO — nunca sobrescribe
'==============================================================================
Public Sub ArchivarAhora()
    Dim fmt As String, carp As String, base As String, ruta As String, i As Long, wbN As Workbook
    fmt = UCase$(Trim$(CStr(Cfg("C13"))))
    If fmt = "NO" Then Exit Sub

    carp = Trim$(CStr(Cfg("C12")))
    If carp = "" Then carp = ThisWorkbook.Path & Application.PathSeparator & "Resumenes"
    If Dir(carp, vbDirectory) = "" Then MkDir carp
    base = carp & Application.PathSeparator & "Contribucion_" & _
           IIf(EsALT(), "Alternativos_", "Tradicionales_") & Format(Date, "yyyymmdd")

    Application.ScreenUpdating = False
    ThisWorkbook.Sheets(Array(SH_PPT, SH_AUM)).Copy
    Set wbN = ActiveWorkbook
    For i = 1 To wbN.Worksheets.Count
        With wbN.Worksheets(i).UsedRange
            .Value = .Value                       ' pega valores, corta los vinculos
        End With
        wbN.Worksheets(i).Columns(1).Hidden = True
        wbN.Worksheets(i).Columns(C_SUB).Hidden = True
    Next i

    ruta = base: i = 1
    Do While Dir(ruta & ".xlsx") <> ""
        i = i + 1
        ruta = base & "_v" & i
    Loop
    If fmt = "XLSX" Or fmt = "AMBOS" Then wbN.SaveAs ruta & ".xlsx", xlOpenXMLWorkbook
    If fmt = "PDF" Or fmt = "AMBOS" Then wbN.ExportAsFixedFormat xlTypePDF, ruta & ".pdf"
    wbN.Close SaveChanges:=False
    Application.ScreenUpdating = True
End Sub

'==============================================================================
' 7) ORQUESTACION
'==============================================================================
Public Sub ActualizarTodo()
    Dim t0 As Double, tFuente As Double, tBD As Double, tCuadros As Double
    On Error GoTo Fin
    t0 = Timer
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    Application.StatusBar = "1/4 Leyendo la hoja 3..."
    CargarFuente
    tFuente = Timer - t0: t0 = Timer

    Application.StatusBar = "2/4 Escribiendo BD..."
    EscribirBD
    tBD = Timer - t0: t0 = Timer

    Application.StatusBar = "3/4 Armando PPT y AUM..."
    ConstruirCuadros
    tCuadros = Timer - t0

    Application.StatusBar = "4/4 Archivando..."
    ArchivarAhora

Fin:
    Application.StatusBar = False
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    If Err.Number <> 0 Then
        MsgBox "Error " & Err.Number & ": " & Err.Description, vbExclamation, "ActualizarTodo"
    Else
        MsgBox nR & " marcas cargadas en el libro " & Cfg("C2") & "." & vbCrLf & vbCrLf & _
               "Fuente " & Format(tFuente, "0.0") & "s  |  BD " & Format(tBD, "0.0") & "s  |  Cuadros " & _
               Format(tCuadros, "0.0") & "s" & vbCrLf & vbCrLf & _
               "Revisa la hoja Mapa antes de usar el cuadro.", vbInformation, "Listo"
    End If
End Sub

' Rearma los cuadros con lo que ya esta en BD. No vuelve a la red.
Public Sub Recalcular()
    On Error GoTo Fin
    Application.ScreenUpdating = False
    LeerBDdeHoja
    ConstruirCuadros
Fin:
    Application.ScreenUpdating = True
    If Err.Number <> 0 Then MsgBox "Error " & Err.Number & ": " & Err.Description, vbExclamation, "Recalcular"
End Sub
