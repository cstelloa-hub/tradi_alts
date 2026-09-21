Attribute VB_Name = "mod_RetornoFondos"
'===============================================================================
' mod_RetornoFondos  -  v1.9  -  21/09/2026
'
' Cuadro de RETORNOS por fondo (no contribucion) a partir de la hoja 3 de
' "03. Marcas de Fondos v5 vb222.xlsx".
'
' Dos libros, este mismo modulo en los dos. Lo unico que cambia es Config!C2:
'    Retornos_Tradicionales.xlsm   C2 = TRAD   -> bloque = Moneda (col CI) -> fondos
'    Retornos_Alternativos.xlsm    C2 = ALT    -> bloque = INTERNACIONALES / LOCALES
'                                                 -> Estrategia -> fondos
'
' Hojas que crea / mantiene:  Config . Clasif . Mapa . BD . CUADRO
'
' INSTALACION
'   1. Libro nuevo .xlsm vacio.
'   2. VBE (Alt+F11) > Archivo > Importar archivo... > mod_RetornoFondos.bas
'   3. Ejecutar CrearConfig   (crea Config y Clasif con los defaults)
'   4. Poner la ruta del libro de marcas en Config!C3 y el corte en C6.
'   5. Ejecutar ActualizarTodo.
'   6. Para el segundo libro: Guardar una copia, cambiar C2, ActualizarTodo.
'
' MACROS
'   CrearConfig     crea/repara Config y Clasif. Conserva C3. Pisa C6 y C11.
'   CrearClasif     rehace solo la tabla de estrategias.
'   ActualizarTodo  va a la red, relee la fuente, rehace BD y el CUADRO.
'   Recalcular      rehace CUADRO desde la BD ya cargada. No toca la red.
'   ArchivarAhora   guarda el CUADRO en valores en la carpeta de archivo.
'
' METODOLOGIA (fija)
'   - Eje = Dia de la marca (col B). Nunca la fecha EEFF.
'   - Retorno compuesto:  EXP( SUM( LN(1+VarAdj) ) ) - 1  sobre base < Dia <= fin.
'     La BD trae LN(1+VarAdj) precalculado (col M) para que el cuadro sea un
'     SUMIFS plano y la formula quede visible.
'   - Var Adj vacio = Var SBS (col E) cuando no hubo flujos.
'   - Pesos F1/F2/F3 = GAP PRF de las cols I / S / AC, el vigente en la ultima
'     marca <= corte.
'   - Filas de grupo: solo division. Suma de pesos, sin retorno.
'   - Nada se corrige en silencio: todo lo raro cae en Mapa como REVISAR.
'
' REQUISITOS: Excel 2019 / 365 (usa MAXIFS y SUMIFS).
'===============================================================================

Option Explicit

'--- nombres de hoja -----------------------------------------------------------
Private Const SH_CFG As String = "Config"
Private Const SH_CLA As String = "Clasif"
Private Const SH_MAP As String = "Mapa"
Private Const SH_BD  As String = "BD"
Private Const SH_CUA As String = "CUADRO"

'--- layout del CUADRO ---------------------------------------------------------
Private Const CU_FILA_FIN   As Long = 4     ' fecha fin de cada ventana
Private Const CU_FILA_BASE  As Long = 5     ' fecha base de cada ventana
Private Const CU_FILA_CAB   As Long = 6     ' cabecera
Private Const CU_FILA_DAT   As Long = 7     ' primera fila de datos
Private Const CU_COL_BLQ    As Long = 1     ' A  oculta - bloque
Private Const CU_COL_NOM    As Long = 2     ' B  nombre
Private Const CU_COL_P1     As Long = 3     ' C  peso F1
Private Const CU_COL_SEP    As Long = 6     ' F  separador
Private Const CU_COL_W1     As Long = 7     ' G  primera ventana
Private Const CU_COL_SUB    As Long = 13    ' M  oculta - subgrupo
Private Const N_VENT        As Long = 6     ' WTD 5D MTD MayoTD YTD FY

'--- layout de la BD -----------------------------------------------------------
Private Const BD_COLS As Long = 17
' 1 Dia  2 Anio  3 Mes  4 Nombre  5 CodSBS  6 Bloque  7 Subgrupo  8 EstrategiaCod
' 9 Exposicion  10 Moneda  11 Vintage  12 VarAdj  13 LN(1+VarAdj)
' 14 PesoF1  15 PesoF2  16 PesoF3  17 Flag

'--- estado de la corrida ------------------------------------------------------
Private mT0 As Double
Private mPasos As String
Private mDetalle As String


'===============================================================================
'  1. CONFIG Y CLASIF
'===============================================================================
Public Sub CrearConfig()
    Dim ws As Worksheet, rutaPrev As String, c2Prev As String
    Dim scrPrev As Boolean

    scrPrev = Application.ScreenUpdating
    Application.ScreenUpdating = False
    On Error GoTo Limpio

    ' conservar lo que ya estaba
    On Error Resume Next
    rutaPrev = CStr(ThisWorkbook.Worksheets(SH_CFG).Range("C3").Value)
    c2Prev = CStr(ThisWorkbook.Worksheets(SH_CFG).Range("C2").Value)
    On Error GoTo Limpio

    Set ws = HojaOCrea(SH_CFG)
    ws.Cells.Clear
    ws.Cells.Interior.Pattern = xlNone

    ws.Range("B1").Value = "CONFIGURACION - Cuadro de retornos por fondo"
    ws.Range("B1").Font.Bold = True
    ws.Range("B1").Font.Size = 13

    EscribeCfg ws, 2, "Libro (TRAD / ALT)", IIf(Len(c2Prev) > 0, c2Prev, "TRAD"), _
               "TRAD = solo estrategia F. Trad. ALT = todas las demas."
    EscribeCfg ws, 3, "Ruta del libro de marcas", rutaPrev, _
               "Vacio = pregunta al ejecutar ActualizarTodo."
    EscribeCfg ws, 4, "Hoja fuente", "3", "Nombre o indice de la hoja."
    EscribeCfg ws, 5, "Inicio del historico", DateSerial(2025, 1, 1), _
               "Se ignoran las marcas anteriores."
    EscribeCfg ws, 6, "Fecha de corte", Empty, _
               "Vacio = la ultima marca cargada."
    EscribeCfg ws, 7, "Fila de cabecera de la fuente", 3, ""
    EscribeCfg ws, 8, "Primera fila de datos de la fuente", 4, ""
    EscribeCfg ws, 9, "Dias de la ventana 5D", 5, "Dias calendario hacia atras."
    EscribeCfg ws, 10, "Base del FY", DateSerial(2025, 10, 31), _
               "FY = 01 nov a 31 oct."
    EscribeCfg ws, 11, "Inicio del periodo libre (MayoTD)", DateSerial(2026, 5, 1), _
               "La base del periodo es esta fecha menos un dia."
    EscribeCfg ws, 12, "Carpeta de archivo", "", _
               "Vacio = subcarpeta Resumenes junto a este libro."
    EscribeCfg ws, 13, "Formato de archivo", "XLSX", "XLSX / PDF / AMBOS / NO"

    ws.Range("B15").Value = "MAPA DE COLUMNAS DE LA FUENTE (letras)"
    ws.Range("B15").Font.Bold = True
    ws.Range("B16").Value = "Campo"
    ws.Range("C16").Value = "Letra"
    ws.Range("D16").Value = "Nota"
    ws.Range("B16:D16").Font.Bold = True

    EscribeMap ws, 17, "Nombre", "A", "Etiqueta del fondo"
    EscribeMap ws, 18, "Fecha (Dia de la marca)", "B", "Eje de todo"
    EscribeMap ws, 19, "Fecha EEFF", "C", "Solo trazabilidad"
    EscribeMap ws, 20, "Var SBS", "E", "Se usa solo si Var Adj viene vacio"
    EscribeMap ws, 21, "Var Adj", "F", "Retorno del fondo"
    EscribeMap ws, 22, "Codigo SBS", "G", "Llave, detecta duplicados"
    EscribeMap ws, 23, "GAP PRF F1", "I", "Peso propio en Fondo 1"
    EscribeMap ws, 24, "GAP PRF F2", "S", "Peso propio en Fondo 2"
    EscribeMap ws, 25, "GAP PRF F3", "AC", "Peso propio en Fondo 3"
    EscribeMap ws, 26, "Basico F1", "N", "No se usa en este cuadro"
    EscribeMap ws, 27, "Basico F2", "X", "No se usa en este cuadro"
    EscribeMap ws, 28, "Basico F3", "AH", "No se usa en este cuadro"
    EscribeMap ws, 29, "Estrategia (codigo)", "CF", "Ruteo TRAD / ALT"
    EscribeMap ws, 30, "Exposicion", "CG", "Geografia. No se usa en el cuadro"
    EscribeMap ws, 31, "Vintage", "CH", "No se usa en el cuadro"
    EscribeMap ws, 32, "Moneda", "CI", "Bloque del libro TRAD"

    ws.Range("B2:B13").HorizontalAlignment = xlLeft
    ws.Range("C5:C6").NumberFormat = "dd/mm/yyyy"
    ws.Range("C10:C11").NumberFormat = "dd/mm/yyyy"
    ws.Columns("B").ColumnWidth = 34
    ws.Columns("C").ColumnWidth = 20
    ws.Columns("D").ColumnWidth = 52
    ws.Range("C2:C13").Interior.Color = RGB(255, 242, 204)
    ws.Range("C17:C32").Interior.Color = RGB(255, 242, 204)

    CrearClasif

    Application.ScreenUpdating = scrPrev
    MsgBox "Config y Clasif listas." & vbCrLf & vbCrLf & _
           "Revisa C2 (" & ws.Range("C2").Value & "), pon la ruta en C3 y el corte en C6, " & _
           "y ejecuta ActualizarTodo.", vbInformation, "CrearConfig"
    Exit Sub

Limpio:
    Application.ScreenUpdating = scrPrev
    If Err.Number <> 0 Then
        MsgBox "CrearConfig se detuvo: " & Err.Description, vbExclamation
    End If
End Sub


Public Sub CrearClasif()
    Dim ws As Worksheet, d() As Variant, i As Long
    Dim scrPrev As Boolean

    scrPrev = Application.ScreenUpdating
    Application.ScreenUpdating = False

    Set ws = HojaOCrea(SH_CLA)
    ws.Cells.Clear

    ws.Range("A1").Value = "Codigo"
    ws.Range("B1").Value = "Nombre largo"
    ws.Range("C1").Value = "Libro"
    ws.Range("D1").Value = "Bloque"
    ws.Range("A1:D1").Font.Bold = True
    ws.Range("A1:D1").Interior.Color = RGB(217, 217, 217)

    d = Array( _
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

    For i = LBound(d) To UBound(d)
        ws.Cells(i + 2, 1).Value = d(i)(0)
        ws.Cells(i + 2, 2).Value = d(i)(1)
        ws.Cells(i + 2, 3).Value = d(i)(2)
        ws.Cells(i + 2, 4).Value = d(i)(3)
    Next i

    ws.Range("A1:D" & (UBound(d) + 2)).Borders.LineStyle = xlContinuous
    ws.Range("A1:D" & (UBound(d) + 2)).Borders.Color = RGB(191, 191, 191)
    ws.Range("A:D").ColumnWidth = 38
    ws.Columns("A").ColumnWidth = 20
    ws.Columns("C").ColumnWidth = 10
    ws.Columns("D").ColumnWidth = 22

    ws.Range("F1").Value = "El orden de las filas manda el orden de las estrategias en el CUADRO."
    ws.Range("F2").Value = "Codigo que no este en esta lista -> SIN CLASIFICAR + REVISAR en Mapa."
    ws.Range("F1:F2").Font.Color = RGB(128, 128, 128)

    Application.ScreenUpdating = scrPrev
End Sub


'===============================================================================
'  2. ACTUALIZAR TODO  (va a la red)
'===============================================================================
Public Sub ActualizarTodo()
    Dim wsCfg As Worksheet
    Dim ruta As String, hojaSrc As String
    Dim wbSrc As Workbook, wsSrc As Worksheet
    Dim v As Variant
    Dim colMap As Object, cabeceras As Object
    Dim nTrad As Long, nAlt As Long, nSin As Long, nCarg As Long
    Dim sinCodigos As Object
    Dim avisos As Object
    Dim calcPrev As XlCalculation, scrPrev As Boolean, evPrev As Boolean
    Dim filaCab As Long, filaDat As Long
    Dim libro As String
    Dim bYaAbierto As Boolean, sNotaFuente As String

    mT0 = Timer
    mPasos = ""

    On Error GoTo Falla
    scrPrev = Application.ScreenUpdating
    evPrev = Application.EnableEvents
    calcPrev = Application.Calculation
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    Paso "Validando Config"
    Set wsCfg = ThisWorkbook.Worksheets(SH_CFG)
    libro = UCase$(Trim$(CStr(wsCfg.Range("C2").Value)))
    If libro <> "TRAD" And libro <> "ALT" Then
        Err.Raise vbObjectError + 1, , "Config!C2 debe decir TRAD o ALT. Dice: '" & _
                  wsCfg.Range("C2").Value & "'."
    End If
    filaCab = CLng(Val(wsCfg.Range("C7").Value))
    filaDat = CLng(Val(wsCfg.Range("C8").Value))
    If filaCab < 1 Then filaCab = 3
    If filaDat <= filaCab Then filaDat = filaCab + 1

    ruta = Trim$(CStr(wsCfg.Range("C3").Value))
    If Len(ruta) = 0 Then
        ruta = PedirRuta()
        If Len(ruta) = 0 Then
            Application.Calculation = xlCalculationAutomatic
            Application.ScreenUpdating = scrPrev
            Application.EnableEvents = evPrev
            Application.StatusBar = False
            MsgBox "Cancelado: no se eligio ningun archivo de marcas.", vbInformation, "ActualizarTodo"
            Exit Sub
        End If
        wsCfg.Range("C3").Value = ruta
    End If
    If Dir(ruta) = "" Then
        Err.Raise vbObjectError + 2, , "No encuentro el archivo de marcas:" & vbCrLf & ruta
    End If

    hojaSrc = Trim$(CStr(wsCfg.Range("C4").Value))
    If Len(hojaSrc) = 0 Then hojaSrc = "3"

    ' Si el libro de marcas YA esta abierto en este Excel, Workbooks.Open devuelve
    ' el que ya esta (ignora ReadOnly) y cerrarlo seria cerrarle el archivo al usuario.
    ' En ese caso se lee sin abrir y sin cerrar.
    Set wbSrc = LibroYaAbierto(ruta)
    bYaAbierto = Not wbSrc Is Nothing

    If bYaAbierto Then
        Paso "La fuente ya estaba abierta: se lee sin cerrarla"
        sNotaFuente = "La fuente ya estaba abierta en tu Excel." & vbCrLf & _
                      "Se leyo la version en memoria (puede diferir del disco si tienes cambios sin guardar)" & vbCrLf & _
                      "y NO se cerro." & vbCrLf & vbCrLf
    Else
        Paso "Abriendo la fuente en solo lectura"
        Set wbSrc = Workbooks.Open(Filename:=ruta, ReadOnly:=True, UpdateLinks:=0, _
                                   AddToMru:=False)
        sNotaFuente = ""
    End If
    On Error Resume Next
    Set wsSrc = wbSrc.Worksheets(hojaSrc)
    If wsSrc Is Nothing Then Set wsSrc = wbSrc.Worksheets(CLng(Val(hojaSrc)))
    On Error GoTo Falla
    If wsSrc Is Nothing Then
        If Not bYaAbierto Then wbSrc.Close SaveChanges:=False
        Err.Raise vbObjectError + 3, , "No encuentro la hoja '" & hojaSrc & "' en la fuente."
    End If

    Paso "Leyendo la hoja a memoria"
    Set colMap = LeerMapaColumnas(wsCfg)
    v = LeerBloque(wsSrc, colMap, filaDat)
    Set cabeceras = LeerCabeceras(wsSrc, colMap, filaCab)

    If Not bYaAbierto Then wbSrc.Close SaveChanges:=False
    Set wbSrc = Nothing
    Set wsSrc = Nothing

    Paso "Construyendo la BD"
    Set sinCodigos = CreateObject("Scripting.Dictionary")
    Set avisos = CreateObject("Scripting.Dictionary")
    ConstruirBD v, colMap, libro, wsCfg, nTrad, nAlt, nSin, nCarg, sinCodigos, avisos

    Paso "Escribiendo el Mapa"
    EscribirMapa colMap, cabeceras, libro, nTrad, nAlt, nSin, nCarg, sinCodigos, avisos

    If nCarg = 0 Then
        Application.Calculation = calcPrev
        Application.ScreenUpdating = scrPrev
        Application.EnableEvents = evPrev
        Application.StatusBar = False
        MsgBox "No se cargo ninguna marca para el libro " & libro & "." & vbCrLf & _
               "Revisa la hoja Mapa: puede ser la fecha de C5 o el codigo de estrategia.", _
               vbExclamation, "ActualizarTodo"
        Exit Sub
    End If

    Paso "Armando el CUADRO"
    ArmarCuadro libro

    Application.Calculation = xlCalculationAutomatic
    Application.CalculateFullRebuild

    Application.ScreenUpdating = scrPrev
    Application.EnableEvents = evPrev
    Application.StatusBar = False

    MsgBox "v1.9 - ActualizarTodo terminado en " & Format$(Timer - mT0, "0.0") & " s." & vbCrLf & vbCrLf & _
           sNotaFuente & _
           mPasos & vbCrLf & _
           "Marcas cargadas en este libro (" & libro & "): " & Format$(nCarg, "#,##0") & vbCrLf & _
           "Sin clasificar: " & nSin & vbCrLf & vbCrLf & _
           "Revisa la hoja Mapa antes de usar el cuadro.", vbInformation, "ActualizarTodo"
    Exit Sub

Falla:
    Dim msg As String, errNum As Long
    errNum = Err.Number
    msg = Err.Description
    On Error Resume Next
    If Not wbSrc Is Nothing Then
        If Not bYaAbierto Then wbSrc.Close SaveChanges:=False
    End If
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.StatusBar = False
    Dim texto As String, codigo As String
    codigo = CodigoError(errNum) & " @ " & mDetalle
    texto = "v1.9 - ActualizarTodo se detuvo." & vbCrLf & vbCrLf & msg & vbCrLf & vbCrLf & _
            "Paso a paso:" & vbCrLf & mPasos
    GuardarLog "CODIGO: " & codigo & vbCrLf & vbCrLf & texto
    MsgBox "ESCRIBEME ESTA LINEA:" & vbCrLf & vbCrLf & _
           "   " & codigo & vbCrLf & vbCrLf & _
           String$(40, "-") & vbCrLf & msg & vbCrLf & _
           "(el detalle completo quedo en la hoja Log)", vbCritical, "ActualizarTodo"
End Sub


'===============================================================================
'  3. RECALCULAR  (no toca la red)
'===============================================================================
Public Sub Recalcular()
    Dim wsCfg As Worksheet, wsMap As Worksheet
    Dim libro As String, libroBD As String
    Dim calcPrev As XlCalculation, scrPrev As Boolean, evPrev As Boolean

    On Error GoTo Falla
    mT0 = Timer
    mPasos = ""

    scrPrev = Application.ScreenUpdating
    evPrev = Application.EnableEvents
    calcPrev = Application.Calculation
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    Set wsCfg = ThisWorkbook.Worksheets(SH_CFG)
    libro = UCase$(Trim$(CStr(wsCfg.Range("C2").Value)))
    If libro <> "TRAD" And libro <> "ALT" Then
        Err.Raise vbObjectError + 1, , "Config!C2 debe decir TRAD o ALT."
    End If

    If Not HojaExiste(SH_BD) Then
        Err.Raise vbObjectError + 4, , "No hay BD cargada. Ejecuta ActualizarTodo."
    End If
    If ThisWorkbook.Worksheets(SH_BD).Cells(2, 1).Value = "" Then
        Err.Raise vbObjectError + 4, , "La BD esta vacia. Ejecuta ActualizarTodo."
    End If

    ' guarda: C2 vs lo que realmente esta cargado en BD
    libroBD = ""
    On Error Resume Next
    Set wsMap = ThisWorkbook.Worksheets(SH_MAP)
    If Not wsMap Is Nothing Then libroBD = UCase$(Trim$(CStr(wsMap.Range("B2").Value)))
    On Error GoTo Falla
    If Len(libroBD) > 0 And libroBD <> libro Then
        Application.Calculation = calcPrev
        Application.ScreenUpdating = scrPrev
        Application.EnableEvents = evPrev
        MsgBox "Config!C2 dice " & libro & " pero la BD cargada es de " & libroBD & "." & vbCrLf & _
               "Ejecuta ActualizarTodo para recargar la BD con el filtro nuevo.", _
               vbExclamation, "Recalcular"
        Exit Sub
    End If

    Paso "Armando el CUADRO"
    ArmarCuadro libro

    Application.Calculation = xlCalculationAutomatic
    Application.CalculateFullRebuild

    Application.ScreenUpdating = scrPrev
    Application.EnableEvents = evPrev
    Application.StatusBar = False
    Exit Sub

Falla:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.StatusBar = False
    Dim cod2 As String
    cod2 = CodigoError(Err.Number) & " @ " & mDetalle
    GuardarLog "CODIGO: " & cod2 & vbCrLf & vbCrLf & Err.Description
    MsgBox "ESCRIBEME ESTA LINEA:" & vbCrLf & vbCrLf & _
           "   " & cod2 & vbCrLf & vbCrLf & _
           String$(40, "-") & vbCrLf & Err.Description, vbCritical, "Recalcular"
End Sub


'===============================================================================
'  4. LECTURA DE LA FUENTE
'===============================================================================
Private Function LeerMapaColumnas(wsCfg As Worksheet) As Object
    Dim d As Object, i As Long, campo As String, letra As String
    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = 1

    For i = 17 To 32
        campo = Trim$(CStr(wsCfg.Cells(i, 2).Value))
        letra = UCase$(Trim$(CStr(wsCfg.Cells(i, 3).Value)))
        If Len(campo) > 0 And Len(letra) > 0 Then
            d(campo) = LetraANum(letra)
        End If
    Next i

    ' campos imprescindibles
    Dim req As Variant, k As Variant
    req = Array("Nombre", "Fecha (Dia de la marca)", "Var Adj", "Codigo SBS", _
                "GAP PRF F1", "GAP PRF F2", "GAP PRF F3", "Estrategia (codigo)", "Moneda")
    For Each k In req
        If Not d.Exists(CStr(k)) Then
            Err.Raise vbObjectError + 10, , "Falta la letra de '" & k & "' en Config (C17:C32)."
        End If
    Next k

    Set LeerMapaColumnas = d
End Function


Private Function LeerCabeceras(wsSrc As Worksheet, colMap As Object, filaCab As Long) As Object
    Dim d As Object, k As Variant
    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = 1
    For Each k In colMap.Keys
        d(CStr(k)) = Trim$(CStr(wsSrc.Cells(filaCab, colMap(k)).Text))
    Next k
    Set LeerCabeceras = d
End Function


Private Function LeerBloque(wsSrc As Worksheet, colMap As Object, filaDat As Long) As Variant
    Dim maxCol As Long, ultFila As Long, k As Variant
    Dim colFecha As Long

    maxCol = 0
    For Each k In colMap.Keys
        If colMap(k) > maxCol Then maxCol = colMap(k)
    Next k

    colFecha = colMap("Fecha (Dia de la marca)")
    ultFila = wsSrc.Cells(wsSrc.Rows.Count, colFecha).End(xlUp).Row
    If ultFila < filaDat Then
        Err.Raise vbObjectError + 11, , "La hoja fuente no tiene datos debajo de la fila " & filaDat & "."
    End If

    LeerBloque = wsSrc.Range(wsSrc.Cells(filaDat, 1), wsSrc.Cells(ultFila, maxCol)).Value2
End Function


'===============================================================================
'  5. CONSTRUCCION DE LA BD
'===============================================================================
Private Sub ConstruirBD(v As Variant, colMap As Object, libro As String, _
                        wsCfg As Worksheet, ByRef nTrad As Long, ByRef nAlt As Long, _
                        ByRef nSin As Long, ByRef nCarg As Long, _
                        sinCodigos As Object, avisos As Object)

    Dim ws As Worksheet
    Dim i As Long, n As Long, fila As Long
    Dim out() As Variant
    Dim cla As Object, claOrden As Object
    Dim dFecha As Double, dIni As Double
    Dim sNom As String, sCod As String, sEstr As String, sMon As String
    Dim sVint As String, sExp As String
    Dim sLibroFila As String, sBloque As String, sSub As String
    Dim vAdj As Variant, vSbs As Variant, dVar As Double
    Dim bVarOK As Boolean, bUsoSbs As Boolean
    Dim flag As String
    Dim dupKey As Object
    Dim k As String
    Dim nDup As Long, nSbs As Long, nSinVar As Long, nFueraRango As Long
    Dim minVar As Double, maxVar As Double, bPrimero As Boolean
    Dim vFec As Variant
    Dim nFechaMala As Long, nSinNombre As Long, nPrevias As Long

    Set cla = LeerClasif(claOrden)
    Set dupKey = CreateObject("Scripting.Dictionary")
    dupKey.CompareMode = 1

    dIni = 0
    If IsDate(wsCfg.Range("C5").Value) Then dIni = CDbl(CDate(wsCfg.Range("C5").Value))

    n = UBound(v, 1)
    ReDim out(1 To n, 1 To BD_COLS)
    fila = 0
    bPrimero = True
    minVar = 0: maxVar = 0

    For i = 1 To n
        ' --- fecha (eje) ---
        vFec = v(i, colMap("Fecha (Dia de la marca)"))
        If Not EsNumero(vFec) Then
            ' fecha como texto o basura: no se salta en silencio, se cuenta
            If Len(Trim$(TextoSeguro(vFec))) > 0 Then nFechaMala = nFechaMala + 1
            GoTo Siguiente
        End If
        dFecha = CDbl(vFec)
        If dFecha <= 0 Then
            nFechaMala = nFechaMala + 1
            GoTo Siguiente
        End If
        If dFecha < dIni Then
            nPrevias = nPrevias + 1
            GoTo Siguiente
        End If

        sNom = Trim$(CStr(TextoSeguro(v(i, colMap("Nombre")))))
        If Len(sNom) = 0 Then
            nSinNombre = nSinNombre + 1
            GoTo Siguiente
        End If

        flag = ""

        sCod = Trim$(CStr(TextoSeguro(v(i, colMap("Codigo SBS")))))
        sEstr = Trim$(CStr(TextoSeguro(v(i, colMap("Estrategia (codigo)")))))
        sMon = Trim$(CStr(TextoSeguro(v(i, colMap("Moneda")))))
        sVint = Trim$(CStr(TextoSeguro(v(i, colMap("Vintage")))))
        sExp = Trim$(CStr(TextoSeguro(v(i, colMap("Exposicion")))))

        ' --- ruteo por estrategia ---
        If Len(sEstr) = 0 Then
            sLibroFila = "SIN CLASIFICAR"
        ElseIf cla.Exists(sEstr) Then
            sLibroFila = cla(sEstr)(0)
        Else
            sLibroFila = "SIN CLASIFICAR"
        End If

        If sLibroFila = "TRAD" Then
            nTrad = nTrad + 1
        ElseIf sLibroFila = "ALT" Then
            nAlt = nAlt + 1
        Else
            nSin = nSin + 1
            k = IIf(Len(sEstr) = 0, "(vacio)", sEstr)
            If Not sinCodigos.Exists(k) Then sinCodigos(k) = 0
            sinCodigos(k) = sinCodigos(k) + 1
        End If

        If sLibroFila <> libro Then GoTo Siguiente

        ' --- bloque y subgrupo ---
        If libro = "TRAD" Then
            sBloque = IIf(Len(sMon) = 0, "SIN MONEDA", UCase$(sMon))
            sSub = ""
            If Len(sMon) = 0 Then flag = AgregaFlag(flag, "SIN MONEDA")
        Else
            sBloque = cla(sEstr)(2)
            sSub = cla(sEstr)(1)
            If Len(sBloque) = 0 Then sBloque = "SIN BLOQUE"
        End If

        ' --- retorno ---
        vAdj = v(i, colMap("Var Adj"))
        bVarOK = False: bUsoSbs = False
        If EsNumero(vAdj) Then
            dVar = CDbl(vAdj)
            bVarOK = True
        Else
            If colMap.Exists("Var SBS") Then
                vSbs = v(i, colMap("Var SBS"))
                If EsNumero(vSbs) Then
                    dVar = CDbl(vSbs)
                    bVarOK = True
                    bUsoSbs = True
                    nSbs = nSbs + 1
                End If
            End If
        End If

        If Not bVarOK Then
            dVar = 0
            flag = AgregaFlag(flag, "REVISAR: sin Var Adj ni Var SBS")
            nSinVar = nSinVar + 1
        Else
            If bPrimero Then
                minVar = dVar: maxVar = dVar: bPrimero = False
            Else
                If dVar < minVar Then minVar = dVar
                If dVar > maxVar Then maxVar = dVar
            End If
            If Abs(dVar) > 1 Then
                flag = AgregaFlag(flag, "REVISAR: |Var Adj| > 100%")
                nFueraRango = nFueraRango + 1
            End If
        End If
        If bUsoSbs Then flag = AgregaFlag(flag, "Var Adj vacio, se uso Var SBS")

        ' --- duplicados: mismo codigo + nombre + dia ---
        k = sCod & "|" & UCase$(sNom) & "|" & CStr(CLng(dFecha))
        If dupKey.Exists(k) Then
            flag = AgregaFlag(flag, "REVISAR: duplicado")
            nDup = nDup + 1
        Else
            dupKey(k) = 1
        End If

        ' --- volcado ---
        fila = fila + 1
        out(fila, 1) = dFecha
        out(fila, 2) = Year(CDate(dFecha))
        out(fila, 3) = Month(CDate(dFecha))
        out(fila, 4) = sNom
        out(fila, 5) = sCod
        out(fila, 6) = sBloque
        out(fila, 7) = sSub
        out(fila, 8) = sEstr
        out(fila, 9) = sExp
        out(fila, 10) = sMon
        out(fila, 11) = sVint
        out(fila, 12) = dVar

        If bVarOK And (1 + dVar) > 0 Then
            out(fila, 13) = Log(1 + dVar)
        Else
            out(fila, 13) = CVErr(xlErrNA)
            If bVarOK Then
                out(fila, 17) = AgregaFlag(flag, "REVISAR: 1+Var Adj <= 0, no hay LN")
                flag = CStr(out(fila, 17))
            End If
        End If

        out(fila, 14) = NumOVacio(v(i, colMap("GAP PRF F1")))
        out(fila, 15) = NumOVacio(v(i, colMap("GAP PRF F2")))
        out(fila, 16) = NumOVacio(v(i, colMap("GAP PRF F3")))
        out(fila, 17) = flag

Siguiente:
    Next i

    nCarg = fila

    ' --- escribir la hoja BD de un solo volcado ---
    Set ws = HojaOCrea(SH_BD)
    ws.Cells.Clear

    Dim cab As Variant
    cab = Array("Dia", "Anio", "Mes", "Nombre", "Codigo SBS", "Bloque", "Subgrupo", _
                "Estrategia cod", "Exposicion", "Moneda", "Vintage", "Var Adj", _
                "LN(1+Var Adj)", "Peso F1", "Peso F2", "Peso F3", "Flag")
    ws.Range(ws.Cells(1, 1), ws.Cells(1, BD_COLS)).Value = cab
    ws.Range(ws.Cells(1, 1), ws.Cells(1, BD_COLS)).Font.Bold = True
    ws.Range(ws.Cells(1, 1), ws.Cells(1, BD_COLS)).Interior.Color = RGB(217, 217, 217)

    If fila > 0 Then
        ws.Range(ws.Cells(2, 1), ws.Cells(fila + 1, BD_COLS)).Value = out
        ws.Range(ws.Cells(2, 1), ws.Cells(fila + 1, 1)).NumberFormat = "dd/mm/yyyy"
        ws.Range(ws.Cells(2, 12), ws.Cells(fila + 1, 12)).NumberFormat = "0.00%"
        ws.Range(ws.Cells(2, 13), ws.Cells(fila + 1, 13)).NumberFormat = "0.000000"
        ws.Range(ws.Cells(2, 14), ws.Cells(fila + 1, 16)).NumberFormat = "0.00%"
    End If

    ws.Range(ws.Cells(1, 1), ws.Cells(1, BD_COLS)).EntireColumn.AutoFit
    ws.Columns(4).ColumnWidth = 34

    CongelarPaneles ws, 1, 0

    ' --- nombres definidos ---
    Dim ultBD As Long
    ultBD = IIf(fila > 0, fila + 1, 2)
    DefinirNombre "bdDia", SH_BD, "$A$2:$A$" & ultBD
    DefinirNombre "bdNom", SH_BD, "$D$2:$D$" & ultBD
    DefinirNombre "bdBloque", SH_BD, "$F$2:$F$" & ultBD
    DefinirNombre "bdSub", SH_BD, "$G$2:$G$" & ultBD
    DefinirNombre "bdVar", SH_BD, "$L$2:$L$" & ultBD
    DefinirNombre "bdLn", SH_BD, "$M$2:$M$" & ultBD
    DefinirNombre "bdPeso1", SH_BD, "$N$2:$N$" & ultBD
    DefinirNombre "bdPeso2", SH_BD, "$O$2:$O$" & ultBD
    DefinirNombre "bdPeso3", SH_BD, "$P$2:$P$" & ultBD

    ' --- avisos para el Mapa ---
    avisos("Filas saltadas: fecha no numerica o invalida") = nFechaMala
    avisos("Filas saltadas: sin nombre de fondo") = nSinNombre
    avisos("Filas saltadas: anteriores al inicio del historico") = nPrevias
    avisos("Duplicados (mismo codigo+nombre+dia)") = nDup
    avisos("Var Adj vacio, se uso Var SBS") = nSbs
    avisos("Sin Var Adj ni Var SBS") = nSinVar
    avisos("|Var Adj| > 100%") = nFueraRango
    avisos("Var Adj minimo") = minVar
    avisos("Var Adj maximo") = maxVar
End Sub


Private Function LeerClasif(ByRef orden As Object) As Object
    Dim ws As Worksheet, d As Object, i As Long, ult As Long
    Dim cod As String

    If Not HojaExiste(SH_CLA) Then CrearClasif
    Set ws = ThisWorkbook.Worksheets(SH_CLA)

    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = 1
    Set orden = CreateObject("Scripting.Dictionary")
    orden.CompareMode = 1

    ult = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For i = 2 To ult
        cod = Trim$(CStr(ws.Cells(i, 1).Value))
        If Len(cod) > 0 Then
            d(cod) = Array(UCase$(Trim$(CStr(ws.Cells(i, 3).Value))), _
                           Trim$(CStr(ws.Cells(i, 2).Value)), _
                           UCase$(Trim$(CStr(ws.Cells(i, 4).Value))))
            If Not orden.Exists(Trim$(CStr(ws.Cells(i, 2).Value))) Then
                orden(Trim$(CStr(ws.Cells(i, 2).Value))) = i
            End If
        End If
    Next i

    Set LeerClasif = d
End Function


'===============================================================================
'  6. HOJA MAPA
'===============================================================================
Private Sub EscribirMapa(colMap As Object, cabeceras As Object, libro As String, _
                         nTrad As Long, nAlt As Long, nSin As Long, nCarg As Long, _
                         sinCodigos As Object, avisos As Object)
    Dim ws As Worksheet, r As Long, k As Variant
    Dim estado As String, cab As String, campo As String

    Set ws = HojaOCrea(SH_MAP)
    ws.Cells.Clear

    ws.Range("A1").Value = "MAPA DE COLUMNAS Y CONTROLES"
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Font.Size = 13
    ws.Range("A2").Value = "Libro cargado en BD:"
    ws.Range("B2").Value = libro
    ws.Range("B2").Font.Bold = True
    ws.Range("A3").Value = "Corrida:"
    ws.Range("B3").Value = Now
    ws.Range("B3").NumberFormat = "dd/mm/yyyy hh:mm"

    r = 5
    ws.Cells(r, 1).Value = "Campo"
    ws.Cells(r, 2).Value = "Columna"
    ws.Cells(r, 3).Value = "Cabecera encontrada en la fuente"
    ws.Cells(r, 4).Value = "Estado"
    ws.Range(ws.Cells(r, 1), ws.Cells(r, 4)).Font.Bold = True
    ws.Range(ws.Cells(r, 1), ws.Cells(r, 4)).Interior.Color = RGB(217, 217, 217)

    For Each k In colMap.Keys
        r = r + 1
        campo = CStr(k)
        cab = ""
        If cabeceras.Exists(campo) Then cab = CStr(cabeceras(campo))

        estado = "OK"
        If Len(cab) = 0 Then estado = "REVISAR: cabecera vacia"
        If campo = "Fecha (Dia de la marca)" Then
            If InStr(1, cab, "Fecha", vbTextCompare) = 0 Then estado = "REVISAR: no dice Fecha"
        ElseIf campo = "Var Adj" Then
            If InStr(1, cab, "Var", vbTextCompare) = 0 Then estado = "REVISAR: no dice Var"
        End If

        ws.Cells(r, 1).Value = campo
        ws.Cells(r, 2).Value = NumALetra(CLng(colMap(k)))
        ws.Cells(r, 3).Value = cab
        ws.Cells(r, 4).Value = estado
        If Left$(estado, 7) = "REVISAR" Then
            ws.Range(ws.Cells(r, 1), ws.Cells(r, 4)).Interior.Color = RGB(255, 199, 206)
            ws.Range(ws.Cells(r, 1), ws.Cells(r, 4)).Font.Color = RGB(156, 0, 6)
        End If
    Next k

    r = r + 2
    ws.Cells(r, 1).Value = "RUTEO"
    ws.Cells(r, 1).Font.Bold = True
    r = r + 1: ws.Cells(r, 1).Value = "Marcas TRAD (desde el inicio del historico)": ws.Cells(r, 2).Value = nTrad
    r = r + 1: ws.Cells(r, 1).Value = "Marcas ALT": ws.Cells(r, 2).Value = nAlt
    r = r + 1: ws.Cells(r, 1).Value = "Sin clasificar": ws.Cells(r, 2).Value = nSin
    If nSin > 0 Then
        ws.Range(ws.Cells(r, 1), ws.Cells(r, 2)).Interior.Color = RGB(255, 199, 206)
        ws.Range(ws.Cells(r, 1), ws.Cells(r, 2)).Font.Color = RGB(156, 0, 6)
    End If
    r = r + 1: ws.Cells(r, 1).Value = "Cargadas en este libro (" & libro & ")": ws.Cells(r, 2).Value = nCarg
    ws.Cells(r, 2).Font.Bold = True

    If sinCodigos.Count > 0 Then
        r = r + 2
        ws.Cells(r, 1).Value = "CODIGOS DE ESTRATEGIA SIN CLASIFICAR"
        ws.Cells(r, 1).Font.Bold = True
        For Each k In sinCodigos.Keys
            r = r + 1
            ws.Cells(r, 1).Value = CStr(k)
            ws.Cells(r, 2).Value = sinCodigos(k)
            ws.Cells(r, 3).Value = "Agregalo a la hoja Clasif y vuelve a correr ActualizarTodo."
        Next k
    End If

    r = r + 2
    ws.Cells(r, 1).Value = "CONTROLES DE LA BD"
    ws.Cells(r, 1).Font.Bold = True
    For Each k In avisos.Keys
        r = r + 1
        ws.Cells(r, 1).Value = CStr(k)
        If CStr(k) = "Var Adj minimo" Or CStr(k) = "Var Adj maximo" Then
            ws.Cells(r, 2).Value = avisos(k)
            ws.Cells(r, 2).NumberFormat = "0.00%"
            ws.Cells(r, 3).Value = "Si esto sale en miles, la fuente trae el retorno en puntos, no en decimal."
        Else
            ws.Cells(r, 2).Value = avisos(k)
            If avisos(k) > 0 _
               And CStr(k) <> "Var Adj vacio, se uso Var SBS" _
               And CStr(k) <> "Filas saltadas: anteriores al inicio del historico" Then
                ws.Range(ws.Cells(r, 1), ws.Cells(r, 2)).Interior.Color = RGB(255, 235, 156)
            End If
            If Left$(CStr(k), 14) = "Filas saltadas" Then
                ws.Cells(r, 3).Value = "Estas filas NO entraron a la BD."
            Else
                ws.Cells(r, 3).Value = "Filtra la columna Flag de la BD para verlas."
            End If
        End If
    Next k

    r = r + 2
    ws.Cells(r, 1).Value = "CONTROLES DE CIERRE"
    ws.Cells(r, 1).Font.Bold = True
    r = r + 1: ws.Cells(r, 1).Value = "1. Ningun REVISAR arriba. Fecha en B, Var Adj en F."
    r = r + 1: ws.Cells(r, 1).Value = "2. Var Adj minimo/maximo dentro de un rango creible (decimal, no puntos)."
    r = r + 1: ws.Cells(r, 1).Value = "3. TRAD + ALT + sin clasificar = total de marcas desde el inicio del historico."
    r = r + 1: ws.Cells(r, 1).Value = "4. Una marca conocida contra su retorno esperado en la BD."

    ws.Columns("A").ColumnWidth = 48
    ws.Columns("B").ColumnWidth = 16
    ws.Columns("C").ColumnWidth = 62
    ws.Columns("D").ColumnWidth = 30
    CongelarPaneles ws, 5, 0
End Sub


'===============================================================================
'  7. CUADRO
'===============================================================================
Private Sub ArmarCuadro(libro As String)
    Dim ws As Worksheet, wsBD As Worksheet, wsCfg As Worksheet
    Dim ultBD As Long
    Dim dCorte As Double, dBaseMin As Double
    Dim bloques As Object, claOrden As Object, cla As Object
    Dim nExcl As Long

    Marca "ArmarCuadro: abriendo Config y BD"
    Set wsCfg = ThisWorkbook.Worksheets(SH_CFG)
    Set wsBD = ThisWorkbook.Worksheets(SH_BD)
    ultBD = wsBD.Cells(wsBD.Rows.Count, 1).End(xlUp).Row
    If ultBD < 2 Then Err.Raise vbObjectError + 20, , "La BD esta vacia."

    ' --- corte ---
    Marca "ArmarCuadro: determinando la fecha de corte"
    If IsDate(wsCfg.Range("C6").Value) Then
        dCorte = CDbl(CDate(wsCfg.Range("C6").Value))
    Else
        dCorte = Application.WorksheetFunction.Max(wsBD.Range("A2:A" & ultBD))
    End If
    If dCorte <= 0 Then Err.Raise vbObjectError + 21, , "No pude determinar la fecha de corte."

    Marca "ArmarCuadro: calculando las bases de las ventanas"
    dBaseMin = BaseMasAntigua(wsCfg, dCorte)

    Marca "ArmarCuadro: leyendo la hoja Clasif"
    Set cla = LeerClasif(claOrden)

    Marca "ArmarCuadro: armando la estructura de filas"
    Set bloques = ArmarEstructura(wsBD, ultBD, libro, dCorte, dBaseMin, claOrden, nExcl)

    ' --- hoja ---
    Marca "ArmarCuadro: creando y limpiando la hoja CUADRO"
    Set ws = HojaOCrea(SH_CUA)
    On Error Resume Next
    ws.Cells.FormatConditions.Delete
    On Error GoTo 0
    ws.Cells.Clear
    ws.Cells.Interior.Pattern = xlNone

    ws.Range("B2").Value = IIf(libro = "TRAD", "Retornos: Fondos Tradicionales", _
                                               "Retornos: Fondos Alternativos")
    ws.Range("B2").Font.Bold = True
    ws.Range("B2").Font.Size = 15
    ws.Range("B2").Font.Color = RGB(192, 0, 0)

    ws.Range("D2").Value = dCorte
    ws.Range("D2").NumberFormat = "dd/mm/yyyy"
    ws.Range("D2").Font.Bold = True
    ws.Range("E2").Value = "<- fecha de corte (Config!C6, vacio = ultima marca)"
    ws.Range("E2").Font.Color = RGB(128, 128, 128)

    ws.Range("B3").Value = "Fondos sin posicion excluidos: " & nExcl & _
                           "   |   Retorno compuesto sobre Var Adj   |   Filas de grupo: solo division"
    ws.Range("B3").Font.Color = RGB(128, 128, 128)
    ws.Range("B3").Font.Size = 8

    Marca "ArmarCuadro: encabezado de ventanas"
    EscribirEncabezadoVentanas ws

    ' --- filas ---
    Marca "ArmarCuadro: escribiendo las filas"
    EscribirFilas ws, bloques, libro

    Marca "ArmarCuadro: aplicando formato"
    FormatearCuadro ws, libro
End Sub


Private Sub EscribirEncabezadoVentanas(ws As Worksheet)
    Dim nom As Variant, j As Long, c As Long
    nom = Array("WTD", "5D", "MTD", "MayoTD", "YTD", "FY")

    ws.Cells(CU_FILA_FIN, CU_COL_NOM).Value = "fin"
    ws.Cells(CU_FILA_BASE, CU_COL_NOM).Value = "base"
    ws.Cells(CU_FILA_FIN, CU_COL_NOM).HorizontalAlignment = xlRight
    ws.Cells(CU_FILA_BASE, CU_COL_NOM).HorizontalAlignment = xlRight
    ws.Range(ws.Cells(CU_FILA_FIN, CU_COL_NOM), ws.Cells(CU_FILA_BASE, CU_COL_NOM)).Font.Color = RGB(128, 128, 128)

    For j = 0 To N_VENT - 1
        c = CU_COL_W1 + j
        ws.Cells(CU_FILA_FIN, c).Formula = "=$D$2"
        ws.Cells(CU_FILA_CAB, c).Value = nom(j)
    Next j

    ' bases, formulas visibles contra Config
    ws.Cells(CU_FILA_BASE, CU_COL_W1 + 0).Formula = "=$D$2-WEEKDAY($D$2,3)-1"
    ws.Cells(CU_FILA_BASE, CU_COL_W1 + 1).Formula = "=$D$2-" & SH_CFG & "!$C$9"
    ws.Cells(CU_FILA_BASE, CU_COL_W1 + 2).Formula = "=EOMONTH($D$2,-1)"
    ws.Cells(CU_FILA_BASE, CU_COL_W1 + 3).Formula = "=" & SH_CFG & "!$C$11-1"
    ws.Cells(CU_FILA_BASE, CU_COL_W1 + 4).Formula = "=DATE(YEAR($D$2),1,1)-1"
    ws.Cells(CU_FILA_BASE, CU_COL_W1 + 5).Formula = "=" & SH_CFG & "!$C$10"

    ws.Cells(CU_FILA_CAB, CU_COL_P1 + 0).Value = "F1"
    ws.Cells(CU_FILA_CAB, CU_COL_P1 + 1).Value = "F2"
    ws.Cells(CU_FILA_CAB, CU_COL_P1 + 2).Value = "F3"
End Sub


Private Sub EscribirFilas(ws As Worksheet, bloques As Object, libro As String)
    Dim r As Long, j As Long, c As Long
    Dim kB As Variant, kS As Variant, kF As Variant
    Dim subs As Object, fondos As Object
    Dim filasBloque As Object, filasSub As Object, filasFondo As Object
    Dim filaB As Long, filaS As Long
    Dim lista As String
    Dim tipos As Object

    Set tipos = CreateObject("Scripting.Dictionary")
    Set filasBloque = CreateObject("Scripting.Dictionary")

    r = CU_FILA_DAT - 1

    For Each kB In bloques.Keys
        r = r + 1
        filaB = r
        Marca "Filas: bloque '" & TextoSeguro(kB) & "' en la fila " & r
        tipos(r) = "BLQ"
        ws.Cells(r, CU_COL_BLQ).Value = TextoSeguro(kB)
        ws.Cells(r, CU_COL_NOM).Value = TextoSeguro(kB)
        Set filasSub = CreateObject("Scripting.Dictionary")

        Set subs = bloques(kB)
        For Each kS In subs.Keys
            Set filasFondo = CreateObject("Scripting.Dictionary")
            If TextoSeguro(kS) <> "" Then
                r = r + 1
                filaS = r
                Marca "Filas: subgrupo '" & TextoSeguro(kS) & "' en la fila " & r
                tipos(r) = "SUB"
                ws.Cells(r, CU_COL_BLQ).Value = TextoSeguro(kB)
                ws.Cells(r, CU_COL_SUB).Value = TextoSeguro(kS)
                ws.Cells(r, CU_COL_NOM).Value = TextoSeguro(kS)
            Else
                filaS = 0
            End If

            Set fondos = subs(kS)
            For Each kF In fondos.Keys
                r = r + 1
                Marca "Filas: fondo '" & TextoSeguro(kF) & "' en la fila " & r
                tipos(r) = "FND"
                ws.Cells(r, CU_COL_BLQ).Value = TextoSeguro(kB)
                ws.Cells(r, CU_COL_SUB).Value = TextoSeguro(kS)
                ws.Cells(r, CU_COL_NOM).Value = TextoSeguro(kF)
                ws.Cells(r, CU_COL_NOM).IndentLevel = 2

                ' pesos
                Marca "Filas: formulas de peso, fila " & r
                For j = 0 To 2
                    c = CU_COL_P1 + j
                    ws.Cells(r, c).Formula = FormulaPeso(r, j + 1)
                Next j
                ' retornos
                Marca "Filas: formulas de retorno, fila " & r
                For j = 0 To N_VENT - 1
                    c = CU_COL_W1 + j
                    ws.Cells(r, c).Formula = FormulaRetorno(r, c)
                Next j
                filasFondo(r) = 1
            Next kF

            If filaS > 0 Then
                Marca "Filas: suma del subgrupo en la fila " & filaS
                lista = RangoDeFilas(filasFondo)
                If Len(lista) > 0 Then
                    For j = 0 To 2
                        c = CU_COL_P1 + j
                        ws.Cells(filaS, c).Formula = FormulaSumaGrupo(c, filasFondo)
                    Next j
                End If
                filasSub(filaS) = 1
            Else
                ' sin nivel de estrategia: los fondos cuelgan del bloque
                For Each kF In filasFondo.Keys
                    filasSub(kF) = 1
                Next kF
            End If
        Next kS

        Marca "Filas: suma del bloque en la fila " & filaB
        For j = 0 To 2
            c = CU_COL_P1 + j
            ws.Cells(filaB, c).Formula = FormulaSumaGrupo(c, filasSub)
        Next j
        filasBloque(filaB) = 1
        ws.Cells(filaB, CU_COL_SUB).Value = ""
    Next kB

    ' --- Total ---
    Marca "Filas: fila Total"
    r = r + 1
    tipos(r) = "TOT"
    ws.Cells(r, CU_COL_NOM).Value = "Total"
    For j = 0 To 2
        c = CU_COL_P1 + j
        ws.Cells(r, c).Formula = FormulaSumaGrupo(c, filasBloque)
    Next j

    Marca "Filas: guardando los tipos de fila"
    GuardarTipos ws, tipos, r
End Sub


Private Function FormulaPeso(r As Long, cual As Long) As String
    Dim rng As String, mx As String
    rng = "bdPeso" & cual
    mx = "MAXIFS(bdDia,bdNom,$B" & r & ",bdDia,""<=""&$D$2)"
    FormulaPeso = "=IFERROR(IF(AVERAGEIFS(" & rng & ",bdNom,$B" & r & ",bdDia," & mx & ")=0,""""," & _
                  "AVERAGEIFS(" & rng & ",bdNom,$B" & r & ",bdDia," & mx & ")),"""")"
End Function


Private Function FormulaRetorno(r As Long, c As Long) As String
    Dim col As String
    col = NumALetra(c)
    ' si la base es posterior o igual al fin, la ventana no existe todavia -> #N/D
    FormulaRetorno = "=IF(" & col & "$" & CU_FILA_BASE & ">=" & col & "$" & CU_FILA_FIN & ",NA()," & _
                     "EXP(SUMIFS(bdLn,bdNom,$B" & r & _
                     ",bdDia,"">""&" & col & "$" & CU_FILA_BASE & _
                     ",bdDia,""<=""&" & col & "$" & CU_FILA_FIN & "))-1)"
End Function


Private Function FormulaSumaGrupo(c As Long, filas As Object) As String
    ' Suma explicita de las filas hijas (nunca un MAXIFS a nivel de grupo: eso
    ' solo tomaria los fondos que marcaron el ultimo dia).
    ' SUM admite hasta 255 argumentos, asi que se parte en tramos por si acaso.
    Const MAX_ARG As Long = 200
    Dim k As Variant, col As String
    Dim s As String, total As String, n As Long

    col = NumALetra(c)
    s = "": total = "": n = 0

    For Each k In filas.Keys
        If Len(s) > 0 Then s = s & ","
        s = s & col & CStr(k)
        n = n + 1
        If n Mod MAX_ARG = 0 Then
            If Len(total) > 0 Then total = total & "+"
            total = total & "SUM(" & s & ")"
            s = ""
        End If
    Next k

    If Len(s) > 0 Then
        If Len(total) > 0 Then total = total & "+"
        total = total & "SUM(" & s & ")"
    End If

    If Len(total) = 0 Then
        FormulaSumaGrupo = ""
    Else
        FormulaSumaGrupo = "=IF(" & total & "=0,""""," & total & ")"
    End If
End Function


Private Function RangoDeFilas(filas As Object) As String
    Dim k As Variant, s As String
    For Each k In filas.Keys
        If Len(s) > 0 Then s = s & ","
        s = s & CStr(k)
    Next k
    RangoDeFilas = s
End Function


Private Sub GuardarTipos(ws As Worksheet, tipos As Object, ultFila As Long)
    Dim r As Long
    ' la marca del tipo de fila viaja en la columna oculta A, al lado del bloque
    For r = CU_FILA_DAT To ultFila
        If tipos.Exists(r) Then
            ws.Cells(r, CU_COL_SUB + 1).Value = tipos(r)   ' col N, oculta
        End If
    Next r
End Sub


Private Function ArmarEstructura(wsBD As Worksheet, ultBD As Long, libro As String, _
                                 dCorte As Double, dBaseMin As Double, _
                                 claOrden As Object, ByRef nExcl As Long) As Object
    Dim v As Variant, i As Long
    Dim bloques As Object, subs As Object, fondos As Object
    Dim conPos As Object, vistos As Object
    Dim sB As String, sS As String, sN As String
    Dim dFecha As Double
    Dim tienePeso As Boolean
    Dim arrBloques As Variant, arrSubs As Variant, nombres As Variant
    Dim ib As Long, isb As Long, inm As Long

    Marca "Estructura: leyendo la BD a memoria"
    v = wsBD.Range(wsBD.Cells(2, 1), wsBD.Cells(ultBD, BD_COLS)).Value2

    Set conPos = CreateObject("Scripting.Dictionary"): conPos.CompareMode = 1
    Set vistos = CreateObject("Scripting.Dictionary"): vistos.CompareMode = 1

    ' 1) que fondos tienen posicion dentro de la ventana mas larga
    Marca "Estructura: filtro de posicion"
    For i = 1 To UBound(v, 1)
        If EsNumero(v(i, 1)) Then
            dFecha = CDbl(v(i, 1))
            sN = TextoSeguro(v(i, 4))
            If Len(sN) > 0 Then
                If Not vistos.Exists(sN) Then vistos(sN) = TextoSeguro(v(i, 6)) & "|" & TextoSeguro(v(i, 7))
                If dFecha > dBaseMin And dFecha <= dCorte Then
                    tienePeso = (Abs(NumCero(v(i, 14))) > 0.0000001) Or _
                                (Abs(NumCero(v(i, 15))) > 0.0000001) Or _
                                (Abs(NumCero(v(i, 16))) > 0.0000001)
                    If tienePeso Then conPos(sN) = 1
                End If
            End If
        End If
    Next i

    ' 2) estructura ordenada
    Set bloques = CreateObject("Scripting.Dictionary"): bloques.CompareMode = 1

    Marca "Estructura: ordenando los bloques"
    arrBloques = OrdenBloques(v, libro)

    For ib = LBound(arrBloques) To UBound(arrBloques)
        sB = TextoSeguro(arrBloques(ib))
        Marca "Estructura: bloque '" & sB & "'"
        Set subs = CreateObject("Scripting.Dictionary"): subs.CompareMode = 1

        arrSubs = OrdenSubs(v, sB, libro, claOrden)

        For isb = LBound(arrSubs) To UBound(arrSubs)
            sS = TextoSeguro(arrSubs(isb))
            Marca "Estructura: bloque '" & sB & "' subgrupo '" & sS & "'"
            Set fondos = CreateObject("Scripting.Dictionary"): fondos.CompareMode = 1

            nombres = NombresDe(v, sB, sS)
            If IsArray(nombres) Then
                For inm = LBound(nombres) To UBound(nombres)
                    sN = TextoSeguro(nombres(inm))
                    If conPos.Exists(sN) Then
                        fondos(sN) = 1
                    Else
                        nExcl = nExcl + 1
                    End If
                Next inm
            End If

            ' OJO: guardar un objeto en un Dictionary EXIGE Set. Sin Set, VBA intenta
            ' asignar la propiedad por defecto (Item) y tira "No coinciden los tipos".
            Marca "Estructura: guardando subgrupo '" & sS & "' (" & fondos.Count & " fondos)"
            If fondos.Count > 0 Then Set subs(sS) = fondos
        Next isb

        Marca "Estructura: guardando bloque '" & sB & "' (" & subs.Count & " subgrupos)"
        If subs.Count > 0 Then Set bloques(sB) = subs
    Next ib

    If bloques.Count = 0 Then
        Err.Raise vbObjectError + 22, , "Ningun fondo paso el filtro de posicion. " & _
                  "Revisa la fecha de corte (Config!C6) y los pesos de las cols I/S/AC."
    End If

    Set ArmarEstructura = bloques
End Function


Private Function OrdenBloques(v As Variant, libro As String) As Variant
    Dim d As Object, i As Long, s As String
    Dim res() As String, n As Long, k As Variant

    Set d = CreateObject("Scripting.Dictionary"): d.CompareMode = 1
    For i = 1 To UBound(v, 1)
        s = Trim$(TextoSeguro(v(i, 6)))
        If Len(s) > 0 Then d(s) = 1
    Next i

    If libro = "ALT" Then
        ReDim res(0 To d.Count + 1)
        n = -1
        If d.Exists("INTERNACIONALES") Then n = n + 1: res(n) = "INTERNACIONALES"
        If d.Exists("LOCALES") Then n = n + 1: res(n) = "LOCALES"
        For Each k In d.Keys
            If CStr(k) <> "INTERNACIONALES" And CStr(k) <> "LOCALES" Then
                n = n + 1: res(n) = CStr(k)
            End If
        Next k
    Else
        ReDim res(0 To d.Count + 1)
        n = -1
        If d.Exists("PEN") Then n = n + 1: res(n) = "PEN"
        If d.Exists("USD") Then n = n + 1: res(n) = "USD"
        For Each k In d.Keys
            If CStr(k) <> "PEN" And CStr(k) <> "USD" Then
                n = n + 1: res(n) = CStr(k)
            End If
        Next k
    End If

    If n < 0 Then
        OrdenBloques = Array()
    Else
        ReDim Preserve res(0 To n)
        OrdenBloques = res
    End If
End Function


Private Function OrdenSubs(v As Variant, bloque As String, libro As String, _
                           claOrden As Object) As Variant
    Dim d As Object, i As Long, s As String
    Dim res() As String, n As Long, k As Variant
    Dim mejor As Long, mejorNom As String, pos As Long

    Set d = CreateObject("Scripting.Dictionary"): d.CompareMode = 1
    For i = 1 To UBound(v, 1)
        If StrComp(Trim$(TextoSeguro(v(i, 6))), bloque, vbTextCompare) = 0 Then
            s = Trim$(TextoSeguro(v(i, 7)))
            d(s) = 1
        End If
    Next i

    If d.Count = 0 Then
        OrdenSubs = Array("")
        Exit Function
    End If

    ' orden = orden de filas de Clasif; lo no listado va al final
    ReDim res(0 To d.Count - 1)
    n = -1
    Do While d.Count > 0
        mejor = 999999
        mejorNom = ""
        For Each k In d.Keys
            If claOrden.Exists(CStr(k)) Then
                pos = CLng(claOrden(CStr(k)))
            Else
                pos = 900000
            End If
            If pos < mejor Then
                mejor = pos
                mejorNom = CStr(k)
            End If
        Next k
        If mejorNom = "" Then
            For Each k In d.Keys
                mejorNom = CStr(k)
                Exit For
            Next k
        End If
        n = n + 1
        res(n) = mejorNom
        d.Remove mejorNom
    Loop

    OrdenSubs = res
End Function


Private Function NombresDe(v As Variant, bloque As String, subg As String) As Variant
    Dim d As Object, i As Long, s As String
    Dim arr() As String, n As Long, k As Variant
    Dim j As Long, tmp As String

    Set d = CreateObject("Scripting.Dictionary"): d.CompareMode = 1
    For i = 1 To UBound(v, 1)
        If StrComp(Trim$(TextoSeguro(v(i, 6))), bloque, vbTextCompare) = 0 Then
            If StrComp(Trim$(TextoSeguro(v(i, 7))), subg, vbTextCompare) = 0 Then
                s = Trim$(TextoSeguro(v(i, 4)))
                If Len(s) > 0 Then d(s) = 1
            End If
        End If
    Next i

    If d.Count = 0 Then
        NombresDe = Array()
        Exit Function
    End If

    ReDim arr(0 To d.Count - 1)
    n = -1
    For Each k In d.Keys
        n = n + 1
        arr(n) = CStr(k)
    Next k

    ' orden alfabetico
    For i = LBound(arr) To UBound(arr) - 1
        For j = i + 1 To UBound(arr)
            If StrComp(arr(i), arr(j), vbTextCompare) > 0 Then
                tmp = arr(i): arr(i) = arr(j): arr(j) = tmp
            End If
        Next j
    Next i

    NombresDe = arr
End Function


Private Function BaseMasAntigua(wsCfg As Worksheet, dCorte As Double) As Double
    Dim b(1 To 6) As Double, i As Long, mn As Double
    Dim dc As Date
    dc = CDate(CLng(Int(dCorte)))

    b(1) = dCorte - (Weekday(dc, vbMonday) - 1) - 1          ' domingo anterior
    b(2) = dCorte - NumCero(wsCfg.Range("C9").Value)
    b(3) = CDbl(DateSerial(Year(dc), Month(dc), 0))          ' fin del mes anterior
    If IsDate(wsCfg.Range("C11").Value) Then
        b(4) = CDbl(CDate(wsCfg.Range("C11").Value)) - 1
    Else
        b(4) = dCorte
    End If
    b(5) = CDbl(DateSerial(Year(dc), 1, 1)) - 1
    If IsDate(wsCfg.Range("C10").Value) Then
        b(6) = CDbl(CDate(wsCfg.Range("C10").Value))
    Else
        b(6) = CDbl(DateSerial(Year(dc) - 1, 10, 31))
    End If

    mn = b(1)
    For i = 2 To 6
        If b(i) < mn And b(i) > 0 Then mn = b(i)
    Next i
    BaseMasAntigua = mn
End Function


'===============================================================================
'  8. FORMATO DEL CUADRO
'===============================================================================
Private Sub FormatearCuadro(ws As Worksheet, libro As String)
    Dim ultFila As Long, r As Long, j As Long, c As Long
    Dim tipo As String
    Dim rngDat As Range

    ultFila = ws.Cells(ws.Rows.Count, CU_COL_NOM).End(xlUp).Row
    If ultFila < CU_FILA_DAT Then Exit Sub

    ' --- anchos ---
    Marca "Formato: anchos de columna"
    ws.Columns(CU_COL_BLQ).ColumnWidth = 18
    ws.Columns(CU_COL_NOM).ColumnWidth = 40
    ws.Range(NumALetra(CU_COL_P1) & ":" & NumALetra(CU_COL_P1 + 2)).ColumnWidth = 8
    ws.Columns(CU_COL_SEP).ColumnWidth = 1.6
    ws.Range(NumALetra(CU_COL_W1) & ":" & NumALetra(CU_COL_W1 + N_VENT - 1)).ColumnWidth = 10
    ws.Columns(CU_COL_SUB).ColumnWidth = 30
    ws.Columns(CU_COL_SUB + 1).ColumnWidth = 6

    ' --- formatos numericos ---
    Marca "Formato: formatos numericos"
    ws.Range(ws.Cells(CU_FILA_FIN, CU_COL_W1), ws.Cells(CU_FILA_BASE, CU_COL_W1 + N_VENT - 1)).NumberFormat = "dd/mm/yy"
    ws.Range(ws.Cells(CU_FILA_FIN, CU_COL_W1), ws.Cells(CU_FILA_FIN, CU_COL_W1 + N_VENT - 1)).Font.Color = RGB(89, 89, 89)
    ws.Range(ws.Cells(CU_FILA_BASE, CU_COL_W1), ws.Cells(CU_FILA_BASE, CU_COL_W1 + N_VENT - 1)).Font.Color = RGB(166, 166, 166)
    ws.Range(ws.Cells(CU_FILA_FIN, CU_COL_W1), ws.Cells(CU_FILA_BASE, CU_COL_W1 + N_VENT - 1)).Font.Size = 8
    ws.Range(ws.Cells(CU_FILA_FIN, CU_COL_W1), ws.Cells(CU_FILA_BASE, CU_COL_W1 + N_VENT - 1)).HorizontalAlignment = xlCenter

    ws.Range(ws.Cells(CU_FILA_DAT, CU_COL_P1), ws.Cells(ultFila, CU_COL_P1 + 2)).NumberFormat = "0.0%"
    ws.Range(ws.Cells(CU_FILA_DAT, CU_COL_W1), ws.Cells(ultFila, CU_COL_W1 + N_VENT - 1)).NumberFormat = "0.0%"
    ws.Range(ws.Cells(CU_FILA_DAT, CU_COL_P1), ws.Cells(ultFila, CU_COL_W1 + N_VENT - 1)).HorizontalAlignment = xlCenter

    ' --- cabecera ---
    Marca "Formato: cabecera"
    With ws.Range(ws.Cells(CU_FILA_CAB, CU_COL_W1), ws.Cells(CU_FILA_CAB, CU_COL_W1 + N_VENT - 1))
        .Interior.Color = RGB(192, 0, 0)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
    End With
    With ws.Range(ws.Cells(CU_FILA_CAB, CU_COL_P1), ws.Cells(CU_FILA_CAB, CU_COL_P1 + 2))
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).Color = RGB(128, 128, 128)
    End With

    ' --- bordes suaves en el cuerpo ---
    Marca "Formato: bordes"
    Set rngDat = ws.Range(ws.Cells(CU_FILA_DAT, CU_COL_NOM), ws.Cells(ultFila, CU_COL_W1 + N_VENT - 1))
    With rngDat.Borders
        .LineStyle = xlContinuous
        .Color = RGB(217, 217, 217)
        .Weight = xlThin
    End With
    ws.Range(ws.Cells(CU_FILA_DAT, CU_COL_SEP), ws.Cells(ultFila, CU_COL_SEP)).Borders.LineStyle = xlNone

    ' --- pintar por tipo de fila ---
    For r = CU_FILA_DAT To ultFila
        Marca "Formato: pintando la fila " & r
        tipo = TextoSeguro(ws.Cells(r, CU_COL_SUB + 1).Value)
        Select Case tipo
            Case "BLQ"
                With ws.Range(ws.Cells(r, CU_COL_NOM), ws.Cells(r, CU_COL_P1 + 2))
                    .Interior.Color = RGB(128, 128, 128)
                    .Font.Color = RGB(255, 255, 255)
                    .Font.Bold = True
                End With
            Case "SUB"
                With ws.Range(ws.Cells(r, CU_COL_NOM), ws.Cells(r, CU_COL_P1 + 2))
                    .Interior.Color = RGB(217, 217, 217)
                    .Font.Bold = True
                End With
                ws.Cells(r, CU_COL_NOM).IndentLevel = 1
            Case "TOT"
                With ws.Range(ws.Cells(r, CU_COL_NOM), ws.Cells(r, CU_COL_P1 + 2))
                    .Font.Bold = True
                End With
                With ws.Range(ws.Cells(r, CU_COL_NOM), ws.Cells(r, CU_COL_W1 + N_VENT - 1)).Borders(xlEdgeTop)
                    .LineStyle = xlContinuous
                    .Color = RGB(64, 64, 64)
                    .Weight = xlMedium
                End With
        End Select
    Next r

    ' --- escala de color, una por columna, solo filas de fondo ---
    For j = 0 To N_VENT - 1
        c = CU_COL_W1 + j
        Marca "Formato: escala de color en la columna " & NumALetra(c)
        AplicarEscala ws.Range(ws.Cells(CU_FILA_DAT, c), ws.Cells(ultFila, c))
    Next j

    ' --- ocultar auxiliares ---
    Marca "Formato: ocultando columnas auxiliares"
    ws.Columns(CU_COL_BLQ).Hidden = True
    ws.Columns(CU_COL_SUB).Hidden = True
    ws.Columns(CU_COL_SUB + 1).Hidden = True

    Marca "Formato: congelando paneles"
    CongelarPaneles ws, CU_FILA_CAB, CU_COL_NOM
    Marca "Formato: terminado"
End Sub


Private Sub AplicarEscala(rng As Range)
    Dim fc As ColorScale
    On Error Resume Next
    rng.FormatConditions.Delete
    On Error GoTo 0

    Set fc = rng.FormatConditions.AddColorScale(ColorScaleType:=3)
    With fc.ColorScaleCriteria(1)
        .Type = xlConditionValueLowestValue
        .FormatColor.Color = RGB(230, 60, 60)
    End With
    With fc.ColorScaleCriteria(2)
        .Type = xlConditionValueNumber
        .Value = 0
        .FormatColor.Color = RGB(255, 255, 255)
    End With
    With fc.ColorScaleCriteria(3)
        .Type = xlConditionValueHighestValue
        .FormatColor.Color = RGB(99, 190, 123)
    End With
End Sub


'===============================================================================
'  9. ARCHIVO
'===============================================================================
Public Sub ArchivarAhora()
    Dim ws As Worksheet, wbNew As Workbook
    Dim carpeta As String, base As String, ruta As String
    Dim fmt As String, libro As String
    Dim dCorte As Variant, sFecha As String
    Dim n As Long
    Dim scrPrev As Boolean, evPrev As Boolean

    On Error GoTo Falla
    scrPrev = Application.ScreenUpdating
    evPrev = Application.EnableEvents
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    If Not HojaExiste(SH_CUA) Then
        Err.Raise vbObjectError + 30, , "No hay CUADRO que archivar. Ejecuta ActualizarTodo."
    End If
    Set ws = ThisWorkbook.Worksheets(SH_CUA)

    libro = UCase$(Trim$(CStr(ThisWorkbook.Worksheets(SH_CFG).Range("C2").Value)))
    fmt = UCase$(Trim$(CStr(ThisWorkbook.Worksheets(SH_CFG).Range("C13").Value)))
    If Len(fmt) = 0 Then fmt = "XLSX"
    If fmt = "NO" Then
        Application.ScreenUpdating = scrPrev
        Application.EnableEvents = evPrev
        MsgBox "Config!C13 dice NO. No archive nada.", vbInformation, "ArchivarAhora"
        Exit Sub
    End If

    carpeta = Trim$(CStr(ThisWorkbook.Worksheets(SH_CFG).Range("C12").Value))
    If Len(carpeta) = 0 Then carpeta = ThisWorkbook.Path & Application.PathSeparator & "Resumenes"
    If Right$(carpeta, 1) = Application.PathSeparator Then carpeta = Left$(carpeta, Len(carpeta) - 1)
    If Dir(carpeta, vbDirectory) = "" Then MkDir carpeta

    dCorte = ws.Range("D2").Value
    If IsDate(dCorte) Then
        sFecha = Format$(CDate(dCorte), "yyyymmdd")
    Else
        sFecha = Format$(Date, "yyyymmdd")
    End If

    base = IIf(libro = "TRAD", "Retornos_Tradicionales_", "Retornos_Alternativos_") & sFecha

    ' append-only: nunca se pisa
    ruta = carpeta & Application.PathSeparator & base
    n = 1
    Do While Dir(ruta & ".xlsx") <> "" Or Dir(ruta & ".pdf") <> ""
        n = n + 1
        ruta = carpeta & Application.PathSeparator & base & "_v" & n
    Loop

    ws.Copy
    Set wbNew = ActiveWorkbook
    With wbNew.Worksheets(1)
        .UsedRange.Value = .UsedRange.Value
        .Columns(CU_COL_BLQ).Hidden = True
        .Columns(CU_COL_SUB).Hidden = True
        .Columns(CU_COL_SUB + 1).Hidden = True
    End With

    If fmt = "XLSX" Or fmt = "AMBOS" Then
        wbNew.SaveAs Filename:=ruta & ".xlsx", FileFormat:=xlOpenXMLWorkbook
    End If
    If fmt = "PDF" Or fmt = "AMBOS" Then
        wbNew.Worksheets(1).PageSetup.Orientation = xlLandscape
        wbNew.Worksheets(1).PageSetup.Zoom = False
        wbNew.Worksheets(1).PageSetup.FitToPagesWide = 1
        wbNew.Worksheets(1).PageSetup.FitToPagesTall = False
        wbNew.ExportAsFixedFormat Type:=xlTypePDF, Filename:=ruta & ".pdf", _
                                  Quality:=xlQualityStandard, OpenAfterPublish:=False
    End If

    wbNew.Close SaveChanges:=False
    Set wbNew = Nothing

    Application.ScreenUpdating = scrPrev
    Application.EnableEvents = evPrev
    MsgBox "Archivado en:" & vbCrLf & ruta, vbInformation, "ArchivarAhora"
    Exit Sub

Falla:
    On Error Resume Next
    If Not wbNew Is Nothing Then wbNew.Close SaveChanges:=False
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    MsgBox "ArchivarAhora se detuvo." & vbCrLf & vbCrLf & Err.Description, vbCritical, "ArchivarAhora"
End Sub


'===============================================================================
' 10. UTILIDADES
'===============================================================================
Private Function CodigoError(n As Long) As String
    ' Los errores propios se lanzan como vbObjectError + k; se muestran como Ekk
    ' para que la linea a reportar quede corta.
    If n < 0 Then
        CodigoError = "E" & Format$(n - vbObjectError, "00")
    Else
        CodigoError = CStr(n)
    End If
End Function


Public Sub GuardarLog(texto As String)
    ' Deja el detalle del error en la hoja Log y en un .txt junto al libro,
    ' para poder copiarlo y pegarlo en vez de fotografiar la pantalla.
    Dim ws As Worksheet, ruta As String, f As Integer, i As Long
    Dim partes As Variant

    On Error Resume Next

    Set ws = HojaOCrea("Log")
    If Not ws Is Nothing Then
        ws.Cells.Clear
        ws.Range("A1").Value = "Ultimo error - " & Format$(Now, "dd/mm/yyyy hh:mm:ss")
        ws.Range("A1").Font.Bold = True
        partes = Split(texto, vbCrLf)
        For i = LBound(partes) To UBound(partes)
            ws.Cells(i + 3, 1).Value = "'" & partes(i)
        Next i
        ws.Columns("A").ColumnWidth = 110
    End If

    If Len(ThisWorkbook.Path) > 0 Then
        ruta = ThisWorkbook.Path & Application.PathSeparator & "Log_Retornos.txt"
        f = FreeFile
        Open ruta For Output As #f
        Print #f, "=== " & Format$(Now, "dd/mm/yyyy hh:mm:ss") & " ==="
        Print #f, texto
        Close #f
    End If

    On Error GoTo 0
End Sub


Private Sub Marca(texto As String)
    ' marcador fino: se nombra en el mensaje de error para ubicar la linea exacta
    mDetalle = texto
End Sub


Private Sub Paso(texto As String)
    mPasos = mPasos & "  - " & texto & " (" & Format$(Timer - mT0, "0.0") & " s)" & vbCrLf
    Application.StatusBar = "Retornos: " & texto & "..."
    DoEvents
End Sub


Private Function HojaOCrea(nombre As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(nombre)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = nombre
    End If
    Set HojaOCrea = ws
End Function


Private Function HojaExiste(nombre As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(nombre)
    On Error GoTo 0
    HojaExiste = Not ws Is Nothing
End Function


Private Sub EscribeCfg(ws As Worksheet, fila As Long, etiqueta As String, _
                       valor As Variant, nota As String)
    ws.Cells(fila, 2).Value = etiqueta
    If Not IsEmpty(valor) Then ws.Cells(fila, 3).Value = valor
    ws.Cells(fila, 4).Value = nota
    ws.Cells(fila, 4).Font.Color = RGB(128, 128, 128)
End Sub


Private Sub EscribeMap(ws As Worksheet, fila As Long, campo As String, _
                       letra As String, nota As String)
    ws.Cells(fila, 2).Value = campo
    ws.Cells(fila, 3).Value = letra
    ws.Cells(fila, 4).Value = nota
    ws.Cells(fila, 4).Font.Color = RGB(128, 128, 128)
End Sub


Private Sub DefinirNombre(nombre As String, hoja As String, refLocal As String)
    Dim ws As Worksheet, rng As Range, nm As Name
    Dim i As Long, eDesc As String

    Set ws = ThisWorkbook.Worksheets(hoja)

    ' Borra CUALQUIER nombre que se llame igual, sea de ambito de libro o de hoja.
    ' Un libro copiado de otra herramienta hereda nombres de ambito de hoja que no
    ' aparecen como ThisWorkbook.Names(nombre) y hacen chocar el Add.
    ' Excel rechaza un nombre definido que se pueda leer como direccion de celda.
    ' Ojo: "bdP1" ES una direccion valida (columna BDP, fila 1) y por eso fallaba.
    If PareceReferencia(nombre) Then
        Err.Raise vbObjectError + 51, , _
            "El nombre '" & nombre & "' se puede leer como una direccion de celda, " & _
            "asi que Excel no lo acepta como nombre definido. Hay que usar uno que no " & _
            "termine en digitos detras de letras de columna validas (A..XFD)."
    End If

    For i = ThisWorkbook.Names.Count To 1 Step -1
        Set nm = ThisWorkbook.Names(i)
        If StrComp(NombreCorto(nm.Name), nombre, vbTextCompare) = 0 Then
            On Error Resume Next
            nm.Delete
            On Error GoTo 0
        End If
    Next i

    Set rng = ws.Range(refLocal)

    ' RefersTo con un objeto Range, no con texto: evita el problema de A1 vs R1C1
    ' y el del separador de listas segun el idioma de Excel.
    On Error Resume Next
    ThisWorkbook.Names.Add Name:=nombre, RefersTo:=rng
    eDesc = ""
    If Err.Number <> 0 Then eDesc = Err.Description
    On Error GoTo 0

    If Len(eDesc) > 0 Then
        Err.Raise vbObjectError + 50, , _
            "No pude crear el nombre definido '" & nombre & "' para " & _
            hoja & "!" & refLocal & "." & vbCrLf & vbCrLf & _
            "Excel dijo: " & eDesc & vbCrLf & vbCrLf & _
            "Casi siempre es un nombre heredado de otro libro. Ejecuta la macro " & _
            "LimpiarNombres y vuelve a correr ActualizarTodo."
    End If
End Sub


Private Function PareceReferencia(ByVal nombre As String) As Boolean
    ' True si el texto es una direccion de celda estilo A1 valida (p. ej. "bdP1" = BDP1)
    ' o estilo R1C1 (p. ej. "R1C1"). Esos no sirven como nombres definidos.
    Dim i As Long, nLet As Long, nDig As Long, ch As String
    Dim letras As String, digitos As String, col As Long

    nombre = Trim$(nombre)
    If Len(nombre) = 0 Then Exit Function

    For i = 1 To Len(nombre)
        ch = UCase$(Mid$(nombre, i, 1))
        If ch >= "A" And ch <= "Z" Then
            If nDig > 0 Then Exit Function      ' letras despues de digitos: no es A1
            letras = letras & ch
            nLet = nLet + 1
        ElseIf ch >= "0" And ch <= "9" Then
            digitos = digitos & ch
            nDig = nDig + 1
        Else
            Exit Function                       ' cualquier otro caracter: no es A1
        End If
    Next i

    If nLet = 0 Or nDig = 0 Then Exit Function
    If nLet > 3 Or nDig > 7 Then Exit Function

    col = 0
    For i = 1 To Len(letras)
        col = col * 26 + (Asc(Mid$(letras, i, 1)) - 64)
    Next i

    If col >= 1 And col <= 16384 Then
        If Val(digitos) >= 1 And Val(digitos) <= 1048576 Then PareceReferencia = True
    End If
End Function


Private Function NombreCorto(s As String) As String
    ' "BD!bdDia" (ambito de hoja) -> "bdDia"
    Dim p As Long
    p = InStrRev(s, "!")
    If p > 0 Then
        NombreCorto = Mid$(s, p + 1)
    Else
        NombreCorto = s
    End If
End Function


Public Sub LimpiarNombres()
    ' Borra los nombres definidos heredados de otro libro. Solo toca los que
    ' empiezan con "bd" y los que quedaron apuntando a #REF!.
    Dim nm As Name, i As Long
    Dim nBd As Long, nRef As Long, det As String

    For i = ThisWorkbook.Names.Count To 1 Step -1
        Set nm = ThisWorkbook.Names(i)
        If LCase$(Left$(NombreCorto(nm.Name), 2)) = "bd" Then
            det = det & "  " & nm.Name & "   ->   " & nm.RefersTo & vbCrLf
            On Error Resume Next
            nm.Delete
            On Error GoTo 0
            nBd = nBd + 1
        ElseIf InStr(1, nm.RefersTo, "#REF!", vbTextCompare) > 0 Then
            On Error Resume Next
            nm.Delete
            On Error GoTo 0
            nRef = nRef + 1
        End If
    Next i

    MsgBox "Nombres borrados:" & vbCrLf & vbCrLf & _
           "  que empezaban con 'bd': " & nBd & vbCrLf & _
           "  rotos (#REF!): " & nRef & vbCrLf & vbCrLf & _
           IIf(Len(det) > 0, "Detalle:" & vbCrLf & det & vbCrLf, "") & _
           "Quedan " & ThisWorkbook.Names.Count & " nombres en el libro." & vbCrLf & _
           "Ahora ejecuta ActualizarTodo.", vbInformation, "LimpiarNombres"
End Sub


Private Sub CongelarPaneles(ws As Worksheet, filas As Long, cols As Long)
    ' sin .Select: se activa la hoja y se usa SplitRow / SplitColumn
    Dim wsPrev As Object
    On Error Resume Next
    ThisWorkbook.Activate
    Set wsPrev = ActiveSheet
    ws.Activate
    ActiveWindow.FreezePanes = False
    ActiveWindow.SplitRow = filas
    ActiveWindow.SplitColumn = cols
    ActiveWindow.FreezePanes = True
    If Not wsPrev Is Nothing Then wsPrev.Activate
    On Error GoTo 0
End Sub


Private Function LibroYaAbierto(ruta As String) As Workbook
    ' Devuelve el Workbook si ese archivo ya esta abierto en esta instancia de Excel.
    Dim wb As Workbook, nom As String, p As Long
    p = InStrRev(ruta, Application.PathSeparator)
    If p = 0 Then p = InStrRev(ruta, "/")
    nom = Mid$(ruta, p + 1)
    For Each wb In Application.Workbooks
        If StrComp(wb.Name, nom, vbTextCompare) = 0 Then
            Set LibroYaAbierto = wb
            Exit Function
        End If
    Next wb
End Function


Private Function PedirRuta() As String
    Dim fd As Object, s As String
    Set fd = Application.FileDialog(3)   ' msoFileDialogFilePicker
    fd.Title = "Elige el libro de marcas"
    fd.Filters.Clear
    fd.Filters.Add "Excel", "*.xlsx; *.xlsm; *.xlsb"
    fd.AllowMultiSelect = False
    If fd.Show = -1 Then s = fd.SelectedItems(1)
    PedirRuta = s
End Function


Private Function LetraANum(ByVal letra As String) As Long
    Dim i As Long, n As Long, ch As Long
    letra = UCase$(Trim$(letra))
    For i = 1 To Len(letra)
        ch = Asc(Mid$(letra, i, 1)) - 64
        If ch < 1 Or ch > 26 Then
            Err.Raise vbObjectError + 40, , "Letra de columna invalida: '" & letra & "'."
        End If
        n = n * 26 + ch
    Next i
    LetraANum = n
End Function


Private Function NumALetra(ByVal n As Long) As String
    Dim s As String, r As Long
    Do While n > 0
        r = ((n - 1) Mod 26)
        s = Chr$(65 + r) & s
        n = (n - r - 1) \ 26
    Loop
    NumALetra = s
End Function


Private Function TextoSeguro(v As Variant) As String
    If IsError(v) Then
        TextoSeguro = ""
    ElseIf IsEmpty(v) Then
        TextoSeguro = ""
    ElseIf IsNull(v) Then
        TextoSeguro = ""
    Else
        TextoSeguro = CStr(v)
    End If
End Function


Private Function EsNumero(v As Variant) As Boolean
    If IsError(v) Then
        EsNumero = False
    ElseIf IsEmpty(v) Then
        EsNumero = False
    ElseIf VarType(v) = vbString Then
        If Len(Trim$(CStr(v))) = 0 Then
            EsNumero = False
        Else
            EsNumero = IsNumeric(v)
        End If
    Else
        EsNumero = IsNumeric(v)
    End If
End Function


Private Function NumOVacio(v As Variant) As Variant
    If EsNumero(v) Then
        NumOVacio = CDbl(v)
    Else
        NumOVacio = Empty
    End If
End Function


Private Function NumCero(v As Variant) As Double
    If EsNumero(v) Then
        NumCero = CDbl(v)
    Else
        NumCero = 0
    End If
End Function


Private Function AgregaFlag(actual As String, nuevo As String) As String
    If Len(actual) = 0 Then
        AgregaFlag = nuevo
    Else
        AgregaFlag = actual & " | " & nuevo
    End If
End Function
