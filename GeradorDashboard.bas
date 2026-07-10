Attribute VB_Name = "Module4"
Option Explicit

' --- MACRO PARA ABRIR O FORMULÁRIO ---
Sub MostrarPainel()
    frmGerador.Show
End Sub

' --- MACRO PRINCIPAL (MOTOR DO DASHBOARD) ---
Sub GerarRelatorioMaster_Pro(diasFiltro As Long, incluirSlide1 As Boolean, incluirSlide2 As Boolean)
    Dim ppApp As Object, ppPres As Object, ppSlide As Object
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim lo As ListObject
    Dim dictTipo As Object, dictMeses As Object, dictEstado As Object, dictBandeira As Object
    Dim i As Long, dataHoje As Date: dataHoje = Date
    Dim dataLimite As Date
    
    ' 1. Localizar Tabela
    On Error Resume Next
    Set lo = ws.ListObjects("tbProjetos")
    On Error GoTo 0
    
    If lo Is Nothing Then
        MsgBox "A tabela 'tbProjetos' não foi encontrada!", vbCritical
        Exit Sub
    End If
    
    ' 2. Inicializar Dicionários
    Set dictTipo = CreateObject("Scripting.Dictionary")
    Set dictMeses = CreateObject("Scripting.Dictionary")
    Set dictEstado = CreateObject("Scripting.Dictionary")
    Set dictBandeira = CreateObject("Scripting.Dictionary")
    
    ' Define o limite de data (0 = Pipeline Completo)
    dataLimite = IIf(diasFiltro = 0, DateAdd("yyyy", 10, dataHoje), dataHoje + diasFiltro)

    ' 3. Processamento de Dados
    Dim totalGeral As Long, emAndamento As Long, concluidos As Long
    totalGeral = lo.ListRows.Count
    
    For i = 1 To totalGeral
        Dim dVal As Variant: dVal = lo.ListRows(i).Range.Cells(1, lo.ListColumns("INAUGURAÇÃO").Index).Value
        Dim tVal As String: tVal = lo.ListRows(i).Range.Cells(1, lo.ListColumns("PROJETO").Index).Value
        Dim eVal As String: eVal = lo.ListRows(i).Range.Cells(1, lo.ListColumns("ESTADO").Index).Value
        Dim bVal As String: bVal = lo.ListRows(i).Range.Cells(1, lo.ListColumns("BANDEIRA").Index).Value
        
        If IsDate(dVal) Then
            If CDate(dVal) < dataHoje Then
                concluidos = concluidos + 1
            Else
                emAndamento = emAndamento + 1
                ' Pipeline Mensal dentro do filtro e cronológico
                If CDate(dVal) <= dataLimite Then
                    Dim chaveM As String: chaveM = Format(dVal, "yyyymm") & "|" & UCase(Format(dVal, "mmm/yyyy"))
                    dictMeses(chaveM) = dictMeses(chaveM) + 1
                End If
            End If
        End If
        
        If tVal <> "" Then dictTipo(tVal) = dictTipo(tVal) + 1
        If eVal <> "" Then dictEstado(eVal) = dictEstado(eVal) + 1
        If bVal <> "" Then dictBandeira(bVal) = dictBandeira(bVal) + 1
    Next i

    ' 4. Criação do PowerPoint
    Set ppApp = CreateObject("PowerPoint.Application")
    Set ppPres = ppApp.Presentations.Add

    ' Gerar Slide 1 (Sumário)
    If incluirSlide1 Then
        Set ppSlide = ppPres.Slides.Add(ppPres.Slides.Count + 1, 12) ' Layout em Branco
        DesenharCabecalho ppSlide, "EXECUTIVE REPORT | STATUS DE PORTFÓLIO", diasFiltro
        CriarCard ppSlide, 30, 95, 150, 110, "PORTFÓLIO ATIVO", totalGeral, RGB(15, 35, 75)
        CriarCard ppSlide, 195, 95, 150, 110, "EM ANDAMENTO", emAndamento, RGB(48, 127, 226)
        CriarCard ppSlide, 360, 95, 150, 110, "CONCLUÍDOS", concluidos, RGB(46, 204, 113)
        DesenharBarras ppSlide, dictTipo, 30, 240, "COMPOSIÇÃO POR TIPO", totalGeral
        DesenharTabelaPipeline ppSlide, dictMeses, 380, 240, "PIPELINE DE ENTREGAS"
    End If

    ' Gerar Slide 2 (Geográfico)
    If incluirSlide2 Then
        Set ppSlide = ppPres.Slides.Add(ppPres.Slides.Count + 1, 12)
        DesenharCabecalho ppSlide, "ANÁLISE GEOGRÁFICA E POR BANDEIRA", diasFiltro
        DesenharBarras ppSlide, dictEstado, 30, 110, "DISTRIBUIÇÃO POR ESTADO", totalGeral
        DesenharBarras ppSlide, dictBandeira, 380, 110, "VOLUME POR BANDEIRA", totalGeral
    End If

    ppApp.Visible = True
End Sub

' --- FUNÇÕES DE DESIGN (SEM BORDAS PRETAS) ---

Sub DesenharCabecalho(sld As Object, tit As String, dias As Long)
    Dim txtF As String: txtF = IIf(dias = 0, "Pipeline Completo", dias & " dias")
    Dim tb As Object
    Set tb = sld.Shapes.AddTextbox(1, 30, 20, 600, 40)
    tb.TextFrame.TextRange.Text = tit
    tb.TextFrame.TextRange.Font.Name = "Arial": tb.TextFrame.TextRange.Font.Bold = True
    tb.TextFrame.TextRange.Font.Size = 22: tb.TextFrame.TextRange.Font.Color.RGB = RGB(15, 35, 75)
    
    Set tb = sld.Shapes.AddTextbox(1, 30, 50, 400, 20)
    tb.TextFrame.TextRange.Text = "Filtro: " & txtF & " | Atualizado em: " & Now
    tb.TextFrame.TextRange.Font.Size = 9: tb.TextFrame.TextRange.Font.Color.RGB = RGB(130, 130, 130)
    
    ' Linha divisória suave
    sld.Shapes.AddConnector(1, 30, 80, 700, 80).Line.ForeColor.RGB = RGB(235, 235, 235)
End Sub

Sub CriarCard(sld As Object, x As Single, y As Single, w As Single, h As Single, tit As String, val As Long, cor As Long)
    Dim bg As Object: Set bg = sld.Shapes.AddShape(1, x, y, w, h)
    bg.Fill.ForeColor.RGB = RGB(255, 255, 255)
    bg.Line.Visible = False ' Remove borda preta
    bg.Shadow.Type = 21: bg.Shadow.Blur = 10: bg.Shadow.Transparency = 0.8: bg.Shadow.Visible = True
    
    Dim topo As Object: Set topo = sld.Shapes.AddShape(1, x, y, w, 4)
    topo.Fill.ForeColor.RGB = cor: topo.Line.Visible = False
    
    Dim tbV As Object: Set tbV = sld.Shapes.AddTextbox(1, x, y + 15, w, 60)
    tbV.TextFrame.TextRange.Text = val: tbV.TextFrame.TextRange.Font.Bold = True: tbV.TextFrame.TextRange.Font.Size = 42
    tbV.TextFrame.TextRange.Font.Color.RGB = cor: tbV.TextFrame.TextRange.ParagraphFormat.Alignment = 2

    Dim tbT As Object: Set tbT = sld.Shapes.AddTextbox(1, x, y + 75, w, 20)
    tbT.TextFrame.TextRange.Text = tit: tbT.TextFrame.TextRange.Font.Bold = True: tbT.TextFrame.TextRange.Font.Size = 9
    tbT.TextFrame.TextRange.Font.Color.RGB = RGB(100, 100, 100): tbT.TextFrame.TextRange.ParagraphFormat.Alignment = 2
End Sub

Sub DesenharBarras(sld As Object, dict As Object, x As Single, y As Single, tit As String, total As Long)
    Dim tbT As Object: Set tbT = sld.Shapes.AddTextbox(1, x, y, 300, 20)
    tbT.TextFrame.TextRange.Text = tit: tbT.TextFrame.TextRange.Font.Bold = True: tbT.TextFrame.TextRange.Font.Size = 12
    
    Dim yOff As Single: yOff = y + 35: Dim k As Variant, i As Integer: i = 0
    For Each k In dict.Keys
        If i > 8 Then Exit For
        Dim pct As Double: pct = dict(k) / total
        Dim tbi As Object: Set tbi = sld.Shapes.AddTextbox(1, x, yOff, 200, 20)
        tbi.TextFrame.TextRange.Text = k & " (" & dict(k) & ")": tbi.TextFrame.TextRange.Font.Size = 8: tbi.TextFrame.TextRange.Font.Bold = True
        
        ' Barra de fundo (Suave)
        Dim b1 As Object: Set b1 = sld.Shapes.AddShape(1, x, yOff + 14, 280, 8)
        b1.Fill.ForeColor.RGB = RGB(245, 245, 245): b1.Line.Visible = False
        
        ' Barra progresso
        Dim b2 As Object: Set b2 = sld.Shapes.AddShape(1, x, yOff + 14, 280 * pct, 8)
        b2.Fill.ForeColor.RGB = RGB(48, 127, 226): b2.Line.Visible = False
        
        yOff = yOff + 35: i = i + 1
    Next k
End Sub

Sub DesenharTabelaPipeline(sld As Object, dict As Object, x As Single, y As Single, tit As String)
    Dim tbTit As Object: Set tbTit = sld.Shapes.AddTextbox(1, x, y, 300, 20)
    tbTit.TextFrame.TextRange.Text = tit: tbTit.TextFrame.TextRange.Font.Bold = True: tbTit.TextFrame.TextRange.Font.Size = 12
    
    ' Cabeçalho Tabela
    Dim h As Object: Set h = sld.Shapes.AddShape(1, x, y + 30, 320, 25)
    h.Fill.ForeColor.RGB = RGB(25, 70, 110): h.Line.Visible = False
    
    Call TextoTabela(sld, "MÊS / PERÍODO", x + 10, y + 33, 150, True, RGB(255, 255, 255))
    Call TextoTabela(sld, "ENTREGAS", x + 170, y + 33, 100, True, RGB(255, 255, 255))

    ' Ordenação e desenho dinâmico das linhas
    If Not dict Is Nothing Then
        If dict.Count > 0 Then
            Dim arr As Variant: arr = dict.Keys
            Call SortArray(arr)
            
            Dim r As Long, yRow As Single: yRow = y + 55
            For r = LBound(arr) To UBound(arr)
                ' Linha Zebra (Sem borda)
                Dim rb As Object: Set rb = sld.Shapes.AddShape(1, x, yRow, 320, 22)
                rb.Line.Visible = False
                If r Mod 2 = 0 Then rb.Fill.ForeColor.RGB = RGB(248, 248, 248) Else rb.Fill.ForeColor.RGB = RGB(235, 240, 245)
                
                Dim nomeMes As String: nomeMes = Split(arr(r), "|")(1)
                Call TextoTabela(sld, nomeMes, x + 10, yRow + 3, 150, False, RGB(50, 50, 50))
                Call TextoTabela(sld, CStr(dict(arr(r))), x + 170, yRow + 3, 100, True, RGB(25, 70, 110))
                
                yRow = yRow + 23
                If r > 12 Then Exit For ' Limite para não sair do slide
            Next r
        End If
    End If
End Sub

' --- FUNÇÕES AUXILIARES (DEFINIDAS APENAS UMA VEZ) ---

Sub TextoTabela(s As Object, t As String, x As Single, y As Single, w As Single, b As Boolean, c As Long)
    Dim tb As Object: Set tb = s.Shapes.AddTextbox(1, x, y, w, 20)
    With tb.TextFrame.TextRange
        .Text = t
        .Font.Name = "Arial": .Font.Size = 10
        .Font.Bold = b: .Font.Color.RGB = c
    End With
End Sub

Sub SortArray(vArray As Variant)
    Dim i As Long, j As Long, temp As Variant
    If UBound(vArray) > LBound(vArray) Then
        For i = LBound(vArray) To UBound(vArray) - 1
            For j = i + 1 To UBound(vArray)
                If vArray(i) > vArray(j) Then
                    temp = vArray(i): vArray(i) = vArray(j): vArray(j) = temp
                End If
            Next j
        Next i
    End If
End Sub
