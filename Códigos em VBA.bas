

VBA MAIS RÁPIDO

Início:
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.DisplayAlerts = False
Application.DisplayStatusBar = False
Activesheet.DisplayPageBreaks = False


Fim:
Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.DisplayAlerts = True
Application.DisplayStatusBar = True
Activesheet.DisplayPageBreaks = True
OCULTAR POPUP DE ATUALIZAR VÍNCULOS
Workbooks.Open (Arquivo1), UpdateLinks:=0

REMOVER DUPLICATAS

Sub test()
 Sheets("IQ").Cells.RemoveDuplicates Columns:=2, Header:=xlYes
End Sub

COPIAR CONTEUDO DE UMA ABA EM UM NOVO ARQUIVO

Sub CopiaPlanilhaAtiva()
    Dim lPlanilha As String
    Dim lNome As String
    Dim lNovaPlanilha As String
   
    lPlanilha = ActiveWorkbook.Name
  
    lNome = ActiveSheet.Name
  
    Sheets(lNome).Select
    Sheets(lNome).Copy
   
    lNovaPlanilha = ActiveWorkbook.Name
   
    Windows(lPlanilha).Activate
    Cells.Select
    Selection.Copy
    Windows(lNovaPlanilha).Activate
    Cells.Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Range("A1").Select
    Windows(lPlanilha).Activate
    Range("A1").Select
    Application.CutCopyMode = False
End Sub

SELECIONAR PASTA PARA ABRIR ARQUIVO

'Procedimento para selecionar arquivos

Sub SelecionarArq1()
    Dim Folder As FileDialog
    Dim lArquivo As String
   
    'Chama o objeto passando os parâmetros
    Set Folder = Application.FileDialog(FileDialogType:=msoFileDialogOpen)
    With Folder
        'Alterar esta propriedade para True permitirá a seleção de vários arquivos
        .AllowMultiSelect = False
       
        'Determina a forma de visualização dos aruqivos
        .InitialView = msoFileDialogViewDetails
       
        'Determina qual o drive inicial
        .InitialFileName = "C:\"
       
        'Filtro de arquivos, pode ser colocado mais do que um filtro separando com ; por exemplo: "*.xls;*.xlsm"
        .Filters.Add "Texto", "*.txt;*.xls;*.xlsm;*.csv;*.xlsx", 1
    End With
   
    'Retorna o arquivo selecionado
    If Folder.Show = -1 Then
        lArquivo = Folder.SelectedItems(1)
        MsgBox "O arquivo selecionado está em: " & lArquivo
        Cells(4, 8).Value = lArquivo
    Else
        MsgBox "Nenhum arquivo foi selecionado."
    End If
End Sub

ABRI O ARQUIVO DE ACORDO COM A MACRO ACIMA

A
Sub ExecImport()
Application.DisplayAlerts = False
' ARQUIVO 1 = PLANILHA EXTRAÍDA DA PLATAFORMA FINANCEIRA '
' ARQUIVO 2 = PLANILHA DE PROPOSTAS COM ERRO '
Dim Arquivo1 As String
Dim Arquivo2 As String
   
    Sheets("EXECUTAVEL").Select
   
Arquivo1 = Range("H4") 'LOCAL CONTENDO A NOMENCLATURA DO RELATÓRIO'
Arquivo2 = Range("H6") 'LOCAL CONTENDO A NOMENCLATURA DO RELATÓRIO'
    NameArquivo1 = Dir(Arquivo1) 'APANHAR O NOME DO ARQUIVO 1'
    NameArquivo2 = Dir(Arquivo2) 'APANHAR O NOME DO ARQUIVO 2'

VERIFICAR SE TEM FILTRO E LIMPA TODOS OS FILTROS

Sub LimpFiltros()
'--------------------------------------------------------------------------------------------'
' LIMPAR FILTROS DA ABA SCRIPT '
'--------------------------------------------------------------------------------------------'
Sheets("SCRIPT").Select
        If ActiveSheet.FilterMode Then ' Verifica se está filtrada '
        ActiveSheet.ShowAllData     ' Limpa todos os filtros '
        End If
        Range("A3").Select
End Sub

ALTERAR FORMATO DATA AMERICA PARA PORTUGUES NO VBA

Columns("C:C").Select
    Selection.TextToColumns Destination:=Range("C1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, _
        Semicolon:=False, Comma:=True, Space:=False, Other:=False, FieldInfo _
        :=Array(Array(1, 4), Array(2, 1), Array(3, 1)), TrailingMinusNumbers:=True

LIMPAR ERRO EM NOME DE CONFLITOS - _FILTERDATABASE  

Sub ClearAutoFilters()
Dim ws As Worksheet
   On Error Resume Next
   For Each ws In ActiveWorkbook.Worksheets
       ws.AutoFilterMode = False
       ws.Names("_FilterDatabase").Delete
   Next ws
End Sub

MACRO MENSAGEM BOX QUANDO DAR ERRO NO VBA

Sub DESFAZER()
' DESFAZER AÇÃO DO RELATÓRIO '
On Error GoTo ErrMsg
Application.Undo
Range("A4").Select
Exit Sub
ErrMsg:
MsgBox ("ERROR: NÃO HÁ COMO DESFAZER A ALTERAÇÃO."), , "INDISPONÍVEL"
End Sub
' FIM '

DESFAZER

Application.Undo

ALTERAR FORMULAS PARA PORTUGUES E INGLES

Sub Ingles_Portugues()
ActiveCell = ActiveCell.Formula
End Sub
Sub Portugues_Ingles()
ActiveCell = "'" & ActiveCell.Formula
End Sub

EXCLUIR LINHAS COM DETERMINADO CRITÉRIO

Sub EXCLUIR()
    Dim lLin As Long
    Application.ScreenUpdating = False
    '.. ABA DETAILED HABITADO ..'
Sheets("ReF").Select
Range("A2").Select
    With Sheets("ReF")
        For lLin = .Cells(.Rows.Count, "AW").End(xlUp).Row To 2 Step -1
            If .Cells(lLin, "G") = "GUI" Then .Rows(lLin).Delete
           
            'Desafoga os processos pendentes do Windows a cada 100 linhas iteradas:
            If lLin Mod 100 = 0 Then DoEvents
        Next lLin
    End With
End Sub

INSERIR FÓRMULA NA LINHA ATÉ A ULTIMA PREENCHIDA

Sub InserirFormula()
Sheets("Detailed Comércio").Select
Range("A2").Select
Dim lr As Long
    lr = Cells(Rows.Count, "A").End(xlUp).Row
    Application.ScreenUpdating = 0
   
    Range("AJ2").Formula = "=IF(C2="""","""",IF(C2=C3,""NÃO"",""SIM""))"
    Range("AJ2").AutoFill Destination:=Range("AJ2:AJ" & lr)
   
    Application.ScreenUpdating = 1
End Sub

COLA COM VALORES A LINHA COM FÓRMULA

Sub ColarValores ()
    Range("AJ2:AJ10000").Copy
    Range("AJ2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone
    Range("A2").Select
    Application.CutCopyMode = False
   
End Sub
   
'......................................................................................................'
Application.StatusBar = "Processando, por favor aguarde..."
Application.Wait (Now + TimeValue("0:00:01"))
Application.ScreenUpdating = False
Application.ScreenUpdating = True
Application.StatusBar = "Importação finalizada com sucesso!"

IMPORTAR DADOS DE OUTRA PLANILHA

Sub Regionais()
dim caminho
CAMINHO = Environ("USERPROFILE") & "\Desktop"   `BUSCAR AUTOMATICO O DESKTOP DO USUÁRIO`
    Worksheets("BancoDeDados").Select `SELECIONAR ABA`
    Range("A5:M100000").ClearContents `LIMPAR RANGE`
    Range("A5").Select `SELECIONAR CÉLUDA`
    Workbooks.Open (CAMINHO & "\Ferramenta - Tratamento de Ordens\Bases\ABC.xlsx") `ABRIR PASTA NO DESKTOP DO USUÁRIO`
    Windows("ABC.xlsx").Activate `ABRIR A PASTA ABC NO CAMINHO ACIMA MENCIONADO`
    Range("N2:N100000", "Z2:Z100000").Copy `COPIAR RANGE DA PASTA CIMA`
    Windows("Tratamento de Ordens - v1").Activate `VOLTAR PARA A PASTA QUE SERÁ INCLUÍDO OS DADOS`
    Range("A5").Select `SELECIONAR CÉLULA`
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False `COLAR COMO VALORES`
    Windows("ABC.xlsx").Activate `SELECIONAR PASTA ABC`
    ActiveSheet.Range("A1").Copy `IMPEDI QUE A CAIXA DE COLAR TRANSFERENCIA SEJA EXIBIDA`
    ActiveWorkbook.Close savechanges:=False `FECHAR PASTA ABC`

IR PARA PRÓXIMA LINHA EM BRANCO

    Range("A5").Select
    Selection.End(xlDown).Select
    ActiveCell.Offset(1, 0).Select

PULAR PARA A PRÓXIMA LINHA

    ActiveSheet.Cells(ActiveCell.Row + 1, ActiveCell.Column).Select

TIRAR SELEÇÃO DA CÉLULA

Application.CutCopyMode = False

CHAMAR USEFORM

    Sub Formulario()
    UseForm1.Show
    End Sub

ABRIR SOMENTE O USEFORM

Sub Aplicativo_invisivel()
Application.Visible = False
UseForm1.Show
End Sub
Private Sub Workbook_Open() `ESTA PASTA DE TRABALHO`
Call Aplicativo_invisivel
Call TelaCheia_On
End Sub

SUB PARA OCULTAR E EXIBIR AS GUIAS, ABAS E LINHA DE REGRA

Inserir em “Estava pasta de Trabalho” – Selecionar “workbook” na listagem”
Private Sub Workbook_Open()
Call ocultar_tudo
End Sub
Sub TelaCheia_On() ‘OCULTAR GUIAS, ABAS E LINHAS DE GRADE DO EXCEL’
    On Error GoTo TelaCheia_On_Error
    Application.ExecuteExcel4Macro "SHOW.TOOLBAR(""Ribbon"",False)"
      Application.DisplayFormulaBar = False
    ActiveWindow.DisplayHeadings = False
    With ActiveWindow
        .DisplayHorizontalScrollBar = False
        .DisplayVerticalScrollBar = False
        .DisplayWorkbookTabs = False
        .DisplayHeadings = False
        .DisplayZeros = False
        .DisplayHeadings = False
        .DisplayGridlines = False
    End With
    With Application
        .DisplayFormulaBar = False
        .DisplayStatusBar = False
        .DisplayNoteIndicator = False
        .Caption = "RELATÓRIO RPC"
    End With
ActiveWindow.Caption = "COMGAS"
    On Error GoTo 0
    Exit Sub
TelaCheia_On_Error:
    MsgBox "Error"
End Sub
Sub TelaCheia_Off() ‘MOSTRAR GUIAS, ABAS E LINHAS DE GRADE DO EXCEL’
    On Error GoTo TelaCheia_Off_Error
    Application.ExecuteExcel4Macro "SHOW.TOOLBAR(""Ribbon"",True)"
    Application.DisplayFormulaBar = True
    ActiveWindow.DisplayHeadings = True
    With ActiveWindow
        .DisplayHorizontalScrollBar = True
        .DisplayVerticalScrollBar = True
        .DisplayWorkbookTabs = True
        .DisplayHeadings = True
        .DisplayZeros = True
        .DisplayHeadings = True
        .DisplayGridlines = True
    End With
    With Application
        .DisplayFormulaBar = True
        .DisplayStatusBar = True
        .DisplayNoteIndicator = True
       
    End With
    On Error GoTo 0
    Exit Sub
TelaCheia_Off_Error:
MsgBox "Error"
   
End Sub
DIMINUIR O CÁLCULO DA MACRO EXCEL

Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
MsgBox "Planilha salva com sucesso”   ‘Aparece a caixa com a mensagem ao finalizar Macro’
Desbloquear Planilha

Sub DesprotegerPlanilhaAtiva() ‘LIBERA PLANILHA COM SENHA”

Dim i, i1, i2, i3, i4, i5, i6 As Integer, j As Integer, k As Integer, l As Integer, m As Integer, n As Integer
On Error Resume Next
For i = 65 To 66
For j = 65 To 66
For k = 65 To 66
For l = 65 To 66
For m = 65 To 66
For i1 = 65 To 66
For i2 = 65 To 66
For i3 = 65 To 66
For i4 = 65 To 66
For i5 = 65 To 66
For i6 = 65 To 66
For n = 32 To 126
ActiveSheet.Unprotect Chr(i) & Chr(j) & Chr(k) & Chr(l) & Chr(m) & Chr(i1) & Chr(i2) & Chr(i3) & Chr(i4) & Chr(i5) & Chr(i6) & Chr(n)
If ActiveSheet.ProtectContents = False Then
MsgBox "Planilha desbloqueada com sucesso"
Exit Sub
End If
Next
Next
Next
Next
Next
Next
Next
Next
Next
Next
Next
Next
End Sub

COLOCAR MENSAGEM AO ABRIR EXCEL

Colocar o código abaixo em EstaPasta_de_tabalho
Private Sub Workbook_Open() 
MsgBox "Bem-Vindo ao Excel", vbInformation, "Pop-Up"
End Sub

PEDIR PARA SALVAR O EXCEL APÓS CLICAR EM SAIR PELA BOTÃO

Sub Fechar()
 bye = True
 ThisWorkbook.Close
 End Sub

ATUALIZAR PLANILHAS QUE POSSUEM SENHA

Sub Atualizar()
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
 'A linha abaixo insere a senha e pressiona enter
 Application.SendKeys ("banco1{ENTER}") ‘COLOCAR A SENHA DA PLANILHA’
ActiveWorkbook.UpdateLink Name:="A:\Banco de Horas - Luciana\01. Banco de Horas Gestão de Chamados.xlsm", Type:= _
 XlExcelLinks ‘COLOCAR O NOME DA PLANILHA’
‘SE POSSUIR MAIS PLANILHAS, PREENCHER CONFORME ACIMA’
 Application.SendKeys ("banco2{ENTER}")
 ActiveWorkbook.UpdateLink Name:="A:\Banco de Horas - Luciana\02. Banco de Horas Performance Operacional.xlsm", Type:= _
 xlExcelLinks
 Application.SendKeys ("banco3{ENTER}")
 ActiveWorkbook.UpdateLink Name:="A:\Banco de Horas - Luciana\03. Banco de Horas Obras em Andamento.xlsm", Type:= _
 xlExcelLinks
 Application.SendKeys ("banco4{ENTER}")
 ActiveWorkbook.UpdateLink Name:="A:\Banco de Horas - Luciana\04. Banco de Horas Gestão de Contratos.xlsm", Type:= _
 xlExcelLinks
 
 Application.SendKeys ("banco5{ENTER}")
 ActiveWorkbook.UpdateLink Name:="A:\Banco de Horas - Luciana\05. Banco de Horas Gestão de Danos.xlsm", Type:= _
 xlExcelLinks
 
 Application.SendKeys ("banco6{ENTER}")
 ActiveWorkbook.UpdateLink Name:="A:\Banco de Horas - Luciana\06. Banco de Horas Excelência Comercial.xlsm", Type:= _
 xlExcelLinks
 
  Application.SendKeys ("banco7{ENTER}")
 ActiveWorkbook.UpdateLink Name:="A:\Banco de Horas - Luciana\07. Banco de Horas Integridade Cadastral 1.xlsm", Type:= _
 xlExcelLinks
 
 Application.SendKeys ("banco8{ENTER}")
 ActiveWorkbook.UpdateLink Name:="A:\Banco de Horas - Luciana\08. Banco de Horas Integridade Cadastral 2.xlsm", Type:= _
 xlExcelLinks
 
 Application.ScreenUpdating = True 'ADICIONAR ANTERIOR A LINHA END"
Application.Calculation = xlCalculationAutomatic
 End Sub

OCULTAR BOTOES MINIMIAR, MAXIMIZAR E FECHAR DA BARRA DO EXCEL

Inserir em “Estava pasta de Trabalho” – Selecionar “workbook” na listagem”
Private Sub Workbook_Open()
    Application.ScreenUpdating = False
          Call RetiraXdaBarra
    Application.ScreenUpdating = True
End Sub
Private Sub Workbook_BeforeClose(Cancel As Boolean)
    Application.ScreenUpdating = False
   
        Call RepoeXdaBarra
       
    Application.ScreenUpdating = True
End Sub
Inserir em MODULO
Option Explicit  'OCULTA OS BOTÕES MINIMIZAR / MAXIMIZAR E FECHAR DA BARRA DO EXCEL
'Sem as declarações abaixo as macros para retirar e repor os botões não funcionará
Declare Function FindWindow32 Lib "user32" Alias "FindWindowA" _
        (ByVal lpClassName As Any, ByVal lpWindowName As Any) As Integer
Declare Function GetWindowLong32 Lib "user32" Alias "GetWindowLongA" _
        (ByVal hWnd As Integer, ByVal nIndex As Integer) As Long
Declare Function SetWindowLong32 Lib "user32" Alias "SetWindowLongA" _
        (ByVal hWnd As Integer, ByVal nIndex As Integer, _
        ByVal dwNewLong As Long) As Long
       
Global Const GWL_STYLE = (-16)
Global Const WS_SYSMENU = &H80000
Sub RepoeXdaBarra()
    Application.ScreenUpdating = True
    Dim WindowStyle As Long
    Dim hWnd As Integer
    Dim WindowName As String
    Dim Result As Variant
    WindowName = Application.Caption
    hWnd = FindWindow32(0&, ByVal WindowName)
    WindowStyle = GetWindowLong32(hWnd, GWL_STYLE)
   WindowStyle = WindowStyle Or WS_SYSMENU
    Result = SetWindowLong32(hWnd, GWL_STYLE, WindowStyle)
   
    'Força a barra de títulos a se atualizar, Retornando os Botões
    Application.Caption = "PROJETO COMGÁS"
    ActiveWindow.Caption = "BANCO DE HORAS"
End Sub
Sub RetiraXdaBarra()
Application.ScreenUpdating = False
    Dim WindowStyle As Long
    Dim hWnd As Integer
    Dim WindowName As String
    Dim Result As Variant
    WindowName = Application.Caption
    hWnd = FindWindow32(0&, ByVal WindowName)
    WindowStyle = GetWindowLong32(hWnd, GWL_STYLE)
    WindowStyle = WindowStyle And (Not WS_SYSMENU)
    Result = SetWindowLong32(hWnd, GWL_STYLE, WindowStyle)
    'Força a barra de títulos a se atualizar, Ocultando os Botões
    Application.Caption = "PROJETO COMGÁS"
    ActiveWindow.Caption = "BANCO DE HORAS"
End Sub

ESCREVE VALOR POR EXTENSO
INSERIR EM UM “NOVO MÓDULO” E INSESIR A FÓRMULA =EXTENSO

Function Extenso(nValor)
'
'escreve o valor em Reais por extenso
'
'
'Faz a validação do argumento
If IsNull(nValor) Or nValor <= 0 Or nValor > 9999999.99 Then
Exit Function
End If
'Declara as variáveis da função
Dim nContador, nTamanho As Integer
Dim cValor, cParte, cFinal As String
ReDim aGrupo(4), aTexto(4) As String
'Define matrizes com extensos parciais
ReDim aUnid(19) As String
aUnid(1) = "Um ": aUnid(2) = "Dois ": aUnid(3) = "Três "
aUnid(4) = "Quatro ": aUnid(5) = "Cinco ": aUnid(6) = "Seis "
aUnid(7) = "Sete ": aUnid(8) = "Oito ": aUnid(9) = "Nove "
aUnid(10) = "Dez ": aUnid(11) = "Onze ": aUnid(12) = "Doze "
aUnid(13) = "Treze ": aUnid(14) = "Quatorze ": aUnid(15) = "Quinze "
aUnid(16) = "Dezesseis ": aUnid(17) = "Dezessete ": aUnid(18) = "Dezoito "
aUnid(19) = "Dezenove "
ReDim aDezena(9) As String
aDezena(1) = "Dez ": aDezena(2) = "Vinte ": aDezena(3) = "Trinta "
aDezena(4) = "Quarenta ": aDezena(5) = "Cinquenta "
aDezena(6) = "Sessenta ": aDezena(7) = "Setenta ": aDezena(8) = "Oitenta "
aDezena(9) = "Noventa "
ReDim aCentena(9) As String
aCentena(1) = "Cento ": aCentena(2) = "Duzentos "
aCentena(3) = "Trezentos ": aCentena(4) = "Quatrocentos "
aCentena(5) = "Quinhentos ": aCentena(6) = "Seiscentos "
aCentena(7) = "Setecentos ": aCentena(8) = "Oitocentos "
aCentena(9) = "Novecentos "
'Divide o valor em vários grupos
cValor = Format$(nValor, "0000000000.00")
aGrupo(1) = Mid$(cValor, 2, 3)
aGrupo(2) = Mid$(cValor, 5, 3)
aGrupo(3) = Mid$(cValor, 8, 3)
aGrupo(4) = "0" + Mid$(cValor, 12, 2)
'Processa cada grupo
For nContador = 1 To 4
cParte = aGrupo(nContador)
nTamanho = Switch(Val(cParte) < 10, 1, Val(cParte) < 100, 2, Val(cParte) < 1000, 3)
If nTamanho = 3 Then
If Right$(cParte, 2) <> "00" Then
aTexto(nContador) = aTexto(nContador) + aCentena(Left(cParte, 1)) + "e "
nTamanho = 2
Else
aTexto(nContador) = aTexto(nContador) + IIf(Left$(cParte, 1) = "1", "Cem ", aCentena(Left(cParte, 1)))
End If
End If
If nTamanho = 2 Then
If Val(Right(cParte, 2)) < 20 Then
aTexto(nContador) = aTexto(nContador) + aUnid(Right(cParte, 2))
Else
aTexto(nContador) = aTexto(nContador) + aDezena(Mid(cParte, 2, 1))
If Right$(cParte, 1) <> "0" Then
aTexto(nContador) = aTexto(nContador) + "e "
nTamanho = 1
End If
End If
End If
If nTamanho = 1 Then
aTexto(nContador) = aTexto(nContador) + aUnid(Right(cParte, 1))
End If
Next
'Gera o formato final do texto
If Val(aGrupo(1) + aGrupo(2) + aGrupo(3)) = 0 And Val(aGrupo(4)) <> 0 Then
cFinal = aTexto(4) + IIf(Val(aGrupo(4)) = 1, "Centavo", "Centavos")
Else
cFinal = ""
cFinal = cFinal + IIf(Val(aGrupo(1)) <> 0, aTexto(1) + IIf(Val(aGrupo(1)) > 1, "Milhões ", "Milhão "), "")
If Val(aGrupo(2) + aGrupo(3)) = 0 Then
cFinal = cFinal + "De "
Else
cFinal = cFinal + IIf(Val(aGrupo(2)) <> 0, aTexto(2) + "Mil ", "")
End If
cFinal = cFinal + aTexto(3) + IIf(Val(aGrupo(1) + aGrupo(2) + aGrupo(3)) = 1, "Real", "Reais ")
cFinal = cFinal + IIf(Val(aGrupo(4)) <> 0, "e " + aTexto(4) + IIf(Val(aGrupo(4)) = 1, "Centavo", "Centavos"), "")
End If
Extenso = cFinal
End Function

BOTÃO GERAR PDF
INSERIR EM UM “NOVO MÓDULO” E ALTERAR O DESTINO DA PASTA

Sub Impressão()
'
' Relatório Macro
' imprimir nome
Range("I1").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    Range("I1").Select
    ActiveCell.Replace What:="/", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Find(What:="/", After:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
        xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False _
        , SearchFormat:=False).Activate
    Range("A1").Select
ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:= _
"A:\Gestão de Danos\OFICIOS 2015\" & ActiveSheet.Range("=I1").Value& & ".pdf", Quality:=xlQualityStandard, _
IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=False
MsgBox "PDF GERADO COM SUCESSO. FOI SALVO NA REDE (PASTA OFÍCIO 2015)"
End Sub

ABRIR SOMENTO O USEFORM
INSERIR EM “ESTA PASTA DE TRABALHO

Private Sub Workbook_Open()
Call Aplicativo_invisivel
End Sub
INSERIR EM “NOVO MÓDULO”
Sub Aplicativo_invisivel()
Application.Visible = False
UserForm1.Show
End Sub

EXEMPLOS DE BOTÕES DO USEFORM

Private Sub CommandButton1_Click()
    Sheets("CAPS FINANCEIRO PCON").Select
    Application.Visible = True
    UserForm1.Hide
    Range("A5").Select
   
End Sub
Private Sub CommandButton2_Click()
    Sheets("CAPS FINANCEIRO PDES").Select
    Application.Visible = True
    UserForm1.Hide
    Range("A5").Select
End Sub
Private Sub CommandButton3_Click()
    Sheets("GESTORES O E A").Select
    Application.Visible = True
    UserForm1.Hide
    Range("A3").Select
End Sub
Private Sub CommandButton4_Click()
    Sheets("APROVACAO").Select
    Application.Visible = True
    UserForm1.Hide
    Range("A3").Select
End Sub
Private Sub Label1_Click()
End Sub
Private Sub UserForm_Click()
End Sub
                
EXCLUIR LINHAS EM BRANCO

Sub DeletarLinhasVazias()

Dim UltimaLinha As Long
Dim r As Long
Dim Counter As Long

Application.ScreenUpdating = False

UltimaLinha = ActiveSheet.UsedRange.Rows.Count + ActiveSheet.UsedRange.Rows(1).Row - 1

For r = UltimaLinha To 1 Step -1
    If Application.WorksheetFunction.CountA(Rows(r)) = 0 Then
        Rows(r).Delete
        Counter = Counter + 1
    End If
Next r

Range("a1").Select

Application.ScreenUpdating = True

MsgBox Counter & " linhas vazias apagada(s).", vbInformation, "Linhas vazias"

End Sub

INSERIR LINHA EM BRANCO

Sub InsereLinha()
Dim i As Long
For i = 10007 To 1 Step -1
If Cells(i, "A") = "AVG" Then
Cells(i, "A").EntireRow.Insert
End If
Next i
For i = 10007 To 1 Step -1
If Cells(i, "A") = "AVG - D4" Then
Cells(i, "A").EntireRow.Insert
End If
Next i
For i = 10007 To 1 Step -1
If Cells(i, "A") = "AVG - D3" Then
Cells(i, "A").EntireRow.Insert
End If
Next i
For i = 10007 To 1 Step -1
If Cells(i, "A") = "AVG - D2" Then
Cells(i, "A").EntireRow.Insert
End If
Next i
For i = 10007 To 1 Step -1
If Cells(i, "A") = "AVG - D1" Then
Cells(i, "A").EntireRow.Insert
Exit Sub
End If
Next i
End Sub

CRITÉRIO PARA FILTRAR AUTOMÁTICO
    '===========================================================
    'Incluir somente essas linhas se desejar
   
    ActiveSheet.Range("$A$3:$AE$1000").AutoFilter Field:=31, Criteria1:= _
        "<>ZONA NÃO ENCONTRADA", Operator:=xlAnd, Criteria2:="<>"
    Range("AE3").Select
    '===================================================
    'Incluir somente essas linhas se desejar
   
    ActiveSheet.Range("$A$3:$AE$1000").AutoFilter Field:=22, Criteria1:= _
        "="
    Range("AE3").Select

CÉLULA SER IGUAL A OUTRA

Private Sub Worksheet_Calculate()
If Range("B5") <> Range("H5") Then
Range("H5") = Range("B5")
End If
End Sub

DESABILITAR SELECIONAR MAIS QUE UM ITEM NA PIVOT

Private Sub Worksheet_Change(ByVal Target As Range)
    If PivotTables("Tabela dinâmica2").PivotFields("MÊS").CurrentPage = "(All)" Then
        MsgBox "Selecione apenas um campo", vbCritical + vbOKOnly, "Filter Selection" 'optional message
        Application.Undo
        Exit Sub
    End If
End Sub

MANDAR E-MAIL (HABILITAR BIBLIOTECA)

Sub Envio()

Dim OutApp As Outlook.Application
Dim OutMail As Outlook.MailItem

'Criação e chamada do Objeto Outlook
Set OutApp = CreateObject("Outlook.Application")
Set OutMail = OutApp.CreateItem(olMailItem)

Application.DisplayAlerts = False

With OutMail
.To = "l.b.de.oliveira@accenture.com"
.CC = ""
.BCC = ""
.Subject = "Este é um e-mail de teste"
'no corpo de e-mail
.HTMLBody = "<b>Prezados.</b> <br/><br/>" & _
" Segue planilha. <br/><br/> Qualquer dúvida, estou à disposição<br/><br/><br/><br/>" & _
"Att,<br/> <b>Leonardo Oliveira<b/> <br/> Accenture<br/>"


'O trecho abaixo anexa a planilha ao e-mail
.Attachments.Add ActiveWorkbook.FullName
.Send 'Ou .Display para mostrar o email
End With

Application.DisplayAlerts = True

'Resetando a sessão
Set OutMail = Nothing
Set OutApp = Nothing

End Sub

#Criar hiperlink

Sub CriarHiperlink()
    
    Dim lUltimaLinhaAtiva As Long
    Dim lControle As Long

    Application.ScreenUpdating = False

    lUltimaLinhaAtiva = Worksheets("Lista de Encomendas").Cells(Worksheets("Lista de Encomendas").Rows.Count, 1).End(xlUp).Row
        
    For lControle = 2 To lUltimaLinhaAtiva
        Range("A" & lControle).Select
        ActiveSheet.Hyperlinks.Add Anchor:=Selection, Address:= _
            "http://websro.correios.com.br/sro_bin/txect01$.QueryList?P_LINGUA=001&P_TIPO=001&P_COD_UNI=" & _
            Range("A" & lControle).Value, TextToDisplay:="" & Range("A" & lControle).Value
    Next lControle
    Application.ScreenUpdating = True
    
End Sub

Sub RemoverHiperlink()

    Dim lUltimaLinhaAtiva As Long
    Dim lControle As Long

    Application.ScreenUpdating = False

    lUltimaLinhaAtiva = Worksheets("Lista de Encomendas").Cells(Worksheets("Lista de Encomendas").Rows.Count, 1).End(xlUp).Row
        
    For lControle = 2 To lUltimaLinhaAtiva
        Range("A" & lControle).Select
        Selection.Hyperlinks.Delete
    Next lControle

    Application.ScreenUpdating = True
End Sub

#ABRIR OUTRA PLANILHA

Private Sub CommandButton1_Click()
    Workbooks.Open ("C:\Documents and Settings\Administrador\Meus documentos\B.xls")
    Worksheets("Plan1").Activate
End Sub

#ABRIR ARQUIVO DO WORD

Dim objWord As Object

Set objWord = CreateObject("Word.application")

objWord.documents.Open "C:\Users\GilliardPacheco\Desktop\novo sistig\cartas\carta de colaboração.doc"

objWord.Visible = True
' SEU CODIGO ......

'objWord.Close
Sub Teste_Path()
Dim msg

msg = "C:\Users\" & VBA.Environ$("USERNAME") & "\Desktop\novo sistig\cartas\carta de colaboração.doc" & VBA.vbCr
msg = msg & "C:\Users\" & VBA.Environ$("USERNAME") & "\Desktop\novo sistig\cartas\carta de colaboração.doc"

MsgBox msg

End Sub