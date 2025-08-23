---
layout: default
title: Manual (PT)
permalink: /manual_pt
---

# LibreMacro – Manual Completo

## Índice
1. [Introdução](#introdução)  
2. [Primeiros passos](#primeiros-passos)  
3. [Funcionalidades](#funcionalidades)  
   - [Funções de Diálogo](#funções-de-diálogo)  
   - [Funções de Planilhas](#funções-de-planilhas)  
   - [Funções de Decoração e Formatação](#funções-de-decoração-e-formatação)  
   - [Funções de Animação](#funções-de-animação)  
   - [Funções de Modelos](#funções-de-modelos)  
   - [Funções de Conexão](#funções-de-conexão)  

---

## Introdução

A **LibreMacro** é uma biblioteca para facilitar a criação de macros no **LibreOffice Calc**.  
Ela oferece funções para manipular planilhas, criar diálogos, formatar células, animar efeitos visuais, gerar planilhas‑modelo e consumir dados externos.

Repositório oficial: <https://github.com/LibreMacro/Library>  
Curso gratuito: [Playlist no YouTube](https://www.youtube.com/playlist?list=PLw7mAr9L5qIy_19pSCvfQtlq2ro0digtn)

---

## Primeiros passos

1. Baixe o arquivo `LibreMacro.oxt` no repositório.  
2. LibreOffice → **Ferramentas** → **Gerenciador de Extensões** → **Adicionar…** → escolha o `.oxt`.  
3. Reinicie o LibreOffice. As funções ficarão disponíveis às suas macros Basic.

---

## Funcionalidades

### Funções de Diálogo

#### ConfirmDialog
Caixa de diálogo com **OK** e **Cancelar**.  
```basic
ConfirmDialog(pQuestion As String, Optional pDialogTitle As String) As Boolean
```
**Retorno:** `True` se OK; `False` se Cancelar.  

**Ex.:**
```basic
If ConfirmDialog("Deseja salvar?", "Confirmação") Then
    ' Ação se usuário clicou em ok
else
    ' Ação se usuário clicou em Cancelar
End If
```

---

#### QuestionDialog
Caixa de diálogo com **Sim** e **Não**.  
```basic
QuestionDialog(pQuestion As String, Optional pDialogTitle As String) As Boolean
```
**Retorno:** `True` se Sim; `False` se Não.  
**Ex.:**
```basic
If QuestionDialog("Imprimir agora?", "Pergunta") Then
    ' Ação se usuário clicou em Sim
else
    ' Ação se usuário clicou em Não
End If
```

---

#### RetryDialog
Caixa de diálogo com **Tentar de novo** e **Cancelar**.  
```basic
RetryDialog(pQuestion As String, Optional pDialogTitle As String) As Boolean
```
**Retorno:** `True` se Tentar de novo; `False` se Cancelar.  
**Ex.:**
```basic
Do While RetryDialog("Falha ao conectar. Tentar novamente?", "Erro")
    ' ...
Loop
```

---

#### Dialog  
Função utilitária interna usada pelas três acima para montar a MessageBox com ícone adequado.  
Você normalmente não chama `Dialog` diretamente — use `ConfirmDialog`, `QuestionDialog` ou `RetryDialog`.

#### Dialog3  
Variante utilitária de `Dialog` para cenários/ícones adicionais. Uso direto não é necessário na maioria dos casos.

---

### Funções de Planilhas

#### Cell
Acessa uma célula específica e retorna o objeto `CellRange`.  
```basic
Cell(pSheet As String, pCell As String) As Object
```
**Uso:**  
- `Cell(...).Value` → números  
- `Cell(...).String` → textos  
**Ex.:**
```basic
Cell("Planilha1","A1").Value = 42
Cell("Planilha1","B1").String = "Olá"
```

---

#### Sheet
Retorna uma referência para determinada planilha, cujo nome está informado no parâmetro de entrada pSheet.  
```basic
Sheet(pSheet As String) As Object
```
**Ex.:**
```basic
Dim s As Object
s = Sheet("Planilha1") 
```
Neste exemplo, a variável s terá uma referência para "Planilha1". Com isso podemos trabalhar com esse objeto como quisermos.

#### Row
Retorna uma referência para determinada linha, conforme parâmetro de entrada pRowindex.  
```basic
Row(pSheet As String, pRowIndex As Long) As Object
```
**Ex.:**
```basic
Row("Planilha1", 5).CellBackColor = RGB(240,240,240)  ' pinta a 6ª linha
```

#### ActiveSheet
Retorna a planilha ativa.  
```basic
ActiveSheet() As Object
```

#### ActiveSheetName
Retorna o **nome** da planilha ativa.  
```basic
ActiveSheetName() As String
```

#### SelectCell
Seleciona uma célula específica de determinada planilha.  
```basic
SelectCell(pSheet As String, pCell As String)
```
**Ex.:**
```basic
SelectCell("Planilha1","C3") 
```
Neste exemplo, a célula C3 da "Planilha1" será selecionada.

#### SelectRow
Seleciona uma **linha** inteira de determinada planilha.   
```basic
SelectRow(pSheet As String, pRowIndex As Long)
```

---

#### CreateSheet
Cria uma nova planilha.  
```basic
CreateSheet(pName As String)
```
**Ex.:**
```basic
CreateSheet("Relatório")
```

#### RemoveSheet
Remove uma planilha existente pelo nome.  
```basic
RemoveSheet(pName As String) 
```

**Ex.:**
```basic
RemoveSheet("Planilha1")
```
Neste exemplo, a "Planilha1" será removida do projeto. 

---

#### FindTextInCell
Procura um texto dentro de uma **célula**; `True` se contiver o trecho.  
```basic
FindTextInCell(pText As String, pCell As String) As Boolean
```
**Ex.:**
```basic
If FindTextInCell("OK","C5") Then ... End If 
```
Neste caso, a macro buscará o texto "OK" dentro da célula C5. Caso encontre, executará o que está determinado dentro da condição.

#### FindTextInColumn
Procura um texto em toda a **coluna** (da planilha indicada).  
```basic
FindTextInColumn(pSheet As String, pColIndex As Long, pText As String) As Boolean   ' (ver notas)
```
**Ex.:**
```basic
If FindTextInColumn("Planilha1", 0, "Cliente") Then ... End If 
```
Neste caso, a macro buscará a palavra "Cliente" dentro da "Planilha1" e, em específico, na primeira coluna (0 -> representa coluna A, ou seja, primeira coluna).

---

#### InsertRows
Insere **n** linhas a partir de um índice (1‑based para linha).  
```basic
InsertRows(pSheet As String, IndexL As Long, Units As Long)
```
**Ex.:**
```basic
InsertRows("Planilha1", 3, 2)  ' insere 2 linhas antes da linha 3
```

#### DeleteRows
Remove **n** linhas a partir de um índice.  
```basic
DeleteRows(pSheet As String, IndexL As Long, Units As Long)
```

---

#### InsertColumns
Insere **n** colunas a partir do índice **0‑based**.  
```basic
InsertColumns(pSheet As String, IndexC As Long, Units As Long)
```
**Ex.:**
```basic
InsertColumns("Planilha1", 2, 1)  ' insere antes da coluna C (A=0, B=1, C=2)
```

#### DeleteColumns
Remove **n** colunas a partir do índice **0‑based**.  
```basic
DeleteColumns(pSheet As String, IndexC As Long, Units As Long)
```

---

#### InsertCellNote
Insere/define **anotação** (nota) na célula.  
```basic
InsertCellNote(pSheet As String, pCell As String, pNote As String)
```
**Ex.:**
```basic
InsertCellNote("Planilha1","D4","Atenção: valor estimado")
```

#### RemoveCellNote
Remove a anotação da célula.  
```basic
RemoveCellNote(pSheet As String, pCell As String)
```

---

#### ClearContents
Limpa conteúdo de um **intervalo** (texto, números, fórmulas).  
```basic
ClearContents(pSheet As String, pRange As String)
```
**Ex.:**
```basic
ClearContents("Planilha1","A1:C100")
```

---

#### SortAsc
Ordena **crescente** um intervalo baseado em uma coluna de referência (0‑based).  
```basic
SortAsc(pSheet As String, pRange As String, pIndexC As Long)
```
**Ex.:**
```basic
SortAsc("Planilha1","A2:D100", 0)  ' ordena pela coluna A
```

#### SortDesc
Ordena **decrescente** um intervalo baseado em uma coluna de referência (0‑based).  
```basic
SortDesc(pSheet As String, pRange As String, pIndexC As Long)
```

---

### Funções de Decoração e Formatação

As funções abaixo alteram fonte, cor, estilo e criam padrões visuais.  

#### ChangeFontSize
Altera o **tamanho** da fonte em um intervalo.  
```basic
ChangeFontSize(pSheet As String, pRange As String, pSize As Integer)
```
**Ex.:** 
```basic
ChangeFontSize("Planilha1","A1:B5", 12)`
```
Neste caso, a macro irá alterar a fonte - independentemente do tamanho que esteja - para o tamanho 12.

#### ChangeFontColor
Altera a **cor** da fonte.  
```basic
ChangeFontColor(pSheet As String, pRange As String, pColor As Long)
```

**Ex.:** 
```basic
ChangeFontColor("Planilha1","A1:B5", RGB(0,90,180))
```

Neste caso, as cores do intervalo de células citado acima (A1:B5) ficará com a cor <span style="color:rgb(0,90,180)">RGB(0,90,180)</span>

#### ChangeCellColor
Altera a **cor de fundo** das células.  
```basic
ChangeCellColor(pSheet As String, pRange As String, pColor As Long)
```

**Ex.:** 
```basic
ChangeFontColor("Planilha1","A1:B5", "yellow")
```

Neste caso, as cores do intervalo de células citado acima (A1:B5) ficará com a cor <span style="color:yellow"> amarela</span>.

#### ChangeCellColor
Altera a **cor de fundo** das células.  
```basic
ChangeCellColor(pSheet As String, pRange As String, pColor As Long)
```


#### ChangeCellStyle
Aplica um **estilo** de célula existente no documento.  
```basic
ChangeCellStyle(pSheet As String, pRange As String, pStyleName As String)
```
**Ex.:** `ChangeCellStyle("Planilha1","A1:D1","Heading 1")`

#### ChangeFontFormat
Ativa/desativa formatações como **Negrito/Itálico/Sublinhado**.  
```basic
ChangeFontFormat(pSheet As String, pRange As String, pBold As Boolean, _
                 Optional pItalic As Boolean, Optional pUnderline As Boolean)
```
**Ex.:** `ChangeFontFormat("Planilha1","B2:B10", True, True, False)`

#### CreateStripedLines
Cria **linhas zebradas** (listras alternadas) em um intervalo.  
```basic
CreateStripedLines(pSheet As String, pRange As String, _
                   Optional pColor1 As Long, Optional pColor2 As Long, _
                   Optional pStripeHeight As Integer)
```
**Ex.:**
```basic
CreateStripedLines("Planilha1","A2:D100", RGB(248,248,248), RGB(235,235,235), 1)
```

#### ChangeFont
Altera a **família/tipo** da fonte.  
```basic
ChangeFont(pSheet As String, pRange As String, pFamily As String)
```
**Ex.:** `ChangeFont("Planilha1","A1:C10","Liberation Sans")`

#### CopyFontColor
Copia a **cor da fonte** de uma célula origem para um intervalo destino.  
```basic
CopyFontColor(pSheet As String, pSourceCell As String, pTargetRange As String)
```
**Ex.:** `CopyFontColor("Planilha1","A1","B1:B50")`

---

### Funções de Animação

Efeitos visuais básicos (passo a passo) sobre um intervalo.

#### AnimateFontSize
Anima o **tamanho da fonte** entre dois valores.  
```basic
AnimateFontSize(pSheet As String, pRange As String, _
                pFrom As Integer, pTo As Integer, _
                Optional pSteps As Integer, Optional pDelayMs As Long)
```
**Ex.:**
```basic
AnimateFontSize("Planilha1","A1", 10, 16, 6, 60)
```

Neste exemplo, a célula A1 da "Planilha1" fará uma animação partindo do tamanho 10 até alcançar o tamanho 16, sendo essa anumação feito em 6 etapas e durante 60ms.


#### AnimateFontColor
Anima a **cor da fonte** (interpolando entre duas cores).  
```basic
AnimateFontColor(pSheet As String, pRange As String, _
                 pColorFrom As Long, pColorTo As Long, _
                 Optional pSteps As Integer, Optional pDelayMs As Long)
```

```basic
AnimateFontColor("Planilha1", "B2:B10", "gray", "blue")
```
Neste exemplo, o intervalo de células (B2:B10) recebe uma animação em que a fonte inicialmente aparece na cor cinza e vai transicionanando para a cor azul. 

#### ChangeFont (animação)
Atalho animado para alterar a fonte com efeito gradual.  
```basic
ChangeFont(pSheet As String, pRange As String, pFamily As String, _
           Optional pSteps As Integer, Optional pDelayMs As Long)
```

#### ToggleCellColor
Cria uma animação em que a **cor de fundo** de determinada célula fica alternando entre duas cores, N vezes. Por padrão, essa repetição ocorre 5 vezes.

```basic
ToggleCellColor(pSheet As String, pRange As String, _
                pColor1 As Long, pColor2 As Long, _
                Optional pTimes As Integer, Optional pDelayMs As Long)
```
**Ex.:** 
```basic 
ToggleCellColor("Planilha1","A1", "yellow", "red", 10)
```
Neste exemplo, a cor de fundo da célula A1 fica alterando entre amarelho e vermelho por 10 vezes. Ao término, fica estabelecida a cor informada por último, ou seja, vermelho.

---

### Funções de Modelos


#### CreateTemplateSheet
Cria planilhas‑modelo de forma rápida (nomes e numeração).  

```basic
CreateTemplateSheet(pOption As String, Optional pQty As Integer)
```
**Ex.:**
```basic
' Cria 12 planilhas: Jan..Dez ou “Página 1..n”, dependente da opção
CreateTemplateSheet("mensal", 12)
```

---

### Funções de Conexão

#### GetXMLContent
Busca conteúdo **XML** e atribui a determinada célula.  

```basic
GetXMLContent(pUrl As String, pTag As String) As String
```
 
**Ex.:**
```basic
Dim valor As String
valor = GetXMLContent("https://meu-servidor/api.xml", "/preco/teste")
```


## Créditos
Um software livre criado por Marcos Cabanas Esteves e Thiago Andrade. O projeto pode ser usado gratuitamente, sem custos, e é aberto à participação de todos que desejarem contribuir.
