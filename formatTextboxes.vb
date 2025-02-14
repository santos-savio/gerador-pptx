' Função 1: Deletar Títulos Vazios em Todos os Slides
Function DeletarTituloVazioEmTodosSlides()
    Dim slideAtual As slide
    Dim shapeAtual As shape
    Dim i As Integer
    Dim slide As slide

    ' Loop para verificar cada slide na apresentação
    For Each slide In ActivePresentation.Slides
        ' Loop para verificar cada shape no slide atual
        For i = slide.Shapes.Count To 1 Step -1
            Set shapeAtual = slide.Shapes(i)
            
            ' Verifica se o shape é um placeholder de título
            If shapeAtual.Type = msoPlaceholder Then
                If shapeAtual.PlaceholderFormat.Type = ppPlaceholderTitle Then
                    ' Verifica se o título está vazio
                    If Trim(shapeAtual.TextFrame.TextRange.Text) = "" Then
                        ' Deletar o placeholder de título vazio
                        shapeAtual.Delete
                    End If
                End If
            End If
        Next i
    Next slide
    
    ' Altera a aresolução para 16x9
    Dim pptPres As Presentation
    Set pptPres = ActivePresentation
    
    ' Definir tamanho do slide para widescreen (10 x 5.625 polegadas)
    With pptPres.PageSetup
        .SlideWidth = 10 * 72 ' 10 polegadas em pontos
        .SlideHeight = 5.625 * 72 ' 5.625 polegadas em pontos
    End With
    
End Function

' Função 2: Aplicar Estilos em Todos os Slides
Function AplicarEstilosEmTodosSlides()
    Dim slideAtual As slide
    Dim shape As shape
    Dim Fonte As String
    Dim TamanhoFonte As Integer
    Dim FractionVertical As Single
    Dim FractionHorizontal As Single

    ' Definir fonte, tamanho e posicionamento padrão
    Fonte = "BANDEX"
    TamanhoFonte = 40
    FractionVertical = 50 ' Ajuste conforme necessário
    FractionHorizontal = 30 ' Ajuste conforme necessário

    ' Loop para verificar cada slide na apresentação
    For Each slideAtual In ActivePresentation.Slides
        ' Loop para verificar cada shape no slide atual
        For Each shape In slideAtual.Shapes
            ' Verifica se o shape é uma caixa de texto
            If shape.HasTextFrame Then
                With shape.TextFrame.TextRange
                    .Font.Size = TamanhoFonte ' Define o tamanho da fonte
                    .Font.Name = Fonte ' Define a fonte
                    .Font.Color = RGB(255, 255, 255) ' Define a cor da fonte
                    .ParagraphFormat.Alignment = ppAlignCenter ' Centraliza o texto
                End With
                
                ' Posiciona a caixa de texto no slide
                With shape
                    .Left = (ActivePresentation.PageSetup.SlideWidth - .Width) * (FractionVertical / 100)
                    .Top = (ActivePresentation.PageSetup.SlideHeight - .Height) * (FractionHorizontal / 100)
                End With
            End If
        Next shape
    Next slideAtual
End Function

' Chamada das funções em sequência
Sub ExecutarMacrosEmSequencia()
    DeletarTituloVazioEmTodosSlides
    AplicarEstilosEmTodosSlides
End Sub