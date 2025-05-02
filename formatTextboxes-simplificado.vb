' Esta função vba para powerpoint, formata todas as caixas de texto em uma apresentação
'para um formato padrão de fonte, cor e tamanho

Sub formatTextboxes()
    Dim shape As shape
    Dim slideatual As slide
    Dim pptPres As Presentation

    ' Utilize a apresentaÃ§Ã£o atualmente aberta
    Set pptPres = ActivePresentation

    ' Loop para verificar cada slide na apresentaÃ§Ã£o
    For Each slideatual In pptPres.Slides
        ' Loop para verificar cada shape no slide atual
        For Each shape In slideatual.Shapes
            ' Verifica se o shape Ã© uma caixa de texto
            If shape.HasTextFrame Then
                With shape.TextFrame.TextRange
                    .Text = UCase(.Text)  ' Converte o texto para maiúsculas
                    .Font.Name = "Montserrat"
                    .Font.Color = RGB(255, 255, 255) ' Define a cor da fonte
                    .Font.Size = 42 ' Define o tamanho da fonte
                End With
            End If
        Next shape
    Next slideatual
End Sub
