Attribute VB_Name = "Módulo3"
Global caminhoImagem As String

Sub SelecionarImagem()
    Dim dlg As FileDialog
    
    ' Cria a janela de diálogo para seleção de arquivo
    Set dlg = Application.FileDialog(msoFileDialogFilePicker)
    
    ' Define filtros para tipos de arquivo (por exemplo, imagens)
    dlg.Filters.Clear
    dlg.Filters.Add "Imagens", "*.jpg; *.jpeg; *.png; *.bmp; *.gif", 1
    
    ' Exibe a janela de diálogo
    If dlg.Show = -1 Then
        ' Salva o caminho do arquivo selecionado na variável
        caminhoImagem = dlg.SelectedItems(1)
        ' MsgBox "Caminho da imagem selecionada: " & caminhoImagem
    Else
        MsgBox "Nenhum arquivo foi selecionado."
    End If
End Sub

Sub DefinirFundoImagem()
    Dim slideatual As Integer
    slideatual = ActiveWindow.Selection.SlideRange.SlideIndex
    
    Dim qtdSlide As Integer
    qtdSlide = ActivePresentation.Slides.Count
    
    Dim slide As slide
    
    ' Chama a função para selecionar a imagem
    SelecionarImagem
    
    ' Verifica se uma imagem foi selecionada
    If caminhoImagem = "" Then
        MsgBox "Por favor, selecione uma imagem antes de continuar."
        Exit Sub
    End If
    
    ' Loop para percorrer todos os slides
    For slideatual = slideatual To qtdSlide
        Set slide = ActivePresentation.Slides(slideatual) ' Inicializa a variável slide
        slide.FollowMasterBackground = msoFalse ' Desativa o fundo padrão do mestre
        
        ' Remove imagens existentes no slide
        For Each shape In slide.Shapes
            If shape.Type = msoPicture Then
                shape.Delete ' Remove a imagem
            End If
        Next shape
        
        ' Define a imagem como fundo do slide
        slide.Background.Fill.UserPicture (caminhoImagem)
    Next slideatual
    
    'MsgBox "Imagens de fundo definidas com sucesso!"
End Sub

