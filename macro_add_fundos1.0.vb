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
    
    MsgBox "Quantidade de slides: " & qtdSlide
    MsgBox "Slide inicial é: " & slideatual
    
    SelecionarImagem
    
    ' Verifica se uma imagem foi selecionada
    If caminhoImagem = "" Then
        MsgBox "Por favor, selecione uma imagem antes de continuar."
        Exit Sub
    End If
    
    Do While slideatual < qtdSlide + 1
        'MsgBox "Esse é o slide " & slideatual
        
        Set slide = ActivePresentation.Slides(slideatual) ' Inicializa a variável slide
        ' Definir a imagem como fundo do slide
        slide.FollowMasterBackground = msoFalse ' Desativar fundo padrão do mestre
        slide.Background.Fill.UserPicture (caminhoImagem)
        
        ActivePresentation.Slides(slideatual).Select
        slideatual = slideatual + 1
    Loop
    
    
    MsgBox "Imagens de fundo definidas com sucesso!"
End Sub
