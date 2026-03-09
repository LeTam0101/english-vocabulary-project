Attribute VB_Name = "Module1"
Sub ThemSlideChuan()
    Dim ppApp As PowerPoint.Application
    Dim ppPres As PowerPoint.Presentation
    
    ' K?t n?i v?i PowerPoint dang m?
    On Error Resume Next
    Set ppApp = GetObject(, "PowerPoint.Application")
    If ppApp Is Nothing Then
        MsgBox "M? PowerPoint tru?c khi ch?y macro!", vbExclamation
        Exit Sub
    End If
    On Error GoTo 0
    
    Set ppPres = ppApp.ActivePresentation
    
    ' Ðu?ng d?n file
    Dim txtFilePath As String: txtFilePath = "E:\ipa.txt"
    Dim audioFolderPath As String: audioFolderPath = "E:\ipa\"
    
    ' Ð?c file text
    Dim fileContent As String
    fileContent = ReadTextFileUTF8(txtFilePath)
    
    ' X? lý t?ng dòng
    Dim fileLines As Variant, i As Integer
    fileLines = Split(fileContent, vbLf)
    
    For i = UBound(fileLines) To LBound(fileLines) Step -1
        Dim line As String: line = Trim(fileLines(i))
        If line <> "" And InStr(line, ":") > 0 Then
            Dim parts As Variant: parts = Split(line, ":")
            Dim word As String: word = Trim(parts(0))
            Dim phonetic As String: phonetic = Trim(parts(1))
            
            ' Thêm slide vào d?u
            ThemSlideVoiAudioTuDong ppPres, word, phonetic, audioFolderPath, 1
        End If
    Next i
    
    ppApp.Activate
    MsgBox "Ðã thêm " & UBound(fileLines) + 1 & " slide vào d?u presentation!", vbInformation
End Sub

Sub ThemSlideVoiAudioTuDong(pres As Presentation, word As String, phonetic As String, audioFolder As String, pos As Integer)
    Dim ppSlide As Slide, ppShape As Shape
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim slideWidth As Single, slideHeight As Single
    
    ' T?o slide m?i ? v? trí ch? d?nh
    Set ppSlide = pres.Slides.Add(pos, ppLayoutBlank)
    slideWidth = pres.PageSetup.slideWidth
    slideHeight = pres.PageSetup.slideHeight
    
    ' --- THÊM T? (CAN GI?A) ---
    Set ppShape = ppSlide.Shapes.AddTextbox(msoTextOrientationHorizontal, 0, 0, slideWidth, 100)
    With ppShape
        .Left = 0
        .Top = slideHeight / 3 + 240
        .Width = 500
        .TextFrame.TextRange.Text = word
        .TextFrame.TextRange.Font.Size = 72
        .TextFrame.TextRange.Font.Name = "Arial"
        .TextFrame.TextRange.Font.Color.RGB = RGB(0, 0, 255)
        '.TextFrame.TextRange.Font.Bold = msoTrue
        .TextFrame.HorizontalAnchor = msoAnchorCenter
        .TextFrame.VerticalAnchor = msoAnchorMiddle
        .TextFrame.TextRange.ParagraphFormat.Alignment = ppAlignCenter
    End With
    
    ' --- THÊM PHIÊN ÂM ---
    phonetic = Replace(phonetic, " /", "/")
    phonetic = Replace(phonetic, "/ ", "/")
    If Left(phonetic, 1) <> "/" Then phonetic = "/" & phonetic
    If Right(phonetic, 1) <> "/" Then phonetic = phonetic & "/"
    
    Set ppShape = ppSlide.Shapes.AddTextbox(msoTextOrientationHorizontal, 0, 0, slideWidth, 60)
    With ppShape
        .Left = 390
        .Top = slideHeight / 3 + 250
        .Width = 500
        .TextFrame.TextRange.Text = phonetic
        .TextFrame.TextRange.Font.Size = 48
        .TextFrame.TextRange.Font.Name = "Arial Unicode MS"
        .TextFrame.TextRange.Font.Color.RGB = RGB(128, 0, 0)
        .TextFrame.HorizontalAnchor = msoAnchorCenter
        .TextFrame.VerticalAnchor = msoAnchorMiddle
        .TextFrame.TextRange.ParagraphFormat.Alignment = ppAlignCenter
    End With
    
    ' --- THÊM AUDIO (GÓC PH?I DU?I - KÍCH THU?C CHU?N) ---
    Dim audioPath As String: audioPath = audioFolder & word & ".mp3"
    If fso.FileExists(audioPath) Then
        Set ppShape = ppSlide.Shapes.AddMediaObject2(audioPath, False, True, slideWidth - 120, slideHeight - 80)
        
        ' THI?T L?P PLAYBACK
        With ppShape.AnimationSettings
            .AdvanceMode = ppAdvanceOnTime
            .AdvanceTime = 0
            With .PlaySettings
                .PlayOnEntry = msoTrue
                .PauseAnimation = msoFalse
                .HideWhileNotPlaying = msoTrue
                .LoopUntilStopped = msoTrue ' THÊM DÒNG NÀY Ð? B?T CH? Ð? LOOP
            End With
        End With
        
        ' Ð?M B?O TRONG PLAYBACK ÐÃ CH?N AUTOMATICALLY
        On Error Resume Next
        ppShape.MediaFormat.StartPoint = 0
        ppShape.MediaFormat.EndPoint = ppShape.MediaFormat.Length
        ppShape.MediaFormat.FadeInDuration = 0
        ppShape.MediaFormat.FadeOutDuration = 0
        On Error GoTo 0
    End If
End Sub

Function ReadTextFileUTF8(filePath As String) As String
    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")
    With stream
        .Charset = "UTF-8"
        .Open
        .LoadFromFile filePath
        ReadTextFileUTF8 = .ReadText
        .Close
    End With
    Set stream = Nothing
End Function
