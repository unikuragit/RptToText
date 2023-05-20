Imports CRAXDRT = CRAXDDRT
Imports CRFMRC = CRAXDDRT.CRObjectFormatConditionFormulaType
Imports CRFMRB = CRAXDDRT.CRBorderConditionFormulaType
Imports CRFMSA = CRAXDDRT.CRSectionAreaFormatConditionFormulaType
Imports CRFMFF = CRAXDDRT.CRCommonFieldFormatConditionFormulaType
Imports CRFMNF = CRAXDDRT.CRNumericFieldFormatConditionFormulaType
Imports CRFMFT = CRAXDDRT.CRFontColorConditionFormulaType
Imports CRFMDF = CRAXDDRT.CRDateFieldFormatConditionFormulaType
Imports CRFMTF = CRAXDDRT.CRTimeFieldFormatConditionFormulaType
Module Module1

    'enumFomuras  Array(n,0) Is Japanese.

    Private m_objFomuraSec As Object(,) = New Object(,) {{"CSSクラス名                   ", CRFMSA.crSectionAreaCssClassConditionFormulaType} _
                                                       , {"非表示-ドリルダウン不可       ", CRFMSA.crSectionAreaEnableSuppressConditionFormulaType} _
                                                       , {"非表示-ドリルダウン可         ", CRFMSA.crSectionAreaEnableHideForDrillDownConditionFormulaType} _
                                                       , {"ページ下部へ出力              ", CRFMSA.crSectionAreaEnablePrintAtBottomOfPageConditionFormulaType} _
                                                       , {"出力前に改ページ              ", CRFMSA.crSectionAreaEnableNewPageBeforeConditionFormulaType} _
                                                       , {"出力後に改ページ              ", CRFMSA.crSectionAreaEnableNewPageAfterConditionFormulaType} _
                                                       , {"出力後にページ番号をリセット  ", CRFMSA.crSectionAreaEnableResetPageNumberAfterConditionFormulaType} _
                                                       , {"まとめて表示                  ", CRFMSA.crSectionAreaEnableKeepTogetherConditionFormulaType} _
                                                       , {"空セクションの非表示          ", CRFMSA.crSectionAreaEnableSuppressIfBlankConditionFormulaType} _
                                                       , {"続くセクションをアンダーレイ  ", CRFMSA.crSectionAreaEnableUnderlaySectionConditionFormulaType} _
                                                       , {"?                             ", CRFMSA.crSectionAreaShowAreaConditionFormulaType} _
                                                       , {"背景色                        ", CRFMSA.crSectionAreaBackgroundColorConditionFormulaType}}



    Private m_objFomuraCmn As Object(,) = New Object(,) {{"CSSクラス名                   ", CRFMRC.crCssClassConditionFormulaType} _
                                                       , {"非表示                        ", CRFMRC.crEnableSuppressConditionFormulaType} _
                                                       , {"横位置                        ", CRFMRC.crHorizontalAlignmentConditionFormulaType} _
                                                       , {"オブジェクトごとにまとめて表示", CRFMRC.crEnableKeepTogetherConditionFormulaType} _
                                                       , {"ページ区切りで境界線を閉じる  ", CRFMRC.crEnableCloseAtPageBreakConditionFormulaType} _
                                                       , {"複数行に出力                  ", CRFMRC.crEnableCanGrowConditionFormulaType} _
                                                       , {"ツールヒントテキスト          ", CRFMRC.crToolTipTextConditionFormulaType} _
                                                       , {"テキストの回転                ", CRFMRC.crRotationConditionFormulaType} _
                                                       , {"重複データの非表示            ", CRFMFF.crSuppressIfDuplicatedConditionFormulaType} _
                                                       , {"文字列の表示                  ", CRFMFF.crUseSystemDefaultConditionFormulaType}}


    Private m_objFomuraBdr As Object(,) = New Object(,) {{"線左                          ", CRFMRB.crLeftLineStyleConditionFormulaType} _
                                                       , {"線上                          ", CRFMRB.crTopLineStyleConditionFormulaType} _
                                                       , {"線右                          ", CRFMRB.crRightLineStyleConditionFormulaType} _
                                                       , {"線下                          ", CRFMRB.crBottomLineStyleConditionFormulaType} _
                                                       , {"幅の調節                      ", CRFMRB.crTightHorizontalConditionFormulaType} _
                                                       , {"影付き                        ", CRFMRB.crHasDropShadowConditionFormulaType} _
                                                       , {"境界線                        ", CRFMRB.crBorderColorConditionFormulaType} _
                                                       , {"背景                          ", CRFMRB.crBackgroundColorConditionFormulaType} _
                                                       , {"高さの調節                    ", CRFMRB.crTightVerticalConditionFormulaType}}



    Private m_objFomuraYmd As Object(,) = New Object(,) {{"日付型                        ", CRFMDF.crWindowsDefaultTypeConditionFormulaType} _
                                                       , {"カレンダーの種類              ", CRFMDF.crCalendarTypeConditionFormulaType} _
                                                       , {"形式-月                       ", CRFMDF.crMonthFormatConditionFormulaType} _
                                                       , {"形式-日                       ", CRFMDF.crDayFormatConditionFormulaType} _
                                                       , {"形式-年                       ", CRFMDF.crYearFormatConditionFormulaType} _
                                                       , {"形式-年/期間の種類            ", CRFMDF.crEraTypeConditionFormulaType} _
                                                       , {"順序                          ", CRFMDF.crDateOrderConditionFormulaType} _
                                                       , {"曜日-種類                     ", CRFMDF.crDayOfWeekTypeConditionFormulaType} _
                                                       , {"曜日-区切り                   ", CRFMDF.crDayOfWeekSeparatorConditionFormulaType} _
                                                       , {"曜日-カッコ                   ", CRFMDF.crDayOfWeekEnclosureConditionFormulaType} _
                                                       , {"曜日-位置                     ", CRFMDF.crDayOfWeekPositionConditionFormulaType} _
                                                       , {"区切り-前置記号               ", CRFMDF.crDatePrefixSeparatorConditionFormulaType} _
                                                       , {"区切り-第1区切り              ", CRFMDF.crDateFirstSeparatorConditionFormulaType} _
                                                       , {"区切り-第2区切り              ", CRFMDF.crDateSecondSeparatorConditionFormulaType} _
                                                       , {"区切り-後置記号               ", CRFMDF.crDateSuffixSeparatorConditionFormulaType}}


    Private m_objFomuraTme As Object(,) = New Object(,) {{"記号の位置                    ", CRFMTF.crAMPMFormatConditionFormulaType} _
                                                       , {"午前                          ", CRFMTF.crAMStringConditionFormulaType} _
                                                       , {"時                            ", CRFMTF.crHourFormatConditionFormulaType} _
                                                       , {"時/分の区切り                 ", CRFMTF.crHourMinuteSeparatorConditionFormulaType} _
                                                       , {"分                            ", CRFMTF.crMinuteFormatConditionFormulaType} _
                                                       , {"分/秒の区切り                 ", CRFMTF.crMinuteSecondSeparatorConditionFormulaType} _
                                                       , {"午後                          ", CRFMTF.crPMStringConditionFormulaType} _
                                                       , {"秒                            ", CRFMTF.crSecondFormatConditionFormulaType} _
                                                       , {"12時間制                      ", CRFMTF.crTimeBaseConditionFormulaType}}

    Private m_objFomuraNum As Object(,) = New Object(,) {{"通貨記号を使用可能にする      ", CRFMNF.crCurrencySymbolFormatConditionFormulaType} _
                                                        , {"ページごとに1記号             ", CRFMNF.crHasOneSymbolPerPageConditionFormulaType} _
                                                        , {"位置                          ", CRFMNF.crCurrencyPositionConditionFormulaType} _
                                                        , {"通貨記号                      ", CRFMNF.crCurrencySymbolConditionFormulaType} _
                                                        , {"ゼロを表示しない              ", CRFMNF.crEnableSuppressIfZeroConditionFormulaType} _
                                                        , {"小数点区切り                  ", CRFMNF.crDecimalSymbolConditionFormulaType} _
                                                        , {"小数点以下の桁数              ", CRFMNF.crNDecimalPlacesConditionFormulaType} _
                                                        , {"桁区切りを使用する            ", CRFMNF.crThousandsSeparatorConditionFormulaType} _
                                                        , {"端数処理                      ", CRFMNF.crRoundingFormatConditionFormulaType} _
                                                        , {"記号                          ", CRFMNF.crThousandSymbolConditionFormulaType} _
                                                        , {"負数の表示形式                ", CRFMNF.crNegativeFormatConditionFormulaType} _
                                                        , {"先頭のゼロを表示する          ", CRFMNF.crEnableUseLeadZeroConditionFormulaType} _
                                                        , {"符号を逆にして表示する        ", CRFMNF.crDisplayReverseSignConditionFormulaType} _
                                                        , {"ﾌｨｰﾙﾄﾞｸﾘｯﾋﾟﾝｸﾟを有効にする    ", CRFMNF.crAllowFieldClippingConditionFormulaType} _
                                                        , {"0の値を示す記号               ", CRFMNF.crZeroValueStringConditionFormulaType}}



    Private m_objFomuraFnt As Object(,) = New Object(,) {{"フォント名                    ", CRFMFT.crFontConditionFormulaType} _
                                                       , {"スタイル                      ", CRFMFT.crFontStyleConditionFormulaType} _
                                                       , {"サイズ                        ", CRFMFT.crFontSizeConditionFormulaType} _
                                                       , {"色                            ", CRFMFT.crColorConditionFormulaType} _
                                                       , {"取り消し線                    ", CRFMFT.crFontStrikeOutConditionFormulaType} _
                                                       , {"下線                          ", CRFMFT.crFontUnderLineConditionFormulaType}}



    Private m_objFomuraInd As Object(,) = New Object(,) {{"テキストの形式                ", CRAXDDRT.CRStringFieldConditionFormulaType.crTextInterpretationConditionFormulaType}}



    Private m_objFomuraLnk As Object(,) = New Object(,) {{"ハイパーリンク                ", CRFMRC.crHyperLinkConditionFormulaType}}


    Private buffer As New System.Text.StringBuilder()

    Sub Main()

        Dim app As CRAXDRT.Application

        Dim args As String() = System.Environment.GetCommandLineArgs()
        Dim rptFile As String = String.Empty
        Dim outDir As String = String.Empty
        Dim outFile As String = String.Empty
        Dim isDir As Boolean = False

        If args.Length <= 1 Then
            Console.Error.WriteLine("Invalid args.")
            Return
        End If
        If Not System.IO.File.Exists(args(1)) AndAlso Not System.IO.Directory.Exists(args(1)) Then
            Console.Error.WriteLine("Not exists file.")
            Return
        End If
        rptFile = System.IO.Path.GetFullPath(args(1))

        If System.IO.Directory.Exists(rptFile) Then
            If args.Length <= 2 Then
                Console.Error.WriteLine("Need output dir args.")
                Return
            End If

            If Not System.IO.Directory.Exists(args(2)) Then
                Console.Error.WriteLine("Invalid output dir.")
                Return
            End If
            outDir = System.IO.Path.GetFullPath(args(2))

            app = New CRAXDRT.Application()
            For Each file As String In System.IO.Directory.GetFiles(rptFile, "*.rpt")

                'Load CrystalReports File
                Dim Rpt As CRAXDRT.Report = app.OpenReport(file, 1)

                'Debug Print Formula Code
                Try
                    exp(Rpt, "")
                    Using fio As New System.IO.StreamWriter(System.IO.Path.Combine(outDir, System.IO.Path.GetFileName(file)) & ".txt", False, System.Text.Encoding.GetEncoding("shift_jis"))
                        fio.WriteLine(buffer)
                    End Using
                Catch ex As Exception
                    Console.Error.WriteLine(file & ":" & "Execute error. " & ex.Message)
                End Try
                If buffer.Length > 0 Then
                    buffer.Remove(0, buffer.Length)
                End If

            Next
        Else

            app = New CRAXDRT.Application()
            'Load CrystalReports File
            Dim Rpt As CRAXDRT.Report = app.OpenReport(rptFile, 1)


            'Debug Print Formula Code
            Try
                exp(Rpt, "")
                Console.Out.WriteLine(buffer.ToString())
            Catch ex As Exception
                Console.Error.WriteLine(rptFile & ":" & "Execute error. " & ex.Message)
            End Try


        End If

    End Sub

    Private Sub props(ByVal objRpt As Object, Optional ByVal type As String = "")
        Try
            buffer.AppendLine("AmPmType = " & objRpt.AmPmType)
        Catch ex As Exception
        End Try
        Try
            buffer.AppendLine("AmString = " & objRpt.AmString)
        Catch ex As Exception
        End Try
        Try
            buffer.AppendLine("BackColor = " & objRpt.BackColor)
        Catch ex As Exception
        End Try
        Try
            buffer.AppendLine("BooleanOutputType = " & objRpt.BooleanOutputType)
        Catch ex As Exception
        End Try
        Try
            buffer.AppendLine("BorderColor = " & objRpt.BorderColor)
        Catch ex As Exception
        End Try
        Try
            buffer.AppendLine("Bottom = " & objRpt.Bottom)
        Catch ex As Exception
        End Try
        Try
            buffer.AppendLine("BottomLineStyle = " & objRpt.BottomLineStyle)
        Catch ex As Exception
        End Try
        Try
            buffer.AppendLine("CaseInsensitiveSQLData = " & objRpt.CaseInsensitiveSQLData)
        Catch ex As Exception
        End Try
        Try
            buffer.AppendLine("CanGrow = " & objRpt.CanGrow)
        Catch ex As Exception
        End Try
        Try
            buffer.AppendLine("CharacterSpacing = " & objRpt.CharacterSpacing)
        Catch ex As Exception
        End Try
        Try
            Call GeEnumValue(type, "ConvertDateTimeType", objRpt.ConvertDateTimeType)
            'buffer.AppendLine("ConvertDateTimeType = " & objRpt.ConvertDateTimeType)
        Catch ex As Exception
        End Try
        Try
            buffer.AppendLine("ConvertNullFieldToDefault = " & objRpt.ConvertNullFieldToDefault)
        Catch ex As Exception
        End Try
        Try
            buffer.AppendLine("CloseAtPageBreak = " & objRpt.CloseAtPageBreak)
        Catch ex As Exception
        End Try
        Try
            buffer.AppendLine("CornerEllipseHeight = " & objRpt.CornerEllipseHeight)
        Catch ex As Exception
        End Try
        Try
            buffer.AppendLine("CornerEllipseWidth = " & objRpt.CornerEllipseWidth)
        Catch ex As Exception
        End Try
        Try
            buffer.AppendLine("CssClass = " & objRpt.CssClass)
        Catch ex As Exception
        End Try
        Try
            buffer.AppendLine("CurrencyPositionType = " & objRpt.CurrencyPositionType)
        Catch ex As Exception
        End Try
        Try
            buffer.AppendLine("CurrencySymbol = " & objRpt.CurrencySymbol)
        Catch ex As Exception
        End Try
        Try
            buffer.AppendLine("CurrencySymbolType = " & objRpt.CurrencySymbolType)
        Catch ex As Exception
        End Try
        Try
            buffer.AppendLine("DateCalendarType = " & objRpt.DateCalendarType)
        Catch ex As Exception
        End Try
        Try
            buffer.AppendLine("DateEraType = " & objRpt.DateEraType)
        Catch ex As Exception
        End Try
        Try
            buffer.AppendLine("DateFirstSeparator = " & objRpt.DateFirstSeparator)
        Catch ex As Exception
        End Try
        Try
            buffer.AppendLine("DateOrder = " & objRpt.DateOrder)
        Catch ex As Exception
        End Try
        Try
            buffer.AppendLine("DatePrefixSeparator = " & objRpt.DatePrefixSeparator)
        Catch ex As Exception
        End Try
        Try
            buffer.AppendLine("DateSecondSeparator = " & objRpt.DateSecondSeparator)
        Catch ex As Exception
        End Try
        Try
            buffer.AppendLine("DateWindowsDefaultType = " & objRpt.DateWindowsDefaultType)
        Catch ex As Exception
        End Try
        Try
            buffer.AppendLine("DateSuffixSeparator = " & objRpt.DateSuffixSeparator)
        Catch ex As Exception
        End Try
        Try
            buffer.AppendLine("DayType = " & objRpt.DayType)
        Catch ex As Exception
        End Try
        Try
            buffer.AppendLine("DecimalPlaces = " & objRpt.DecimalPlaces)
        Catch ex As Exception
        End Try
        Try
            buffer.AppendLine("DecimalSymbol = " & objRpt.DecimalSymbol)
        Catch ex As Exception
        End Try
        Try
            buffer.AppendLine("DisplayGrid = " & objRpt.DisplayGrid)
        Catch ex As Exception
        End Try
        Try
            buffer.AppendLine("DisplayHiddenSections = " & objRpt.DisplayHiddenSections)
        Catch ex As Exception
        End Try
        Try
            buffer.AppendLine("DisplayReverseSign = " & objRpt.DisplayReverseSign)
        Catch ex As Exception
        End Try
        Try
            buffer.AppendLine("DisplayRulers = " & objRpt.DisplayRulers)
        Catch ex As Exception
        End Try
        Try
            buffer.AppendLine("EnableOnDemand = " & objRpt.EnableOnDemand)
        Catch ex As Exception
        End Try
        Try
            buffer.AppendLine("EnableSnapToGrid = " & objRpt.EnableSnapToGrid)
        Catch ex As Exception
        End Try
        Try
            buffer.AppendLine("EnableTightHorizontal = " & objRpt.EnableTightHorizontal)
        Catch ex As Exception
        End Try
        Try
            buffer.AppendLine("ExtendToBottomOfSection = " & objRpt.ExtendToBottomOfSection)
        Catch ex As Exception
        End Try
        Try
            buffer.AppendLine("FirstLineIndent = " & objRpt.FirstLineIndent)
        Catch ex As Exception
        End Try
        Try
            Call GeEnumValue(type, "FieldMappingType", objRpt.FieldMappingType)
            'buffer.AppendLine("FieldMappingType = " & objRpt.FieldMappingType)
        Catch ex As Exception
        End Try
        Try
            buffer.AppendLine("FillColor = " & objRpt.FillColor)
        Catch ex As Exception
        End Try
        Try
            Call GeEnumValue(type, "FormulaSyntax", objRpt.FormulaSyntax)
            'buffer.AppendLine("FormulaSyntax = " & objRpt.FormulaSyntax)
        Catch ex As Exception
        End Try
        Try
            buffer.AppendLine("Font.Bold = " & objRpt.Font.Bold)
        Catch ex As Exception
        End Try
        Try
            buffer.AppendLine("Font.GdiCharSet = " & objRpt.Font.GdiCharSet)
        Catch ex As Exception
        End Try
        Try
            buffer.AppendLine("Font.GdiVerticalFont = " & objRpt.Font.GdiVerticalFont)
        Catch ex As Exception
        End Try
        Try
            buffer.AppendLine("Font.Italic = " & objRpt.Font.Italic)
        Catch ex As Exception
        End Try
        Try
            buffer.AppendLine("Font.Name = " & objRpt.Font.Name)
        Catch ex As Exception
        End Try
        Try
            buffer.AppendLine("Font.Size = " & objRpt.Font.Size)
        Catch ex As Exception
        End Try
        Try
            buffer.AppendLine("Font.Strikeout = " & objRpt.Font.Strikeout)
        Catch ex As Exception
        End Try
        Try
            buffer.AppendLine("Font.Underline = " & objRpt.Font.Underline)
        Catch ex As Exception
        End Try
        Try
            buffer.AppendLine("Font.Unit = " & objRpt.Font.Unit)
        Catch ex As Exception
        End Try
        Try
            buffer.AppendLine("GridSize = " & objRpt.GridSize)
        Catch ex As Exception
        End Try
        Try
            buffer.AppendLine("GroupNumber = " & objRpt.GroupNumber)
        Catch ex As Exception
        End Try
        Try
            buffer.AppendLine("HasDropShadow = " & objRpt.HasDropShadow)
        Catch ex As Exception
        End Try
        Try
            buffer.AppendLine("Height = " & objRpt.Height)
        Catch ex As Exception
        End Try
        Try
            buffer.AppendLine("HorAlignment = " & objRpt.HorAlignment)
        Catch ex As Exception
        End Try
        Try
            buffer.AppendLine("HourMinuteSeparator = " & objRpt.HourMinuteSeparator)
        Catch ex As Exception
        End Try
        Try
            buffer.AppendLine("HourType = " & objRpt.HourType)
        Catch ex As Exception
        End Try
        Try
            buffer.AppendLine("HyperlinkText = " & objRpt.HyperlinkText)
        Catch ex As Exception
        End Try
        Try
            buffer.AppendLine("HyperlinkType = " & objRpt.HyperlinkType)
        Catch ex As Exception
        End Try
        Try
            buffer.AppendLine("KeepTogether = " & objRpt.KeepTogether)
        Catch ex As Exception
        End Try
        Try
            Call GeEnumValue(type, "Kind", objRpt.Kind)
            'buffer.AppendLine("Kind = " & objRpt.Kind)
        Catch ex As Exception
        End Try
        Try
            Call GeEnumValue(type, "LastGetFormulaSyntax", objRpt.LastGetFormulaSyntax)
            'buffer.AppendLine("LastGetFormulaSyntax = " & objRpt.LastGetFormulaSyntax)
        Catch ex As Exception
        End Try
        Try
            buffer.AppendLine("Left = " & objRpt.Left)
        Catch ex As Exception
        End Try
        Try
            buffer.AppendLine("LeftIndent = " & objRpt.LeftIndent)
        Catch ex As Exception
        End Try
        Try
            buffer.AppendLine("LeftLineStyle = " & objRpt.LeftLineStyle)
        Catch ex As Exception
        End Try
        Try
            buffer.AppendLine("LineColor = " & objRpt.LineColor)
        Catch ex As Exception
        End Try
        Try
            buffer.AppendLine("LineStyle = " & objRpt.LineStyle)
        Catch ex As Exception
        End Try
        Try
            buffer.AppendLine("LineThickness = " & objRpt.LineThickness)
        Catch ex As Exception
        End Try
        Try
            buffer.AppendLine("LineSpacing = " & objRpt.LineSpacing)
        Catch ex As Exception
        End Try
        Try
            buffer.AppendLine("LineSpacingType = " & objRpt.LineSpacingType)
        Catch ex As Exception
        End Try
        Try
            buffer.AppendLine("MaxNumberOfLines = " & objRpt.MaxNumberOfLines)
        Catch ex As Exception
        End Try
        Try
            buffer.AppendLine("MinimumHeight = " & objRpt.MinimumHeight)
        Catch ex As Exception
        End Try
        Try
            buffer.AppendLine("MinuteSecondSeparator = " & objRpt.MinuteSecondSeparator)
        Catch ex As Exception
        End Try
        Try
            buffer.AppendLine("MinuteType = " & objRpt.MinuteType)
        Catch ex As Exception
        End Try
        Try
            buffer.AppendLine("MonthType = " & objRpt.MonthType)
        Catch ex As Exception
        End Try
        Try
            buffer.AppendLine("Name = " & objRpt.Name)
        Catch ex As Exception
        End Try
        Try
            buffer.AppendLine("NegativeType = " & objRpt.NegativeType)
        Catch ex As Exception
        End Try
        Try
            buffer.AppendLine("NewPageAfter = " & objRpt.NewPageAfter)
        Catch ex As Exception
        End Try
        Try
            Call GeEnumValue(type, "HierarchicalSummaryType", objRpt.HierarchicalSummaryType)
            'buffer.AppendLine("HierarchicalSummaryType = " & objRpt.HierarchicalSummaryType)
        Catch ex As Exception
        End Try
        Try
            buffer.AppendLine("NewPageBefore = " & objRpt.NewPageBefore)
        Catch ex As Exception
        End Try
        Try
            buffer.AppendLine("NextValue = " & objRpt.NextValue)
        Catch ex As Exception
        End Try
        Try
            buffer.AppendLine("Number = " & objRpt.Number)
        Catch ex As Exception
        End Try
        Try
            buffer.AppendLine("NumberOfBytes = " & objRpt.NumberOfBytes)
        Catch ex As Exception
        End Try
        Try
            Call GeEnumValue(type, "PaperOrientation", objRpt.PaperOrientation)
            'buffer.AppendLine("PaperOrientation = " & objRpt.PaperOrientation)
        Catch ex As Exception
        End Try
        Try
            Call GeEnumValue(type, "PaperSize", objRpt.PaperSize)
            'buffer.AppendLine("PaperSize = " & objRpt.PaperSize)
        Catch ex As Exception
        End Try
        Try
            Call GeEnumValue(type, "PaperSource", objRpt.PaperSource)
            'buffer.AppendLine("PaperSource = " & objRpt.PaperSource)
        Catch ex As Exception
        End Try
        Try
            Call GeEnumValue(type, "PrinterDuplex", objRpt.PrinterDuplex)
            'buffer.AppendLine("PrinterDuplex = " & objRpt.PrinterDuplex)
        Catch ex As Exception
        End Try
        Try
            buffer.AppendLine("PmString = " & objRpt.PmString)
        Catch ex As Exception
        End Try
        Try
            buffer.AppendLine("PreviousValue = " & objRpt.PreviousValue)
        Catch ex As Exception
        End Try
        Try
            buffer.AppendLine("PrintAtBottomOfPage = " & objRpt.PrintAtBottomOfPage)
        Catch ex As Exception
        End Try
        Try
            buffer.AppendLine("ResetPageNumberAfter = " & objRpt.PrintAtBottomOfPage)
        Catch ex As Exception
        End Try
        Try
            buffer.AppendLine("Right = " & objRpt.Right)
        Catch ex As Exception
        End Try
        Try
            buffer.AppendLine("RightIndent = " & objRpt.RightIndent)
        Catch ex As Exception
        End Try
        Try
            buffer.AppendLine("RightLineStyle = " & objRpt.RightLineStyle)
        Catch ex As Exception
        End Try
        Try
            buffer.AppendLine("RoundingType = " & objRpt.RoundingType)
        Catch ex As Exception
        End Try
        Try
            buffer.AppendLine("SecondType = " & objRpt.SecondType)
        Catch ex As Exception
        End Try
        Try
            buffer.AppendLine("SubreportName = " & objRpt.SubreportName)
        Catch ex As Exception
        End Try
        Try
            buffer.AppendLine("Suppress = " & objRpt.Suppress)
        Catch ex As Exception
        End Try
        Try
            buffer.AppendLine("SuppressIfDuplicated = " & objRpt.SuppressIfDuplicated)
        Catch ex As Exception
        End Try
        Try
            buffer.AppendLine("SuppressIfBlank = " & objRpt.SuppressIfBlank)
        Catch ex As Exception
        End Try
        Try
            buffer.AppendLine("SuppressIfZero = " & objRpt.SuppressIfZero)
        Catch ex As Exception
        End Try
        Try
            buffer.AppendLine("Text = " & objRpt.Text)
        Catch ex As Exception
        End Try
        Try
            buffer.AppendLine("TextColor = " & objRpt.TextColor)
        Catch ex As Exception
        End Try
        Try
            buffer.AppendLine("TextFormat = " & objRpt.TextFormat)
        Catch ex As Exception
        End Try
        Try
            buffer.AppendLine("TextRotationAngle = " & objRpt.TextRotationAngle)
        Catch ex As Exception
        End Try
        Try
            buffer.AppendLine("ThousandsSeparators = " & objRpt.ThousandsSeparators)
        Catch ex As Exception
        End Try
        Try
            buffer.AppendLine("ThousandSymbol = " & objRpt.ThousandSymbol)
        Catch ex As Exception
        End Try
        Try
            buffer.AppendLine("TimeBase = " & objRpt.TimeBase)
        Catch ex As Exception
        End Try
        Try
            buffer.AppendLine("Top = " & objRpt.Top)
        Catch ex As Exception
        End Try
        Try
            buffer.AppendLine("TopLineStyle = " & objRpt.TopLineStyle)
        Catch ex As Exception
        End Try
        Try
            buffer.AppendLine("UseIndexForSpeed = " & objRpt.UseIndexForSpeed)
        Catch ex As Exception
        End Try
        Try
            buffer.AppendLine("UseLeadingZero = " & objRpt.UseLeadingZero)
        Catch ex As Exception
        End Try
        Try
            buffer.AppendLine("UseOneSymbolPerPage = " & objRpt.UseOneSymbolPerPage)
        Catch ex As Exception
        End Try
        Try
            buffer.AppendLine("UseSystemDefaults = " & objRpt.UseSystemDefaults)
        Catch ex As Exception
        End Try
        Try
            buffer.AppendLine("UnderlaySection = " & objRpt.UnderlaySection)
        Catch ex As Exception
        End Try
        Try
            buffer.AppendLine("value = " & objRpt.value)
        Catch ex As Exception
        End Try
        Try
            Call GeEnumValue(type, "ValueType", objRpt.ValueType)
            'buffer.AppendLine("ValueType = " & objRpt.ValueType)
        Catch ex As Exception
        End Try
        Try
            buffer.AppendLine("Width = " & objRpt.Width)
        Catch ex As Exception
        End Try
        Try
            buffer.AppendLine("VerifyOnEveryPrint = " & objRpt.VerifyOnEveryPrint)
        Catch ex As Exception
        End Try
        Try
            buffer.AppendLine("YearType = " & objRpt.YearType)
        Catch ex As Exception
        End Try
        Try
            buffer.AppendLine("ZeroValueString = " & objRpt.ZeroValueString)
        Catch ex As Exception
        End Try
    End Sub

    Private Sub GeEnumValue(ByVal cls As String, ByVal prop As String, ByVal o As Object)
        Dim nm As String = String.Empty
        Select Case prop
            Case "ValueType"
                Select Case o
                    Case CRAXDDRT.CRFieldValueType.crBitmapField
                        nm = "crBitmapField"
                    Case CRAXDDRT.CRFieldValueType.crBlobField
                        nm = "crBlobField"
                    Case CRAXDDRT.CRFieldValueType.crBooleanField
                        nm = "crBooleanField"
                    Case CRAXDDRT.CRFieldValueType.crChartField
                        nm = "crChartField"
                    Case CRAXDDRT.CRFieldValueType.crCurrencyField
                        nm = "crCurrencyField"
                    Case CRAXDDRT.CRFieldValueType.crDateField
                        nm = "crDateField"
                    Case CRAXDDRT.CRFieldValueType.crDateTimeField
                        nm = "crDateTimeField"
                    Case CRAXDDRT.CRFieldValueType.crIconField
                        nm = "crIconField"
                    Case CRAXDDRT.CRFieldValueType.crInt16sField
                        nm = "crInt16sField"
                    Case CRAXDDRT.CRFieldValueType.crInt16uField
                        nm = "crInt16uField"
                    Case CRAXDDRT.CRFieldValueType.crInt32sField
                        nm = "crInt32sField"
                    Case CRAXDDRT.CRFieldValueType.crInt32uField
                        nm = "crInt32uField"
                    Case CRAXDDRT.CRFieldValueType.crInt8sField
                        nm = "crInt8sField"
                    Case CRAXDDRT.CRFieldValueType.crInt8uField
                        nm = "crInt8uField"
                    Case CRAXDDRT.CRFieldValueType.crNumberField
                        nm = "crNumberField"
                    Case CRAXDDRT.CRFieldValueType.crOleField
                        nm = "crOleField"
                    Case CRAXDDRT.CRFieldValueType.crPersistentMemoField
                        nm = "crPersistentMemoField"
                    Case CRAXDDRT.CRFieldValueType.crPictureField
                        nm = "crPictureField"
                    Case CRAXDDRT.CRFieldValueType.crStringField
                        nm = "crStringField"
                    Case CRAXDDRT.CRFieldValueType.crTimeField
                        nm = "crTimeField"
                    Case CRAXDDRT.CRFieldValueType.crTransientMemoField
                        nm = "crTransientMemoField"
                    Case CRAXDDRT.CRFieldValueType.crUnknownField
                        nm = "crUnknownField"
                End Select

            Case "HierarchicalSummaryType"
                Select Case o
                    Case CRAXDDRT.CRHierarchicalSummaryType.crHierarchicalSummaryNone
                        nm = "crHierarchicalSummaryNone"
                    Case CRAXDDRT.CRHierarchicalSummaryType.crSummaryAcrossHierarchy
                        nm = "crSummaryAcrossHierarchy"
                End Select

            Case "ConvertDateTimeType"
                Select Case o
                    Case CRAXDDRT.CRConvertDateTimeType.crConvertDateTimeToDate
                        nm = "crConvertDateTimeToDate"
                    Case CRAXDDRT.CRConvertDateTimeType.crConvertDateTimeToString
                        nm = "crConvertDateTimeToString"
                    Case CRAXDDRT.CRConvertDateTimeType.crKeepDateTimeType
                        nm = "crKeepDateTimeType"
                End Select

            Case "FieldMappingType"
                Select Case o
                    Case CRAXDDRT.CRFieldMappingType.crAutoFieldMapping
                        nm = "crAutoFieldMapping"
                    Case CRAXDDRT.CRFieldMappingType.crEventFieldMapping
                        nm = "crEventFieldMapping"
                    Case CRAXDDRT.CRFieldMappingType.crPromptFieldMapping
                        nm = "crPromptFieldMapping"
                End Select

            Case "FormulaSyntax", "LastGetFormulaSyntax"
                Select Case o
                    Case CRAXDDRT.CRFormulaSyntax.crBasicSyntaxFormula
                        nm = "crBasicSyntaxFormula"
                    Case CRAXDDRT.CRFormulaSyntax.crCrystalSyntaxFormula
                        nm = "crCrystalSyntaxFormula"
                End Select

            Case "Kind"
                Select Case cls
                    Case "ReportClass"

                        Select Case o
                            Case CRAXDDRT.CRReportKind.crColumnarReport
                                nm = "crColumnarReport"
                            Case CRAXDDRT.CRReportKind.crLabelReport
                                nm = "crLabelReport"
                            Case CRAXDDRT.CRReportKind.crMulColumnReport
                                nm = "crMulColumnReport"
                        End Select

                    Case "ReportObjectClass"
                        Select Case o
                            Case CRAXDDRT.CRObjectKind.crBlobFieldObject
                                nm = "crBlobFieldObject"
                            Case CRAXDDRT.CRObjectKind.crBoxObject
                                nm = "crBoxObject"
                            Case CRAXDDRT.CRObjectKind.crCrossTabObject
                                nm = "crCrossTabObject"
                            Case CRAXDDRT.CRObjectKind.crFieldObject
                                nm = "crFieldObject"
                            Case CRAXDDRT.CRObjectKind.crGraphObject
                                nm = "crGraphObject"
                            Case CRAXDDRT.CRObjectKind.crLineObject
                                nm = "crLineObject"
                            Case CRAXDDRT.CRObjectKind.crMapObject
                                nm = "crMapObject"
                            Case CRAXDDRT.CRObjectKind.crOlapGridObject
                                nm = "crOlapGridObject"
                            Case CRAXDDRT.CRObjectKind.crOleObject
                                nm = "crOleObject"
                            Case CRAXDDRT.CRObjectKind.crSubreportObject
                                nm = "crSubreportObject"
                            Case CRAXDDRT.CRObjectKind.crTextObject
                                nm = "crTextObject"
                        End Select

                    Case "GroupNameFieldDefinition", "RunningTotalFieldDefinition", "FormulaFieldDefinition"
                        Select Case o
                            Case CRAXDDRT.CRFieldKind.crDatabaseField
                                nm = "crDatabaseField"
                            Case CRAXDDRT.CRFieldKind.crFormulaField
                                nm = "crFormulaField"
                            Case CRAXDDRT.CRFieldKind.crGroupNameField
                                nm = "crGroupNameField"
                            Case CRAXDDRT.CRFieldKind.crParameterField
                                nm = "crParameterField"
                            Case CRAXDDRT.CRFieldKind.crRunningTotalField
                                nm = "crRunningTotalField"
                            Case CRAXDDRT.CRFieldKind.crSpecialVarField
                                nm = "crSpecialVarField"
                            Case CRAXDDRT.CRFieldKind.crSQLExpressionField
                                nm = "crSQLExpressionField"
                            Case CRAXDDRT.CRFieldKind.crSummaryField
                                nm = "crSQLExpressionField"
                        End Select
                End Select

            Case "PaperOrientation"
                Select Case o
                    Case CRAXDDRT.CRPaperOrientation.crDefaultPaperOrientation
                        nm = "crDefaultPaperOrientation"
                    Case CRAXDDRT.CRPaperOrientation.crLandscape
                        nm = "crLandscape"
                    Case CRAXDDRT.CRPaperOrientation.crPortrait
                        nm = "crPortrait"
                End Select

            Case "PaperSize"
                Select Case o
                    Case CRAXDDRT.CRPaperSize.crDefaultPaperSize
                        nm = "crDefaultPaperSize"
                    Case CRAXDDRT.CRPaperSize.crPaper10x14
                        nm = "crPaper10x14"
                    Case CRAXDDRT.CRPaperSize.crPaper11x17
                        nm = "crPaper11x17"
                    Case CRAXDDRT.CRPaperSize.crPaperA3
                        nm = "crPaperA3"
                    Case CRAXDDRT.CRPaperSize.crPaperA4
                        nm = "crPaperA4"
                    Case CRAXDDRT.CRPaperSize.crPaperA4Small
                        nm = "crPaperA4Small"
                    Case CRAXDDRT.CRPaperSize.crPaperA5
                        nm = "crPaperA5"
                    Case CRAXDDRT.CRPaperSize.crPaperB4
                        nm = "crPaperB4"
                    Case CRAXDDRT.CRPaperSize.crPaperB5
                        nm = "crPaperB5"
                    Case CRAXDDRT.CRPaperSize.crPaperCsheet
                        nm = "crPaperCsheet"
                    Case CRAXDDRT.CRPaperSize.crPaperDsheet
                        nm = "crPaperDsheet"
                    Case CRAXDDRT.CRPaperSize.crPaperEnvelope10
                        nm = "crPaperEnvelope10"
                    Case CRAXDDRT.CRPaperSize.crPaperEnvelope11
                        nm = "crPaperEnvelope11"
                    Case CRAXDDRT.CRPaperSize.crPaperEnvelope12
                        nm = "crPaperEnvelope12"
                    Case CRAXDDRT.CRPaperSize.crPaperEnvelope14
                        nm = "crPaperEnvelope14"
                    Case CRAXDDRT.CRPaperSize.crPaperEnvelope9
                        nm = "crPaperEnvelope9"
                    Case CRAXDDRT.CRPaperSize.crPaperEnvelopeB4
                        nm = "crPaperEnvelopeB4"
                    Case CRAXDDRT.CRPaperSize.crPaperEnvelopeB5
                        nm = "crPaperEnvelopeB5"
                    Case CRAXDDRT.CRPaperSize.crPaperEnvelopeB6
                        nm = "crPaperEnvelopeB6"
                    Case CRAXDDRT.CRPaperSize.crPaperEnvelopeC3
                        nm = "crPaperEnvelopeC3"
                    Case CRAXDDRT.CRPaperSize.crPaperEnvelopeC4
                        nm = "crPaperEnvelopeC4"
                    Case CRAXDDRT.CRPaperSize.crPaperEnvelopeC5
                        nm = "crPaperEnvelopeC5"
                    Case CRAXDDRT.CRPaperSize.crPaperEnvelopeC6
                        nm = "crPaperEnvelopeC6"
                    Case CRAXDDRT.CRPaperSize.crPaperEnvelopeC65
                        nm = "crPaperEnvelopeC65"
                    Case CRAXDDRT.CRPaperSize.crPaperEnvelopeDL
                        nm = "crPaperEnvelopeDL"
                    Case CRAXDDRT.CRPaperSize.crPaperEnvelopeItaly
                        nm = "crPaperEnvelopeItaly"
                    Case CRAXDDRT.CRPaperSize.crPaperEnvelopeMonarch
                        nm = "crPaperEnvelopeMonarch"
                    Case CRAXDDRT.CRPaperSize.crPaperEnvelopePersonal
                        nm = "crPaperEnvelopePersonal"
                    Case CRAXDDRT.CRPaperSize.crPaperEsheet
                        nm = "crPaperEsheet"
                    Case CRAXDDRT.CRPaperSize.crPaperExecutive
                        nm = "crPaperExecutive"
                    Case CRAXDDRT.CRPaperSize.crPaperFanfoldLegalGerman
                        nm = "crPaperFanfoldLegalGerman"
                    Case CRAXDDRT.CRPaperSize.crPaperFanfoldStdGerman
                        nm = "crPaperFanfoldStdGerman"
                    Case CRAXDDRT.CRPaperSize.crPaperFanfoldUS
                        nm = "crPaperFanfoldUS"
                    Case CRAXDDRT.CRPaperSize.crPaperFolio
                        nm = "crPaperFolio"
                    Case CRAXDDRT.CRPaperSize.crPaperLedger
                        nm = "crPaperLedger"
                    Case CRAXDDRT.CRPaperSize.crPaperLegal
                        nm = "crPaperLegal"
                    Case CRAXDDRT.CRPaperSize.crPaperLetter
                        nm = "crPaperLetter"
                    Case CRAXDDRT.CRPaperSize.crPaperLetterSmall
                        nm = "crPaperLetterSmall"
                    Case CRAXDDRT.CRPaperSize.crPaperNote
                        nm = "crPaperNote"
                    Case CRAXDDRT.CRPaperSize.crPaperQuarto
                        nm = "crPaperQuarto"
                    Case CRAXDDRT.CRPaperSize.crPaperStatement
                        nm = "crPaperStatement"
                    Case CRAXDDRT.CRPaperSize.crPaperTabloid
                        nm = "crPaperTabloid"
                    Case CRAXDDRT.CRPaperSize.crPaperUser
                        nm = "crPaperUser"
                End Select

            Case "PaperSource"
                Select Case o
                    Case CRAXDDRT.CRPaperSource.crPRBinAuto
                        nm = "crPRBinAuto"
                    Case CRAXDDRT.CRPaperSource.crPRBinCassette
                        nm = "crPRBinCassette"
                    Case CRAXDDRT.CRPaperSource.crPRBinEnvelope
                        nm = "crPRBinEnvelope"
                    Case CRAXDDRT.CRPaperSource.crPRBinEnvManual
                        nm = "crPRBinEnvManual"
                    Case CRAXDDRT.CRPaperSource.crPRBinFormSource
                        nm = "crPRBinFormSource"
                    Case CRAXDDRT.CRPaperSource.crPRBinLargeCapacity
                        nm = "crPRBinLargeCapacity"
                    Case CRAXDDRT.CRPaperSource.crPRBinLargeFmt
                        nm = "crPRBinLargeFmt"
                    Case CRAXDDRT.CRPaperSource.crPRBinLower
                        nm = "crPRBinLower"
                    Case CRAXDDRT.CRPaperSource.crPRBinManual
                        nm = "crPRBinManual"
                    Case CRAXDDRT.CRPaperSource.crPRBinMiddle
                        nm = "crPRBinMiddle"
                    Case CRAXDDRT.CRPaperSource.crPRBinSmallFmt
                        nm = "crPRBinSmallFmt"
                    Case CRAXDDRT.CRPaperSource.crPRBinTractor
                        nm = "crPRBinTractor"
                    Case CRAXDDRT.CRPaperSource.crPRBinUpper
                        nm = "crPRBinUpper"
                End Select

            Case "PrinterDuplex"
                Select Case o
                    Case CRAXDDRT.CRPrinterDuplexType.crPRDPDefault
                        nm = "crPRDPDefault"
                    Case CRAXDDRT.CRPrinterDuplexType.crPRDPHorizontal
                        nm = "crPRDPHorizontal"
                    Case CRAXDDRT.CRPrinterDuplexType.crPRDPSimplex
                        nm = "crPRDPSimplex"
                    Case CRAXDDRT.CRPrinterDuplexType.crPRDPVertical
                        nm = "crPRDPVertical"
                End Select

        End Select

        If nm.Length > 0 Then
            buffer.AppendLine(prop & " = " & o & " " & nm)
        Else
            buffer.AppendLine(prop & " = " & o)
        End If

    End Sub

    Private Sub exp(ByVal Rpt As CRAXDRT.Report, ByVal Parent As String)


        buffer.AppendLine("=== " & Parent & "Report === {")

        props(Rpt, "ReportClass")

        buffer.AppendLine("")

        For Each Sec As CRAXDRT.Section In Rpt.Sections



            'SectionName

            buffer.AppendLine("=== " & Parent & Sec.Name & " === {")

            props(Sec, "Section")


            'CommonFormula,ColorFormula

            DspFormulaCode(Sec, m_objFomuraSec)



            For Each objRpt As Object In Sec.ReportObjects



                'GetFieldType

                Dim strFiledType As String = String.Empty

                Select Case CType(objRpt.kind, Integer)

                    Case 1 : strFiledType = "crFieldObject"

                    Case 2 : strFiledType = "crTextObject"

                    Case 3 : strFiledType = "crLineObject"

                    Case 4 : strFiledType = "crBoxObject"

                        'Case 3 : strFiledType = "" 'crLineObject

                        'Case 4 : strFiledType = "" 'crBoxObject

                    Case 5 : strFiledType = "crSubreportObject"

                    Case 6 : strFiledType = "crOleObject"

                    Case 7 : strFiledType = "crGraphObject"

                    Case 8 : strFiledType = "crCrossTabObject"

                    Case 9 : strFiledType = "crBlobFieldObject"

                    Case 10 : strFiledType = "crMapObject"

                    Case 11 : strFiledType = "crOlapGridObject"

                End Select



                If strFiledType <> "" Then

                    'ReportObjectName

                    buffer.AppendLine("=== " & Parent & objRpt.Name & " === {")


                    props(objRpt, "ReportObjectClass")

                    'CommonFormula

                    DspFormulaCode(objRpt, m_objFomuraCmn)



                    'BorderLineFormula

                    DspFormulaCode(objRpt, m_objFomuraBdr)



                    'NumericFormula

                    DspFormulaCode(objRpt, m_objFomuraNum)



                    'DateFormula

                    DspFormulaCode(objRpt, m_objFomuraYmd)



                    'TimeFormula

                    DspFormulaCode(objRpt, m_objFomuraTme)



                    'FontFormula

                    DspFormulaCode(objRpt, m_objFomuraFnt)



                    'IndentFormula

                    DspFormulaCode(objRpt, m_objFomuraInd)



                    'HyperLinkFormula

                    DspFormulaCode(objRpt, m_objFomuraLnk)

                End If





                If CType(objRpt.kind, Integer) = 8 Then

                    'crCrossTabObject

                    Dim crs As CRAXDDRT.CrossTabObject = CType(objRpt, CRAXDDRT.CrossTabObject)



                    'CrossTabObjectName

                    buffer.AppendLine("=== " & Parent & crs.Name & " === {")



                    Try

                        For Each objCold As CRAXDDRT.CrossTabGroup In crs.ColumnGroups

                            'ColumunObject within a CrossTab

                            buffer.AppendLine(Parent & objCold.Field.Name)



                            'xxxxFormula

                            'xxxxxxxxxxxxxxxxxxxxxxxx

                        Next



                        For Each objRowd As CRAXDDRT.CrossTabGroup In crs.RowGroups

                            'RowObject within a CrossTab

                            buffer.AppendLine(Parent & objRowd.Field.Name)



                            'xxxxFormula

                            'xxxxxxxxxxxxxxxxxxxxxxxx

                        Next



                    Catch ex As Exception

                        'buffer.AppendLine(ex.Message)
                        buffer.AppendLine("Not get ColumnGroups")

                    End Try

                    buffer.AppendLine("}")

                End If



                If CType(objRpt.kind, Integer) = 5 Then

                    'crSubreportObjectの場合

                    Dim subr As CRAXDDRT.SubreportObject = CType(objRpt, CRAXDDRT.SubreportObject)



                    'SubreportObjectName

                    buffer.AppendLine(Parent & subr.SubreportName & " {")



                    'Recursive processing

                    exp(subr.OpenSubreport, Parent & subr.SubreportName & ".")



                    buffer.AppendLine("}")
                    buffer.AppendLine("")

                End If

                buffer.AppendLine("}")
                buffer.AppendLine("")

            Next

            buffer.AppendLine("}")
            buffer.AppendLine("")

        Next

        buffer.AppendLine("}")
        buffer.AppendLine("")

        'FormulaFields

        buffer.AppendLine("=== " & Parent & "FormulaFields === {")

        Try

            For Each objFFd As CRAXDDRT.FormulaFieldDefinition In Rpt.FormulaFields

                'fieldName

                buffer.AppendLine(Parent & objFFd.FormulaFieldName & " {")

                props(objFFd, "FormulaFieldDefinition")


                'CommonFormula

                DspFormulaCode(objFFd, m_objFomuraCmn)



                'BorderLineFormula

                DspFormulaCode(objFFd, m_objFomuraBdr)



                'NumericFormula

                DspFormulaCode(objFFd, m_objFomuraNum)



                'DateFormula

                DspFormulaCode(objFFd, m_objFomuraYmd)



                'TimeFormula

                DspFormulaCode(objFFd, m_objFomuraTme)



                'FontFormula

                DspFormulaCode(objFFd, m_objFomuraFnt)



                'IndentFormula

                DspFormulaCode(objFFd, m_objFomuraInd)



                'HyperLinkFormula

                DspFormulaCode(objFFd, m_objFomuraLnk)


                buffer.AppendLine("}")
                buffer.AppendLine("")

            Next



        Catch ex As Exception

            buffer.AppendLine("Not get FormulaFields")
            'buffer.AppendLine(ex.Message)

        End Try

        buffer.AppendLine("}")
        buffer.AppendLine("")



        'GroupNameFields

        buffer.AppendLine("=== " & Parent & "GroupNameFields === {")

        Try

            For Each objGFd As CRAXDDRT.GroupNameFieldDefinition In Rpt.GroupNameFields

                'fieldName

                buffer.AppendLine(Parent & objGFd.GroupNameFieldName & " GroupNameConditionFormula {")

                props(objGFd, "GroupNameFieldDefinition")

                'GroupNameConditionFormula

                buffer.AppendLine(objGFd.GroupNameConditionFormula)

                buffer.AppendLine("}")
                buffer.AppendLine("")

            Next



        Catch ex As Exception

            buffer.AppendLine("Not get GroupNameFields")
            'buffer.AppendLine(ex.Message)

        End Try

        buffer.AppendLine("}")
        buffer.AppendLine("")



        'RunningTotalFields

        buffer.AppendLine("=== " & Parent & "RunningTotalFields === {")

        Try

            For Each objRFd As CRAXDDRT.RunningTotalFieldDefinition In Rpt.RunningTotalFields

                'fieldName

                buffer.AppendLine(Parent & objRFd.RunningTotalFieldName & " EvaluateConditionFormula {")

                props(objRFd, "RunningTotalFieldDefinition")

                Try
                    'EvaluateConditionFormula
                    buffer.AppendLine(objRFd.EvaluateConditionFormula)
                Catch ex As Exception
                    buffer.AppendLine("Not get EvaluateConditionFormula")
                End Try

                buffer.AppendLine("")

                'ResetConditionFormula
                buffer.AppendLine(Parent & objRFd.RunningTotalFieldName & " ResetConditionFormula {")

                Try
                    buffer.AppendLine(objRFd.ResetConditionFormula)
                Catch ex As Exception
                    buffer.AppendLine("Not get ResetConditionFormula")
                End Try
                buffer.AppendLine("}")
                buffer.AppendLine("")

                buffer.AppendLine("}")
                buffer.AppendLine("")

            Next

        Catch ex As Exception

            buffer.AppendLine("Not get RunningTotalFields")
            'buffer.AppendLine(ex.Message)

        End Try

        buffer.AppendLine("}")
        buffer.AppendLine("")



        'SQLExpressionFields

        buffer.AppendLine("=== " & Parent & "SQLExpressionFields === {")

        Try

            For Each objSQLd As CRAXDDRT.SQLExpressionFieldDefinition In Rpt.SQLExpressionFields

                'fieldName

                buffer.AppendLine(Parent & objSQLd.SQLExpressionFieldName & " SQL {")

                'SQL

                buffer.AppendLine(objSQLd.Text)

                buffer.AppendLine("}")
                buffer.AppendLine("")

            Next



        Catch ex As Exception

            buffer.AppendLine("Not get SQLExpressionFields")
            'buffer.AppendLine(ex.Message)

        End Try

        buffer.AppendLine("}")
        buffer.AppendLine("")



        'SelectionFormula

        buffer.AppendLine(Parent & "SelectionFormula > RecordSelectionFormula {")

        'RecordSelectionFormula

        buffer.AppendLine(Rpt.RecordSelectionFormula)

        buffer.AppendLine("}")
        buffer.AppendLine("")



        buffer.AppendLine(Parent & "SelectionFormula > GroupSelectionFormula {")

        'GroupSelectionFormula

        buffer.AppendLine(Rpt.GroupSelectionFormula)

        buffer.AppendLine("}")
        buffer.AppendLine("")



    End Sub



    Private Sub DspFormulaCode(ByVal ObjRpt As Object, ByVal enmArray As Object(,))



        Dim strFormulaCode As String = String.Empty

        For intI As Integer = 0 To enmArray.GetUpperBound(0)

            Try

                'xxxxxConditionFormula

                buffer.AppendLine(enmArray(intI, 0).ToString & " (" & enmArray(intI, 1).ToString & ") {")

                'xxxxxConditionFormula

                buffer.AppendLine(ObjRpt.ConditionFormula(CType(enmArray(intI, 1), Integer)))




            Catch ex As Exception
                buffer.AppendLine("Not get ConditionFormula.")
                'buffer.AppendLine(ex.Message)
                Exit Try
            Finally
                buffer.AppendLine("}")
                buffer.AppendLine("")
            End Try

        Next



    End Sub


End Module
