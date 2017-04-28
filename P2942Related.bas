Attribute VB_Name = "P2942Related"
Option Explicit

Public Sub OutputNames()

    Dim rngOut As Range
    Set rngOut = Range("B3")
    
    Dim namedRange As name
    For Each namedRange In ActiveWorkbook.Names
        
        Dim hasForm As Boolean
        hasForm = False
        
        On Error Resume Next
        hasForm = namedRange.RefersToRange.HasFormula
        On Error GoTo 0
        
        If namedRange.Visible And namedRange.name <> "SELF" Then
            
            rngOut = namedRange.name
            rngOut.Offset(, 1) = namedRange.Comment
            
            Set rngOut = rngOut.Offset(1)
            
        End If
    Next
End Sub

Public Sub GetFormulasForNamedRanges()

    Dim n As Integer
    n = FreeFile()
    Open "C:\Documents\TDA\2942\Mass balance\formulas.ini" For Output As #n

    Dim namedRange As name
    For Each namedRange In ActiveWorkbook.Names
        
        Dim hasForm As Boolean
        hasForm = False
        
        On Error Resume Next
        hasForm = namedRange.RefersToRange.HasFormula
        On Error GoTo 0
        
        If namedRange.Visible And namedRange.name <> "SELF" And InStr(namedRange.name, "IGNORE") = 0 Then
            
            If hasForm Then
                'this allows for self reference
                Dim strFormula As String
                strFormula = namedRange.RefersToRange.Formula
            
                Range("J90").Formula = strFormula
                On Error Resume Next
                strFormula = Range("K90").Value
                
                If Err.Number <> 0 Then
                    Debug.Print "Error"
                End If
                
                On Error GoTo 0
            
                strFormula = Replace(strFormula, "SELF", namedRange.name)
                
                If namedRange.name = "FT_601_slpm" Then
                    strFormula = strFormula & "*SIGN(FT_601_dp)"
                End If
                
                Dim deadband As Double
                deadband = 1000000
                
                If InStr(namedRange.name, "capture") Then
                    deadband = 0.1
                End If
                
                
                strFormula = Replace(strFormula, "*", " * ")
                strFormula = Replace(strFormula, "-", " - ")
                strFormula = Replace(strFormula, "+", " + ")
                strFormula = Replace(strFormula, "/", " / ")
                strFormula = Replace(strFormula, "ABS(", "abs(")
                strFormula = Replace(strFormula, "_", "-")
                
                Print #n, ";" & namedRange.name & " = " & namedRange.Comment
                
                Print #n, "[" & Replace(namedRange.name, "_", "-") & "]"
                
                Print #n, "type = float"
                Print #n, "math = " & strFormula
                Print #n, "daq_ch ="
                Print #n, "daq_io ="
                Print #n, "daq_cal ="
                Print #n, "filter ="
                Print #n, "mb_ip ="
                Print #n, "mb_ch ="
                Print #n, "mb_type ="
                Print #n, "mb_order ="
                Print #n, "mb_io ="
                Print #n, "log_series ="
                Print #n, "daq_type ="
                Print #n, "log_type = float"
                Print #n, "deadband = " & deadband
            
                Print #n, ""
            End If
        End If
    Next
    
    Close #n
    

End Sub

Sub FormatChartsWithAxis()

    Dim chtObj As ChartObject
    For Each chtObj In ActiveSheet.ChartObjects
        
        Dim cht As Chart
        Dim ax As Axis
        
        Set cht = chtObj.Chart
        Set ax = cht.Axes(xlValue)
        
        ax.MinimumScale = 70
        ax.MaximumScale = 270
        
        cht.HasTitle = True
        
        Set ax = cht.Axes(xlCategory)
        ax.TickLabels.NumberFormat = "m/d HH:mm"
    
    Next

End Sub

Sub MakePdf2()

    Dim path As String
    path = "C:\Documents\TDA\2942\Data analysis of files\Bed Cycle analysis\PDF\"
    
    Dim index As Variant
    index = "recent"
 
        
    Sheet1.ExportAsFixedFormat xlTypePDF, path & index & ".pdf"


End Sub

Public Sub Chart_ApplyViridis()
    '---------------------------------------------------------------------------------------
    ' Procedure : Chart_ApplyTrendColors
    ' Author    : @byronwall
    ' Date      : 2015 07 24
    ' Purpose   : Applies the predetermined chart colors to each series
    '---------------------------------------------------------------------------------------
    '
    Dim targetObject As ChartObject
    For Each targetObject In Chart_GetObjectsFromObject(Selection)

        Dim targetSeries As series
        For Each targetSeries In targetObject.Chart.SeriesCollection

            Dim butlSeries As New bUTLChartSeries
            butlSeries.UpdateFromChartSeries targetSeries

            targetSeries.MarkerForegroundColorIndex = xlColorIndexNone
            targetSeries.MarkerStyle = xlMarkerStyleNone
            targetSeries.Format.Line.ForeColor.RGB = Chart_GetViridis(butlSeries.SeriesNumber, targetObject.Chart.SeriesCollection.Count)
            targetSeries.Format.Line.Weight = 1.5

        Next targetSeries
    Next targetObject
End Sub

Private Function Chart_GetViridis(index As Variant, totalCount As Variant)

    'return an array that is the size of the count
    Dim colors(1 To 256) As Variant
    
    colors(1) = RGB(68, 1, 84)
    colors(2) = RGB(68, 2, 86)
    colors(3) = RGB(69, 4, 87)
    colors(4) = RGB(69, 5, 89)
    colors(5) = RGB(70, 7, 90)
    colors(6) = RGB(70, 8, 92)
    colors(7) = RGB(70, 10, 93)
    colors(8) = RGB(70, 11, 94)
    colors(9) = RGB(71, 13, 96)
    colors(10) = RGB(71, 14, 97)
    colors(11) = RGB(71, 16, 99)
    colors(12) = RGB(71, 17, 100)
    colors(13) = RGB(71, 19, 101)
    colors(14) = RGB(72, 20, 103)
    colors(15) = RGB(72, 22, 104)
    colors(16) = RGB(72, 23, 105)
    colors(17) = RGB(72, 24, 106)
    colors(18) = RGB(72, 26, 108)
    colors(19) = RGB(72, 27, 109)
    colors(20) = RGB(72, 28, 110)
    colors(21) = RGB(72, 29, 111)
    colors(22) = RGB(72, 31, 112)
    colors(23) = RGB(72, 32, 113)
    colors(24) = RGB(72, 33, 115)
    colors(25) = RGB(72, 35, 116)
    colors(26) = RGB(72, 36, 117)
    colors(27) = RGB(72, 37, 118)
    colors(28) = RGB(72, 38, 119)
    colors(29) = RGB(72, 40, 120)
    colors(30) = RGB(72, 41, 121)
    colors(31) = RGB(71, 42, 122)
    colors(32) = RGB(71, 44, 122)
    colors(33) = RGB(71, 45, 123)
    colors(34) = RGB(71, 46, 124)
    colors(35) = RGB(71, 47, 125)
    colors(36) = RGB(70, 48, 126)
    colors(37) = RGB(70, 50, 126)
    colors(38) = RGB(70, 51, 127)
    colors(39) = RGB(70, 52, 128)
    colors(40) = RGB(69, 53, 129)
    colors(41) = RGB(69, 55, 129)
    colors(42) = RGB(69, 56, 130)
    colors(43) = RGB(68, 57, 131)
    colors(44) = RGB(68, 58, 131)
    colors(45) = RGB(68, 59, 132)
    colors(46) = RGB(67, 61, 132)
    colors(47) = RGB(67, 62, 133)
    colors(48) = RGB(66, 63, 133)
    colors(49) = RGB(66, 64, 134)
    colors(50) = RGB(66, 65, 134)
    colors(51) = RGB(65, 66, 135)
    colors(52) = RGB(65, 68, 135)
    colors(53) = RGB(64, 69, 136)
    colors(54) = RGB(64, 70, 136)
    colors(55) = RGB(63, 71, 136)
    colors(56) = RGB(63, 72, 137)
    colors(57) = RGB(62, 73, 137)
    colors(58) = RGB(62, 74, 137)
    colors(59) = RGB(62, 76, 138)
    colors(60) = RGB(61, 77, 138)
    colors(61) = RGB(61, 78, 138)
    colors(62) = RGB(60, 79, 138)
    colors(63) = RGB(60, 80, 139)
    colors(64) = RGB(59, 81, 139)
    colors(65) = RGB(59, 82, 139)
    colors(66) = RGB(58, 83, 139)
    colors(67) = RGB(58, 84, 140)
    colors(68) = RGB(57, 85, 140)
    colors(69) = RGB(57, 86, 140)
    colors(70) = RGB(56, 88, 140)
    colors(71) = RGB(56, 89, 140)
    colors(72) = RGB(55, 90, 140)
    colors(73) = RGB(55, 91, 141)
    colors(74) = RGB(54, 92, 141)
    colors(75) = RGB(54, 93, 141)
    colors(76) = RGB(53, 94, 141)
    colors(77) = RGB(53, 95, 141)
    colors(78) = RGB(52, 96, 141)
    colors(79) = RGB(52, 97, 141)
    colors(80) = RGB(51, 98, 141)
    colors(81) = RGB(51, 99, 141)
    colors(82) = RGB(50, 100, 142)
    colors(83) = RGB(50, 101, 142)
    colors(84) = RGB(49, 102, 142)
    colors(85) = RGB(49, 103, 142)
    colors(86) = RGB(49, 104, 142)
    colors(87) = RGB(48, 105, 142)
    colors(88) = RGB(48, 106, 142)
    colors(89) = RGB(47, 107, 142)
    colors(90) = RGB(47, 108, 142)
    colors(91) = RGB(46, 109, 142)
    colors(92) = RGB(46, 110, 142)
    colors(93) = RGB(46, 111, 142)
    colors(94) = RGB(45, 112, 142)
    colors(95) = RGB(45, 113, 142)
    colors(96) = RGB(44, 113, 142)
    colors(97) = RGB(44, 114, 142)
    colors(98) = RGB(44, 115, 142)
    colors(99) = RGB(43, 116, 142)
    colors(100) = RGB(43, 117, 142)
    colors(101) = RGB(42, 118, 142)
    colors(102) = RGB(42, 119, 142)
    colors(103) = RGB(42, 120, 142)
    colors(104) = RGB(41, 121, 142)
    colors(105) = RGB(41, 122, 142)
    colors(106) = RGB(41, 123, 142)
    colors(107) = RGB(40, 124, 142)
    colors(108) = RGB(40, 125, 142)
    colors(109) = RGB(39, 126, 142)
    colors(110) = RGB(39, 127, 142)
    colors(111) = RGB(39, 128, 142)
    colors(112) = RGB(38, 129, 142)
    colors(113) = RGB(38, 130, 142)
    colors(114) = RGB(38, 130, 142)
    colors(115) = RGB(37, 131, 142)
    colors(116) = RGB(37, 132, 142)
    colors(117) = RGB(37, 133, 142)
    colors(118) = RGB(36, 134, 142)
    colors(119) = RGB(36, 135, 142)
    colors(120) = RGB(35, 136, 142)
    colors(121) = RGB(35, 137, 142)
    colors(122) = RGB(35, 138, 141)
    colors(123) = RGB(34, 139, 141)
    colors(124) = RGB(34, 140, 141)
    colors(125) = RGB(34, 141, 141)
    colors(126) = RGB(33, 142, 141)
    colors(127) = RGB(33, 143, 141)
    colors(128) = RGB(33, 144, 141)
    colors(129) = RGB(33, 145, 140)
    colors(130) = RGB(32, 146, 140)
    colors(131) = RGB(32, 146, 140)
    colors(132) = RGB(32, 147, 140)
    colors(133) = RGB(31, 148, 140)
    colors(134) = RGB(31, 149, 139)
    colors(135) = RGB(31, 150, 139)
    colors(136) = RGB(31, 151, 139)
    colors(137) = RGB(31, 152, 139)
    colors(138) = RGB(31, 153, 138)
    colors(139) = RGB(31, 154, 138)
    colors(140) = RGB(30, 155, 138)
    colors(141) = RGB(30, 156, 137)
    colors(142) = RGB(30, 157, 137)
    colors(143) = RGB(31, 158, 137)
    colors(144) = RGB(31, 159, 136)
    colors(145) = RGB(31, 160, 136)
    colors(146) = RGB(31, 161, 136)
    colors(147) = RGB(31, 161, 135)
    colors(148) = RGB(31, 162, 135)
    colors(149) = RGB(32, 163, 134)
    colors(150) = RGB(32, 164, 134)
    colors(151) = RGB(33, 165, 133)
    colors(152) = RGB(33, 166, 133)
    colors(153) = RGB(34, 167, 133)
    colors(154) = RGB(34, 168, 132)
    colors(155) = RGB(35, 169, 131)
    colors(156) = RGB(36, 170, 131)
    colors(157) = RGB(37, 171, 130)
    colors(158) = RGB(37, 172, 130)
    colors(159) = RGB(38, 173, 129)
    colors(160) = RGB(39, 173, 129)
    colors(161) = RGB(40, 174, 128)
    colors(162) = RGB(41, 175, 127)
    colors(163) = RGB(42, 176, 127)
    colors(164) = RGB(44, 177, 126)
    colors(165) = RGB(45, 178, 125)
    colors(166) = RGB(46, 179, 124)
    colors(167) = RGB(47, 180, 124)
    colors(168) = RGB(49, 181, 123)
    colors(169) = RGB(50, 182, 122)
    colors(170) = RGB(52, 182, 121)
    colors(171) = RGB(53, 183, 121)
    colors(172) = RGB(55, 184, 120)
    colors(173) = RGB(56, 185, 119)
    colors(174) = RGB(58, 186, 118)
    colors(175) = RGB(59, 187, 117)
    colors(176) = RGB(61, 188, 116)
    colors(177) = RGB(63, 188, 115)
    colors(178) = RGB(64, 189, 114)
    colors(179) = RGB(66, 190, 113)
    colors(180) = RGB(68, 191, 112)
    colors(181) = RGB(70, 192, 111)
    colors(182) = RGB(72, 193, 110)
    colors(183) = RGB(74, 193, 109)
    colors(184) = RGB(76, 194, 108)
    colors(185) = RGB(78, 195, 107)
    colors(186) = RGB(80, 196, 106)
    colors(187) = RGB(82, 197, 105)
    colors(188) = RGB(84, 197, 104)
    colors(189) = RGB(86, 198, 103)
    colors(190) = RGB(88, 199, 101)
    colors(191) = RGB(90, 200, 100)
    colors(192) = RGB(92, 200, 99)
    colors(193) = RGB(94, 201, 98)
    colors(194) = RGB(96, 202, 96)
    colors(195) = RGB(99, 203, 95)
    colors(196) = RGB(101, 203, 94)
    colors(197) = RGB(103, 204, 92)
    colors(198) = RGB(105, 205, 91)
    colors(199) = RGB(108, 205, 90)
    colors(200) = RGB(110, 206, 88)
    colors(201) = RGB(112, 207, 87)
    colors(202) = RGB(115, 208, 86)
    colors(203) = RGB(117, 208, 84)
    colors(204) = RGB(119, 209, 83)
    colors(205) = RGB(122, 209, 81)
    colors(206) = RGB(124, 210, 80)
    colors(207) = RGB(127, 211, 78)
    colors(208) = RGB(129, 211, 77)
    colors(209) = RGB(132, 212, 75)
    colors(210) = RGB(134, 213, 73)
    colors(211) = RGB(137, 213, 72)
    colors(212) = RGB(139, 214, 70)
    colors(213) = RGB(142, 214, 69)
    colors(214) = RGB(144, 215, 67)
    colors(215) = RGB(147, 215, 65)
    colors(216) = RGB(149, 216, 64)
    colors(217) = RGB(152, 216, 62)
    colors(218) = RGB(155, 217, 60)
    colors(219) = RGB(157, 217, 59)
    colors(220) = RGB(160, 218, 57)
    colors(221) = RGB(162, 218, 55)
    colors(222) = RGB(165, 219, 54)
    colors(223) = RGB(168, 219, 52)
    colors(224) = RGB(170, 220, 50)
    colors(225) = RGB(173, 220, 48)
    colors(226) = RGB(176, 221, 47)
    colors(227) = RGB(178, 221, 45)
    colors(228) = RGB(181, 222, 43)
    colors(229) = RGB(184, 222, 41)
    colors(230) = RGB(186, 222, 40)
    colors(231) = RGB(189, 223, 38)
    colors(232) = RGB(192, 223, 37)
    colors(233) = RGB(194, 223, 35)
    colors(234) = RGB(197, 224, 33)
    colors(235) = RGB(200, 224, 32)
    colors(236) = RGB(202, 225, 31)
    colors(237) = RGB(205, 225, 29)
    colors(238) = RGB(208, 225, 28)
    colors(239) = RGB(210, 226, 27)
    colors(240) = RGB(213, 226, 26)
    colors(241) = RGB(216, 226, 25)
    colors(242) = RGB(218, 227, 25)
    colors(243) = RGB(221, 227, 24)
    colors(244) = RGB(223, 227, 24)
    colors(245) = RGB(226, 228, 24)
    colors(246) = RGB(229, 228, 25)
    colors(247) = RGB(231, 228, 25)
    colors(248) = RGB(234, 229, 26)
    colors(249) = RGB(236, 229, 27)
    colors(250) = RGB(239, 229, 28)
    colors(251) = RGB(241, 229, 29)
    colors(252) = RGB(244, 230, 30)
    colors(253) = RGB(246, 230, 32)
    colors(254) = RGB(248, 230, 33)
    colors(255) = RGB(251, 231, 35)
    colors(256) = RGB(253, 231, 37)

    Dim pullIndex As Variant
    pullIndex = CInt(((index - 1) / (totalCount - 1)) * 256) + 1
    
    If pullIndex < 1 Then
        pullIndex = 1
    ElseIf pullIndex > 256 Then
        pullIndex = 256
    End If

    Chart_GetViridis = colors(pullIndex)
    
    'take the string and make the colors
    
    'use the index to get all those to return

End Function

