Attribute VB_Name = "Module2"
' ******************************************************************************
' SolidWorks Macro with Excel Data Import - Looping through Rows - created by Xiao Liang
' ******************************************************************************

Dim swApp As Object
Dim Part As Object
Dim boolstatus As Boolean
Dim ExcelApp As Object
Dim ExcelWorkbook As Object
Dim ws As Object
Dim longstatus As Long
Dim longwarnings As Long

Sub main()


    Set swApp = Application.SldWorks
    swApp.Visible = False ' 隐藏 SolidWorks 界面
    
   ' Initialize Excel and open workbook
    Set ExcelApp = CreateObject("Excel.Application")
    Set ExcelWorkbook = ExcelApp.Workbooks.Open("D:\Close_loop_in_silico_optimization_showcase\Test.xlsx") ' 修改为你的Excel文件路径
    Set ws = ExcelWorkbook.Sheets("Solidworks") ' 修改为你的工作表名称

    ' Determine the last row with data in the worksheet
    Dim lastRow As Long
    
    lastRow = ws.UsedRange.Columns(2).Cells(ws.UsedRange.Rows.Count).Row
    
    Dim i As Long
    For i = 2 To 3 ' Assuming that the first row is the header and data starts from row 2
    
        ' Read Block 1 coordinates and rotation
        Dim block1Coords(1 To 6) As Double
        block1Coords(1) = ws.Cells(i, "B").Value
        block1Coords(2) = ws.Cells(i, "C").Value
        block1Coords(3) = ws.Cells(i, "D").Value
        block1Coords(4) = ws.Cells(i, "E").Value
        block1Coords(5) = ws.Cells(i, "F").Value
        block1Coords(6) = ws.Cells(i, "G").Value
        Dim block1Rotation As Double
        block1Rotation = ws.Cells(i, "H").Value
        
        ' Read Block 2 coordinates and rotation
        Dim block2Coords(1 To 6) As Double
        block2Coords(1) = ws.Cells(i, "I").Value
        block2Coords(2) = ws.Cells(i, "J").Value
        block2Coords(3) = ws.Cells(i, "K").Value
        block2Coords(4) = ws.Cells(i, "L").Value
        block2Coords(5) = ws.Cells(i, "M").Value
        block2Coords(6) = ws.Cells(i, "N").Value
        Dim block2Rotation As Double
        block2Rotation = ws.Cells(i, "O").Value
        
        ' Read Block 3 coordinates and rotation
        Dim block3Coords(1 To 6) As Double
        block3Coords(1) = ws.Cells(i, "P").Value
        block3Coords(2) = ws.Cells(i, "Q").Value
        block3Coords(3) = ws.Cells(i, "R").Value
        block3Coords(4) = ws.Cells(i, "S").Value
        block3Coords(5) = ws.Cells(i, "T").Value
        block3Coords(6) = ws.Cells(i, "U").Value
        Dim block3Rotation As Double
        block3Rotation = ws.Cells(i, "V").Value
        
        ' Read Block 4 coordinates and rotation
        Dim block4Coords(1 To 6) As Double
        block4Coords(1) = ws.Cells(i, "W").Value
        block4Coords(2) = ws.Cells(i, "X").Value
        block4Coords(3) = ws.Cells(i, "Y").Value
        block4Coords(4) = ws.Cells(i, "Z").Value
        block4Coords(5) = ws.Cells(i, "AA").Value
        block4Coords(6) = ws.Cells(i, "AB").Value
        Dim block4Rotation As Double
        block4Rotation = ws.Cells(i, "AC").Value

        
        ' Now apply these values in your SolidWorks operations

        ' Initialize SolidWorks and get active document
        

        
        ' Open
        Set Part = swApp.OpenDoc6("D:\Close_loop_in_silico_optimization_showcase\Blank.SLDPRT", 1, 0, "", longstatus, longwarnings)
        
        Dim COSMOSWORKSObj As Object
        Dim CWAddinCallBackObj As Object
        Set CWAddinCallBackObj = swApp.GetAddInObject("CosmosWorks.CosmosWorks")
        Set COSMOSWORKSObj = CWAddinCallBackObj.COSMOSWORKS
        
        
        ''''''block 1''''''
        Part.SketchManager.InsertSketch True
        boolstatus = Part.Extension.SelectByID2("base plane", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
        Part.SketchManager.CreateCenterRectangle block1Coords(1), block1Coords(2), block1Coords(3), block1Coords(4), block1Coords(5), block1Coords(6)
        Part.Extension.RotateOrCopy False, 1, True, block1Coords(1), block1Coords(2), block1Coords(3), 0, 0, 1, block1Rotation
        Dim swSketch1 As Object
        Set swSketch1 = Part.SketchManager.ActiveSketch
        swSketch1.Name = "block1_" & i
        Part.ClearSelection2 True
        Part.SketchManager.InsertSketch True
        boolstatus = Part.Extension.SelectByID2("block1_" & i, "SKETCH", 0, 0, 0, False, 0, Nothing, 0)
        Dim myFeature1 As Object
        Set myFeature1 = Part.FeatureManager.FeatureCut4(True, False, False, 0, 0, 0.001, 0.001, False, False, False, False, 1.74532925199433E-02, 1.74532925199433E-02, False, False, False, False, False, True, True, True, True, False, 0, 0, False, False)
        myFeature1.Name = "Block_1_" & i
        Part.SelectionManager.EnableContourSelection = False

        ''''''block 2''''''
        Part.SketchManager.InsertSketch True
        boolstatus = Part.Extension.SelectByID2("base plane", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
        Part.SketchManager.CreateCenterRectangle block2Coords(1), block2Coords(2), block2Coords(3), block2Coords(4), block2Coords(5), block2Coords(6)
        Part.Extension.RotateOrCopy False, 1, True, block2Coords(1), block2Coords(2), block2Coords(3), 0, 0, 1, block2Rotation
        Dim swSketch2 As Object
        Set swSketch2 = Part.SketchManager.ActiveSketch
        swSketch2.Name = "block2_" & i
        Part.ClearSelection2 True
        Part.SketchManager.InsertSketch True
        boolstatus = Part.Extension.SelectByID2("block2_" & i, "SKETCH", 0, 0, 0, False, 0, Nothing, 0)
        Dim myFeature2 As Object
        Set myFeature2 = Part.FeatureManager.FeatureCut4(True, False, False, 0, 0, 0.001, 0.001, False, False, False, False, 1.74532925199433E-02, 1.74532925199433E-02, False, False, False, False, False, True, True, True, True, False, 0, 0, False, False)
        myFeature2.Name = "Block_2_" & i
        Part.SelectionManager.EnableContourSelection = False

        ''''''block 3''''''
        Part.SketchManager.InsertSketch True
        boolstatus = Part.Extension.SelectByID2("base plane", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
        Part.SketchManager.CreateCenterRectangle block3Coords(1), block3Coords(2), block3Coords(3), block3Coords(4), block3Coords(5), block3Coords(6)
        Part.Extension.RotateOrCopy False, 1, True, block3Coords(1), block3Coords(2), block3Coords(3), 0, 0, 1, block3Rotation
        Dim swSketch3 As Object
        Set swSketch3 = Part.SketchManager.ActiveSketch
        swSketch3.Name = "block3_" & i
        Part.ClearSelection2 True
        Part.SketchManager.InsertSketch True
        boolstatus = Part.Extension.SelectByID2("block3_" & i, "SKETCH", 0, 0, 0, False, 0, Nothing, 0)
        Dim myFeature3 As Object
        Set myFeature3 = Part.FeatureManager.FeatureCut4(True, False, False, 0, 0, 0.001, 0.001, False, False, False, False, 1.74532925199433E-02, 1.74532925199433E-02, False, False, False, False, False, True, True, True, True, False, 0, 0, False, False)
        myFeature3.Name = "Block_3_" & i
        Part.SelectionManager.EnableContourSelection = False

        ''''''block 4''''''
        Part.SketchManager.InsertSketch True
        boolstatus = Part.Extension.SelectByID2("base plane", "PLANE", 0, 0, 0, False, 0, Nothing, 0)
        Part.SketchManager.CreateCenterRectangle block4Coords(1), block4Coords(2), block4Coords(3), block4Coords(4), block4Coords(5), block4Coords(6)
        Part.Extension.RotateOrCopy False, 1, True, block4Coords(1), block4Coords(2), block4Coords(3), 0, 0, 1, block4Rotation
        Dim swSketch4 As Object
        Set swSketch4 = Part.SketchManager.ActiveSketch
        swSketch4.Name = "block4_" & i
        Part.ClearSelection2 True
        Part.SketchManager.InsertSketch True
        boolstatus = Part.Extension.SelectByID2("block4_" & i, "SKETCH", 0, 0, 0, False, 0, Nothing, 0)
        Dim myFeature4 As Object
        Set myFeature4 = Part.FeatureManager.FeatureCut4(True, False, False, 0, 0, 0.001, 0.001, False, False, False, False, 1.74532925199433E-02, 1.74532925199433E-02, False, False, False, False, False, True, True, True, True, False, 0, 0, False, False)
        myFeature4.Name = "Block_4_" & i
        Part.SelectionManager.EnableContourSelection = False
        
        
        vBodies = Part.GetBodies2(swAllBodies, False)

        ' 初始化最大体积和索引
        maxVolume = 0
        maxVolumeIndex = -1
    
        ' 遍历所有实体体，找到体积最大的实体体
            For j = 0 To UBound(vBodies)
            Set swBody = vBodies(j)
            swMassProp = swBody.GetMassProperties(1) ' 使用默认单位（米-千克-秒）
        
            If Not IsEmpty(swMassProp) Then
                If swMassProp(3) > maxVolume Then
                    maxVolume = swMassProp(3)
                    maxVolumeIndex = j
                    maxVolumeBodyName = swBody.Name
                End If
            End If
        Next j
    
        ' 遍历所有实体体，删除体积小的实体体
        For j = 0 To UBound(vBodies)
            Set swBody = vBodies(j)
        
            ' 如果不是体积最大的实体体，删除它
            If swBody.Name <> maxVolumeBodyName Then
                boolstatus = Part.Extension.SelectByID2(swBody.Name, "SOLIDBODY", 0, 0, 0, False, 0, Nothing, 0)
                Part.ClearSelection2 True
                boolstatus = Part.Extension.SelectByID2(swBody.Name, "SOLIDBODY", 0, 0, 0, True, 0, Nothing, 0)
                Dim myFeature As Object
                Set myFeature = Part.FeatureManager.InsertDeleteBody2(False)
            

            End If
        Next j

        k = i - 1
        
        ' Save As SLDPRT file
        longstatus = Part.SaveAs3("D:\Close_loop_in_silico_optimization_showcase\Design" & k & ".SLDPRT", 0, 0)
        If longstatus <> 0 Then
            MsgBox "Error saving SLDPRT file for iteration " & k
        End If
        
        ' Ensure the file is properly saved
        Part.ClearSelection2 True

        ' Save As X_T file
        longstatus = Part.SaveAs3("D:\Close_loop_in_silico_optimization_showcase\Design" & k & ".X_T", 2, 2)
        If longstatus <> 0 Then
            MsgBox "Error saving X_T file for iteration " & k
        End If

        ' Close Document
        swApp.CloseDoc "Design" & k & ".SLDPRT"

    Next i
    
    ' Close Excel workbook and cleanup
    ExcelWorkbook.Close False
    ExcelApp.Quit
    Set ExcelApp = Nothing
    
    ' Clean up and release objects
    Set Part = Nothing
    Set swApp = Nothing

End Sub

