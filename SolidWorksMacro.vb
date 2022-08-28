Imports SolidWorks.Interop.sldworks
Imports SolidWorks.Interop.swconst
Imports System.Runtime.InteropServices
Imports System
Imports System.Diagnostics
Imports System.Windows.Forms.ContainerControl
Imports System.Windows.Forms.Form
Imports System.Windows.Forms
Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
Imports System.IO
Imports System.Collections
Imports System.Windows.Forms.Application

Partial Class SolidWorksMacro
    Dim swModel As ModelDoc2
    Dim swAssembly As AssemblyDoc
    Dim CheckInRod As String
    Dim value As Integer
    Dim Counter As Integer = 0
    Dim CheckArr(12) As Double
    Dim oExcel As Object
    Dim oBook As Object
    Dim oSheet As Object
    Dim Exceli As Integer = 1
    Dim Excelj As Integer = 0
    Dim X_len As Double
    Dim Y_len As Double
    Dim Z_len As Double
    Dim MateFreeSafe As Integer = 1
    Dim InitLength, InitWidth, InitHeight As Double
    Dim AssemblyPathFile As String
    Dim NumForMovePart As Double = 0
    Dim CountHole As Integer = 1
    Dim ArrForExcel(200, 35) As String
    Dim ChangeorientationForParallal As Integer = 0
    ''''''fo add the rod'''''
    Dim strCompName As String
    '''''For get direction in panel coord''''''
    Dim MoveSide As String
    Dim MoveForFreeSafe As String
    Dim MainAxisAss As String
    Dim SecondaryAxisAss As String
    Dim ThirdAxisAss As String
    Dim InverseMatrix(3, 3) As Double

    '''''From PD2C_EXT'''''
    Dim Product_id As String
    Dim Catalog_No As String
    Dim Component_ID As String
    Dim PartName As String
    Dim Short_name As String
    Dim Worientation As String
    Dim port As String
    Dim MainAxisPart As String
    Dim SIDE As String
    Dim Height_ID As Integer
    Dim SecondaryAxisPart As String
    Dim ThirdAxisPart As String
    Dim PORT_num As String
    Dim pos_XPart As Double
    Dim pos_YPart As Double
    Dim pos_ZPart As Double
    Dim Radius As Double
    Dim StepSize As Double
    Dim number_ports As Integer
    Dim port_num_indrement As Integer = 1
    Dim FreeDis As Double
    Dim SafeDis As Double
    Dim SafeDisUP As Double
    Dim Construction As String
    ''' User Input'''
    Dim PartsNumber As Integer
    Dim DirToRead As String
    Dim MainDir As String
    Dim DataFile As String
    Dim OutFile As String


    Public Sub Main()
        MainDir = "D:\ProcessSimulate\BGU_Pres\"

        Dim frm1 As New UserInput
        Application.Run(frm1)
        swModel = swApp.ActiveDoc
        swAssembly = swModel
        ''''ReadFromExcelTools''''
        PartsNumber = frm1.GSPartsNumber
        DirToRead = frm1.GSFileNameToRead
        DataFile = MainDir + DirToRead + "\Output\defs.xlsx"
        OutFile = MainDir + DirToRead + "\Output\configurations.xls"
        Dim oExcel As Object = CreateObject("Excel.Application")
        Dim oBookR As Object = oExcel.Workbooks.Open(DataFile)
        Dim oSheetR As Object = oBookR.Worksheets("PD2C_EXT")
        Dim PartToHoleName As String


        Dim RowNumExcel As Integer = 2
        While RowNumExcel <= PartsNumber + 1
            ReadFromExcel(oSheetR, RowNumExcel)
            oExcel = CreateObject("Excel.Application")
            oBook = oExcel.Workbooks.Add
            oSheet = oBook.ActiveSheet

            PartToHoleName = GetPartToHoleSize()
            AddPartToHoleToAssem(PartToHoleName)
            '''''put Rod in assem'''''
            Dim TempArr(2) As String
            GetTransform(PartName) 'part transform matrix
            TempArr = GetInfoWhereToMove()
            CheckInRod = strCompName
            CreateTransformRodPlace(CheckInRod, TempArr)
            '''''GET information'''''
            GetInputFromUser(PartName, MoveSide, MoveForFreeSafe)
            deleteRod()

            RowNumExcel = RowNumExcel + 1

        End While
        '''''''Save to Excel'''''''
        oBookR.Close(DataFile)
        SaveToExcel()
        oBook.SaveAs(OutFile, True)
        oExcel.Quit()




    End Sub
    Sub ReadFromExcel(ByVal oSheetR As Object, ByVal i As Integer)

        port_num_indrement = 1
        Dim j As Integer
        Dim InfoArr(21) As String

        For j = 0 To 21
            InfoArr(j) = oSheetR.cells(i, j + 1).Value
        Next
        Product_id = InfoArr(0)
        Catalog_No = InfoArr(1)
        Component_ID = InfoArr(2)
        PartName = Catalog_No & "-" + Component_ID

        Short_name = InfoArr(3)
        Worientation = InfoArr(4)
        port = InfoArr(5)
        MainAxisPart = InfoArr(6)
        Height_ID = InfoArr(7)
        SecondaryAxisPart = InfoArr(8)
        ThirdAxisPart = InfoArr(9)
        PORT_num = InfoArr(10)
        pos_XPart = InfoArr(11)
        pos_YPart = InfoArr(12)
        pos_ZPart = InfoArr(13)
        Radius = InfoArr(14)
        StepSize = InfoArr(15) / 1000
        number_ports = InfoArr(16)
        port_num_indrement = port_num_indrement + InfoArr(17)
        FreeDis = InfoArr(18) / 1000
        SafeDis = InfoArr(19) / 1000
        SafeDisUP = InfoArr(20)
        Construction = InfoArr(21)
    End Sub
    Public Function GetPartToHoleSize() As String
        Dim PartToHoleName As String
        If Radius < 2 Then
            PartToHoleName = "PartToHole2mm.sldprt"
        End If
        If Radius >= 2 AndAlso Radius < 4 Then
            PartToHoleName = "PartToHole4mm.sldprt"
        End If
        If Radius >= 4 AndAlso Radius < 6 Then
            PartToHoleName = "PartToHole6mm.sldprt"
        End If
        Return PartToHoleName
    End Function
    Sub AddPartToHoleToAssem(ByVal PartToHoleName As String)
        Dim PartFileName As String
        ' Open assembly
        swModel = swApp.ActiveDoc

        ' Set up event
        swAssemblyDoc = swModel
        openAssem = New Hashtable
        AttachEventHandlers()

        ' Get title of assembly document
        AssemblyTitle = swModel.GetTitle()

        ' Split the title into two strings using the period as the delimiter
        strings = Split(AssemblyTitle, ".")

        ' Use AssemblyName when mating the component with the assembly
        AssemblyName = strings(0)

        Debug.Print("Name of assembly: " & AssemblyName)

        boolstat = True
        Dim strCompModelname As String
        strCompModelname = PartToHoleName

        ' Because the component resides in the same folder as the assembly, get
        ' the assembly's path and use it when opening the component
        tmpPath = Microsoft.VisualBasic.Strings.Left(swModel.GetPathName, InStrRev(swModel.GetPathName, "\"))
        PartFileName = tmpPath & PartToHoleName
        ' Open the component
        tmpObj = swApp.OpenDoc6(tmpPath + strCompModelname, swDocumentTypes_e.swDocPART, swOpenDocOptions_e.swOpenDocOptions_Silent, "", errors, warnings)
        ' Check to see if the file is read-only or cannot be found; display error
        ' messages if either
        If warnings = swFileLoadWarning_e.swFileLoadWarning_ReadOnly Then
            MsgBox("This file is read-only.")
            boolstat = False
        End If

        If tmpObj Is Nothing Then
            MsgBox("Cannot locate the file.")
            boolstat = False
        End If

        ' Activate the assembly so that you can add the component to it
        swModel = swApp.ActivateDoc3(AssemblyTitle, True, swRebuildOnActivation_e.swUserDecision, errors)

        ' Add the camtest part to the assembly document
        swComponent = swAssemblyDoc.AddComponent5(strCompModelname, swAddComponentConfigOptions_e.swAddComponentConfigOptions_CurrentSelectedConfig, "", False, "", -1, -1, -1)

        ' Make the component virtual
        stat = swComponent.MakeVirtual2(True)

        ' Get the name of the component for the mate
        strCompName = swComponent.Name2()
        swApp.CloseDoc(PartFileName)


    End Sub

    Sub AttachEventHandlers()
        AttachSWEvents()
    End Sub


    Sub AttachSWEvents()
        AddHandler swAssemblyDoc.AddItemNotify, AddressOf Me.swAssemblyDoc_AddItemNotify
        AddHandler swAssemblyDoc.RenameItemNotify, AddressOf Me.swAssemblyDoc_RenameItemNotify
    End Sub

    Private Function swAssemblyDoc_AddItemNotify(ByVal EntityType As Integer, ByVal itemName As String) As Integer
        Debug.Print("Component added: " & itemName)
    End Function

    Private Function swAssemblyDoc_RenameItemNotify(ByVal EntityType As Integer, ByVal oldName As String, ByVal NewName As String) As Integer
        Debug.Print("Virtual component name: " & NewName)
    End Function
    Function GetInfoWhereToMove() As String()
        Dim j As Integer
        Dim temp As Integer
        Dim TempArr(2) As String

        For j = 0 To 2
            If TransMatrixPart(0, j) > 0.6 Then
                temp = 1
            End If
            If TransMatrixPart(0, j) < -0.6 Then
                temp = -1
            End If
            If TransMatrixPart(1, j) > 0.6 Then
                temp = 2
            End If
            If TransMatrixPart(1, j) < -0.6 Then
                temp = -2
            End If
            If TransMatrixPart(2, j) > 0.6 Then
                temp = 3
            End If
            If TransMatrixPart(2, j) < -0.6 Then
                temp = -3
            End If
            TempArr(j) = temp
        Next

        Return TempArr
    End Function
    Public Sub CreateTransformRodPlace(ByVal PartName As String, ByVal TempArr As String())
        Dim mathUtility As MathUtility
        Dim transform As MathTransform
        Dim boolstatus As Boolean = False
        Dim swModel As ModelDoc2
        Dim swModelDocExt As ModelDocExtension
        Dim swAssembly As AssemblyDoc
        Dim swComp As Object = Nothing
        Dim RotationMatrix(3, 3) As Double
        Dim SideNeg(3, 3) As Double
        Dim xrot, yrot, zrot As Double

        swModel = swApp.ActiveDoc
        swModelDocExt = swModel.Extension
        boolstatus = swModelDocExt.SelectByID2(PartName, "COMPONENT", 0, 0, 0, False, 0, Nothing, 0)
        swAssembly = swModel
        mathUtility = swApp.GetMathUtility

        Dim transformArray(15) As Double

        If (MainAxisPart = "x" OrElse MainAxisPart = "+x" OrElse MainAxisPart = "-x") AndAlso (SecondaryAxisPart = "z" OrElse SecondaryAxisPart = "-z") Then
            RotationMatrix(0, 0) = 0 : RotationMatrix(0, 1) = 0 : RotationMatrix(0, 2) = 1
            RotationMatrix(1, 0) = 0 : RotationMatrix(1, 1) = 1 : RotationMatrix(1, 2) = 0
            RotationMatrix(2, 0) = -1 : RotationMatrix(2, 1) = 0 : RotationMatrix(2, 2) = 0

            MoveForFreeSafe = TempArr(0)
            MoveSide = TempArr(2)
            ThirdAxisAss = TempArr(1)
            If MainAxisPart = "-x" Then
                MoveForFreeSafe = -MoveForFreeSafe

            End If
            If SecondaryAxisPart = "-z" Then
                MoveSide = -MoveSide
            End If
            If ThirdAxisPart = "-y" Then
                ThirdAxisAss = -ThirdAxisAss
            End If
            xrot = pos_XPart : yrot = pos_YPart : zrot = pos_ZPart
        End If

        If (MainAxisPart = "x" OrElse MainAxisPart = "+x" OrElse MainAxisPart = "-x") AndAlso (SecondaryAxisPart = "y" OrElse SecondaryAxisPart = "+y" OrElse SecondaryAxisPart = "-y") Then
            RotationMatrix(0, 0) = 0 : RotationMatrix(0, 1) = 0 : RotationMatrix(0, 2) = 1
            RotationMatrix(1, 0) = 0 : RotationMatrix(1, 1) = 1 : RotationMatrix(1, 2) = 0
            RotationMatrix(2, 0) = -1 : RotationMatrix(2, 1) = 0 : RotationMatrix(2, 2) = 0
            MoveForFreeSafe = TempArr(0)
            MoveSide = TempArr(1)
            ThirdAxisAss = TempArr(2)
            If MainAxisPart = "-x" Then
                MoveForFreeSafe = -MoveForFreeSafe

            End If
            If SecondaryAxisPart = "-y" Then
                MoveSide = -MoveSide
            End If
            If ThirdAxisPart = "-z" Then
                ThirdAxisAss = -ThirdAxisAss
            End If
            xrot = pos_XPart : yrot = pos_YPart : zrot = pos_ZPart
        End If

        If (MainAxisPart = "y" OrElse MainAxisPart = "+y" OrElse MainAxisPart = "-y") AndAlso (SecondaryAxisPart = "x" OrElse SecondaryAxisPart = "+x" OrElse SecondaryAxisPart = "-x") Then
            RotationMatrix(0, 0) = 1 : RotationMatrix(0, 1) = 0 : RotationMatrix(0, 2) = 0
            RotationMatrix(1, 0) = 0 : RotationMatrix(1, 1) = 0 : RotationMatrix(1, 2) = 1
            RotationMatrix(2, 0) = 0 : RotationMatrix(2, 1) = -1 : RotationMatrix(2, 2) = 0
            MoveForFreeSafe = TempArr(1)
            MoveSide = TempArr(0)
            ThirdAxisAss = TempArr(2)
            If MainAxisPart = "-y" Then
                MoveForFreeSafe = -MoveForFreeSafe

            End If
            If SecondaryAxisPart = "-x" Then
                MoveSide = -MoveSide
            End If
            If ThirdAxisPart = "-z" Then
                ThirdAxisAss = -ThirdAxisAss
            End If
            xrot = pos_XPart : yrot = pos_YPart : zrot = pos_ZPart

        End If
        If (MainAxisPart = "y" OrElse MainAxisPart = "+y" OrElse MainAxisPart = "-y") AndAlso (SecondaryAxisPart = "z" OrElse SecondaryAxisPart = "+z" OrElse SecondaryAxisPart = "-z") Then
            RotationMatrix(0, 0) = 0 : RotationMatrix(0, 1) = 0 : RotationMatrix(0, 2) = -1
            RotationMatrix(1, 0) = 1 : RotationMatrix(1, 1) = 0 : RotationMatrix(1, 2) = 0
            RotationMatrix(2, 0) = 0 : RotationMatrix(2, 1) = 1 : RotationMatrix(2, 2) = 0
            MoveForFreeSafe = TempArr(1)
            MoveSide = TempArr(2)
            ThirdAxisAss = TempArr(0)
            If MainAxisPart = "-y" Then
                MoveForFreeSafe = -MoveForFreeSafe

            End If
            If SecondaryAxisPart = "-z" Then
                MoveSide = -MoveSide
            End If
            If ThirdAxisPart = "-x" Then
                ThirdAxisAss = -ThirdAxisAss
            End If
            xrot = pos_XPart : yrot = pos_YPart : zrot = pos_ZPart
        End If

        If (MainAxisPart = "z" OrElse MainAxisPart = "+z" OrElse MainAxisPart = "-z") AndAlso (SecondaryAxisPart = "x" OrElse SecondaryAxisPart = "+x" OrElse SecondaryAxisPart = "-x") Then
            RotationMatrix(0, 0) = -1 : RotationMatrix(0, 1) = 0 : RotationMatrix(0, 2) = 0
            RotationMatrix(1, 0) = 0 : RotationMatrix(1, 1) = -1 : RotationMatrix(1, 2) = 0
            RotationMatrix(2, 0) = 0 : RotationMatrix(2, 1) = 0 : RotationMatrix(2, 2) = 1

            MoveForFreeSafe = TempArr(2)
            MoveSide = TempArr(0)
            ThirdAxisAss = TempArr(1)
            If MainAxisPart = "-z" Then
                MoveForFreeSafe = -MoveForFreeSafe

            End If
            If SecondaryAxisPart = "-x" Then
                MoveSide = -MoveSide
            End If
            If ThirdAxisPart = "-y" Then
                ThirdAxisAss = -ThirdAxisAss
            End If
            xrot = pos_XPart : yrot = pos_YPart : zrot = pos_ZPart

        End If

        If (MainAxisPart = "z" OrElse MainAxisPart = "+z" OrElse MainAxisPart = "-z") AndAlso (SecondaryAxisPart = "y" OrElse SecondaryAxisPart = "+y" OrElse SecondaryAxisPart = "-y") Then
            RotationMatrix(0, 0) = 0 : RotationMatrix(0, 1) = -1 : RotationMatrix(0, 2) = 0
            RotationMatrix(1, 0) = -1 : RotationMatrix(1, 1) = 0 : RotationMatrix(1, 2) = 0
            RotationMatrix(2, 0) = 0 : RotationMatrix(2, 1) = 0 : RotationMatrix(2, 2) = -1
            MoveForFreeSafe = TempArr(2)
            MoveSide = TempArr(1)
            ThirdAxisAss = TempArr(0)
            If MainAxisPart = "-z" Then
                MoveForFreeSafe = -MoveForFreeSafe

            End If
            If SecondaryAxisPart = "-y" Then
                MoveSide = -MoveSide
            End If
            If ThirdAxisPart = "-x" Then
                ThirdAxisAss = -ThirdAxisAss
            End If
            xrot = pos_XPart : yrot = pos_YPart : zrot = pos_ZPart
        End If

        If (MoveForFreeSafe = "-3") Then
            For i = 0 To 3
                For j = 0 To 3
                    SideNeg(i, j) = 0
                Next
            Next

            SideNeg(0, 0) = -1
            SideNeg(1, 1) = 1
            SideNeg(2, 2) = -1

            RotationMatrix = MultiplcationMatrixs(SideNeg, RotationMatrix, "Part in PanelCoord")
        End If

        RotationMatrix(0, 3) = xrot : RotationMatrix(1, 3) = yrot : RotationMatrix(2, 3) = zrot
        RotationMatrix(3, 3) = 1

        Dim PartInPanelCoord(3, 3) As Double
        PartInPanelCoord = MultiplcationMatrixs(TransMatrixPart, RotationMatrix, "Part in PanelCoord")

        'Rotation. Note that there are three assignment per line to shorten the macro
        transformArray(0) = PartInPanelCoord(0, 0) : transformArray(1) = PartInPanelCoord(1, 0) : transformArray(2) = PartInPanelCoord(2, 0)
        transformArray(3) = PartInPanelCoord(0, 1) : transformArray(4) = PartInPanelCoord(1, 1) : transformArray(5) = PartInPanelCoord(2, 1)
        transformArray(6) = PartInPanelCoord(0, 2) : transformArray(7) = PartInPanelCoord(1, 2) : transformArray(8) = PartInPanelCoord(2, 2)
        'Translation

        transformArray(9) = PartInPanelCoord(0, 3) / 1000 : transformArray(10) = PartInPanelCoord(1, 3) / 1000 : transformArray(11) = PartInPanelCoord(2, 3) / 1000
        'Scale
        transformArray(12) = 1


        transform = mathUtility.CreateTransform(transformArray)

        Dim transformData As Object
        transformData = transform.ArrayData
        swComp = swModel.ISelectionManager.GetSelectedObjectsComponent4(1, -1)
        boolstatus = swComp.SetTransformAndSolve2(transform)
        boolstatus = swAssembly.ForceRebuild3(False)
        'swModel.ClearSelection2(True)
    End Sub

    Function GetMax(ByVal Val1 As Double, ByVal Val2 As Double) As Double
        ' Finds maximum of 3 values
        GetMax = Val1
        If Val2 > GetMax Then
            GetMax = Val2
        End If
    End Function

    Function GetMin(ByVal Val1 As Double, ByVal Val2 As Double) As Double
        ' Finds minimum of 3 values
        GetMin = Val1
        If Val2 < GetMin Then
            GetMin = Val2
        End If
    End Function

    Public Sub GetInputFromUser(ByVal PartName As String, ByVal MoveSide As String, ByVal MoveForFreeSafe As String)
        Dim X_move As Double
        Dim Y_move As Double
        Dim Z_move As Double
        Dim MovingTo As Integer
        Dim CheckOrientation As Integer = 0
        If number_ports = 0 Then
            number_ports = 1
        End If

        For CountHole = 1 To number_ports
            If Short_name <> "Panel" Then
                If CountHole = 1 Then
                    'CheckCollisionWithRod(CheckInRod)
                End If
                MoveForFreeSafe = ChangeNamesForAxis(MoveForFreeSafe)
                MoveSide = ChangeNamesForAxis(MoveSide)
                MainAxisAss = MoveForFreeSafe
                SecondaryAxisAss = MoveSide
                While CheckOrientation <> 2
                    MovingTo = 1
                    GetTransform(CheckInRod) ' Mate transform matrix
                    ''''''MoveForFreeMatrix''''''
                    X_move = FreeDis
                    Y_move = FreeDis
                    Z_move = FreeDis
                    OutputCompXform(X_move, Y_move, Z_move, MoveForFreeSafe, MovingTo)
                    ''''''MoveForSafeMatrix''''''
                    X_move = SafeDis
                    Y_move = SafeDis
                    Z_move = SafeDis
                    OutputCompXform(X_move, Y_move, Z_move, MoveForFreeSafe, MovingTo)
                    MateFreeSafe = 1

                    If ThirdAxisPart IsNot Nothing Then
                        MateFreeSafe = 2
                        ThirdAxisAss = ChangeNamesForAxis(ThirdAxisAss)
                        MoveForFreeSafe = ThirdAxisAss
                        MainAxisAss = ThirdAxisAss
                        CheckOrientation = CheckOrientation + 1
                        MainAxisPart = ThirdAxisPart
                        ChangeorientationForParallal = 1
                    Else
                        CheckOrientation = 2
                    End If
                End While
                MateFreeSafe = 1
                ChangeorientationForParallal = 0
            Else
                GetTransform(CheckInRod) ' Mate transform matrix

            End If

            '''''''MoveForTheNextPort''''''
            MovingTo = 2
            If CountHole <> number_ports Then
                CheckOrientation = 1
                OutputCompXform(X_move, Y_move, Z_move, MoveSide, MovingTo)
            End If
            PORT_num += port_num_indrement


        Next
        MoveForFreeSafe = Nothing
        MoveSide = Nothing
        MainAxisAss = Nothing
        SecondaryAxisAss = Nothing
    End Sub
    Public WithEvents swAssemblyDoc As AssemblyDoc
    Dim swDocExt As ModelDocExtension
    Dim openAssem As Hashtable
    Dim tmpPath As String
    Dim tmpObj As ModelDoc2
    Dim boolstat As Boolean, stat As Boolean
    Dim strings As Object
    Dim swComponent As Component2
    Dim matefeature As Feature
    Dim MateName As String
    Dim FirstSelection As String
    Dim SecondSelection As String
    Dim AssemblyTitle As String
    Dim AssemblyName As String
    Dim errors As Integer
    Dim warnings As Integer
    Dim mateError As Integer
    Dim fileName As String


    Public Function FindInterfacesDetNoExc(ByVal PartForCheck As String) As Integer
        Dim pIntMgr As InterferenceDetectionMgr
        Dim vInts As Object
        Dim i As Long, j As Long, k As Long
        Dim interference As IInterference
        Dim vComps As Object = Nothing
        Dim comp As Component2
        Dim vol As Double
        Dim vTrans As Object = Nothing
        Dim ret As Boolean
        Dim CountNumOfBaseInterFaces As Integer = 0, CountNumOfPartInterFaces As Integer = 0
        Dim t As Integer
        Dim ReturnCond As Integer = 0
        Dim CounterForSide As Integer = 0
        Dim CheckVol(1000) As Double
        Dim PartsName(1000) As String


        'Open the Interference Detection pane
        swAssemblyDoc.ToolsCheckInterference()

        pIntMgr = swAssemblyDoc.InterferenceDetectionManager


        'Specify the interference detection settings and options
        pIntMgr.TreatCoincidenceAsInterference = False
        pIntMgr.TreatSubAssembliesAsComponents = True
        pIntMgr.IncludeMultibodyPartInterferences = True
        pIntMgr.MakeInterferingPartsTransparent = True
        pIntMgr.CreateFastenersFolder = False
        pIntMgr.IgnoreHiddenBodies = False
        pIntMgr.ShowIgnoredInterferences = False
        pIntMgr.UseTransform = False


        'Specify how to display non-interfering components
        pIntMgr.NonInterferingComponentDisplay = swNonInterferingComponentDisplay_e.swNonInterferingComponentDisplay_Wireframe


        'Run interference detection
        vInts = pIntMgr.GetInterferences
        'Debug.Print("# of interferences: " & pIntMgr.GetInterferenceCount)


        'Get interfering components and transforms
        ret = pIntMgr.GetComponentsAndTransforms(vComps, vTrans)
        'Get interference information

        If (vInts IsNot Nothing) Then
            For i = 0 To UBound(vInts)
                'Debug.Print("Interference " & (i + 1))
                interference = vInts(i)
                'Debug.Print("  Number of components in this interference: " & interference.GetComponentCount)
                vComps = interference.Components
                For j = 0 To UBound(vComps)

                    comp = vComps(j)
                    If comp.Name2 = PartForCheck Then
                        CountNumOfPartInterFaces = 1
                    End If
                Next j
                If CountNumOfPartInterFaces = 1 Then
                    For t = 0 To UBound(vComps)
                        comp = vComps(t)
                        vol = interference.Volume
                        CheckVol(CounterForSide) = (vol * 1000000000)
                        PartsName(CounterForSide) = comp.Name2
                        'Debug.Print("   " & comp.Name2)
                        CounterForSide += 1

                        ReturnCond += CountNumOfPartInterFaces

                    Next
                End If
                CountNumOfPartInterFaces = 0
                'Debug.Print("  Interference volume is " & (vol * 1000000000) & " mm^3")
            Next i
        End If

        'Stop interference detection and close Interference Detection pane
        pIntMgr.Done()
        Return ReturnCond
    End Function
    Public Sub CheckCollisionWithRod(ByVal CompToMove As String)
        Dim swComp As Component2
        Dim sPadStr As String
        Dim swCompXform As MathTransform
        Dim vXform As Object
        Dim swMathUtil As MathUtility
        Dim swModel As ModelDoc2
        Dim swModelDocExt As ModelDocExtension
        Dim swSelMgr As SelectionMgr
        Dim bRet As Boolean
        Dim TranArr(9) As Double
        Dim Length, Width, Height As Double

        swModel = swApp.ActiveDoc
        swModelDocExt = swModel.Extension
        bRet = swModelDocExt.SelectByID2(CompToMove, "COMPONENT", 0, 0, 0, False, 0, Nothing, 0)

        swMathUtil = swApp.GetMathUtility
        swSelMgr = swModel.SelectionManager
        swComp = swSelMgr.GetSelectedObjectsComponent(1)



        ' Null for root component

        swCompXform = swComp.Transform2

        If Not swCompXform Is Nothing Then

            vXform = swCompXform.ArrayData

            ' Root component has no name

            'Debug.Print(sPadStr & "Component = " & swComp.Name2 & " (" & swComp.ReferencedConfiguration & ")")

            'Debug.Print(sPadStr & "  Suppr   = " & swComp.IsSuppressed)

            'Debug.Print(sPadStr & "  Hidden  = " & swComp.IsHidden(False))

            'Debug.Print(sPadStr & "  Rot1  = (" + Str(vXform(0)) + ", " + Str(vXform(1)) + ", " + Str(vXform(2)) + ")")
            'Debug.Print(sPadStr & "  Rot2  = (" + Str(vXform(3)) + ", " + Str(vXform(4)) + ", " + Str(vXform(5)) + ")")
            'Debug.Print(sPadStr & "  Rot3  = (" + Str(vXform(6)) + ", " + Str(vXform(7)) + ", " + Str(vXform(8)) + ")")
            'Debug.Print(sPadStr & "  Trans = (" + Str(vXform(9)) + ", " + Str(vXform(10)) + ", " + Str(vXform(11)) + ")")
            ' Debug.Print(sPadStr & "  Scale = " + Str(vXform(12)))
            'Debug.Print("")

        End If


        Length = vXform(9)
        Width = vXform(10)
        Height = vXform(11)
        Debug.Print("")
        Debug.Print("The part that check is: " & CompToMove)

        MovePartAndCheck(vXform, CompToMove, Length, Width, Height, -MoveForFreeSafe) 'for inside
        MovePartAndCheck(vXform, CompToMove, Length, Width, Height, MoveSide) 'for side1
        MovePartAndCheck(vXform, CompToMove, Length, Width, Height, -MoveSide) 'for side2
        CreateTransform(vXform, CompToMove, SaveNumForMovePart(0) / 3, SaveNumForMovePart(1) / 3, SaveNumForMovePart(2) / 3)
        SaveNumForMovePart(0) = 0 : SaveNumForMovePart(1) = 0 : SaveNumForMovePart(2) = 0
    End Sub
    Dim SumCollid As Integer = 0
    Dim SaveNumForMovePart(2) As Double
    Public Sub MovePartAndCheck(ByVal vXform() As Double, ByVal CompToMove As String, ByVal Length As Double, ByVal Width As Double, ByVal Height As Double, ByVal Direction As String)
        Dim WhileCond As Integer = 0
        While WhileCond = 0
            NumForMovePart += 0.0005
            If Direction = "+1" OrElse Direction = "1" Then 'For +x
                CreateTransform(vXform, CompToMove, Length + NumForMovePart, Width, Height)
                WhileCond = FindInterfacesDetNoExc(CompToMove)
                SaveNumForMovePart(0) = SaveNumForMovePart(0) + Length + NumForMovePart
                SaveNumForMovePart(1) = SaveNumForMovePart(1) + Width
                SaveNumForMovePart(2) = SaveNumForMovePart(2) + Height
            End If
            If Direction = "-1" Then 'For -x
                CreateTransform(vXform, CompToMove, Length - NumForMovePart, Width, Height)
                WhileCond = FindInterfacesDetNoExc(CompToMove)
                SaveNumForMovePart(0) = SaveNumForMovePart(0) + Length - NumForMovePart
                SaveNumForMovePart(1) = SaveNumForMovePart(1) + Width
                SaveNumForMovePart(2) = SaveNumForMovePart(2) + Height
            End If
            If Direction = "+2" OrElse Direction = "2" Then 'For +y
                CreateTransform(vXform, CompToMove, Length, Width + NumForMovePart, Height)
                WhileCond = FindInterfacesDetNoExc(CompToMove)
                SaveNumForMovePart(0) = SaveNumForMovePart(0) + Length
                SaveNumForMovePart(1) = SaveNumForMovePart(1) + Width + NumForMovePart
                SaveNumForMovePart(2) = SaveNumForMovePart(2) + Height
            End If
            If Direction = "-2" Then 'For -y
                CreateTransform(vXform, CompToMove, Length, Width - NumForMovePart, Height)
                WhileCond = FindInterfacesDetNoExc(CompToMove)
                SaveNumForMovePart(0) = SaveNumForMovePart(0) + Length
                SaveNumForMovePart(1) = SaveNumForMovePart(1) + Width - NumForMovePart
                SaveNumForMovePart(2) = SaveNumForMovePart(2) + Height
            End If
            If Direction = "+3" OrElse Direction = "3" Then 'For +z
                CreateTransform(vXform, CompToMove, Length, Width, Height + NumForMovePart)
                WhileCond = FindInterfacesDetNoExc(CompToMove)
                SaveNumForMovePart(0) = SaveNumForMovePart(0) + Length
                SaveNumForMovePart(1) = SaveNumForMovePart(1) + Width
                SaveNumForMovePart(2) = SaveNumForMovePart(2) + Height + NumForMovePart
            End If
            If Direction = "-3" Then 'For -z
                CreateTransform(vXform, CompToMove, Length, Width, Height - NumForMovePart)
                WhileCond = FindInterfacesDetNoExc(CompToMove)
                SaveNumForMovePart(0) = SaveNumForMovePart(0) + Length
                SaveNumForMovePart(1) = SaveNumForMovePart(1) + Width
                SaveNumForMovePart(2) = SaveNumForMovePart(2) + Height - NumForMovePart
            End If
            If WhileCond > 0 Then
                Debug.Print("There Is Collision In Direction " & Direction)
                SumCollid = SumCollid + 1
                CreateTransform(vXform, CompToMove, Length, Width, Height)
            Else
                If NumForMovePart >= 0.008 Then
                    WhileCond = 1
                    Debug.Print("There is No Collision in Direction - Check What the problem " & Direction)
                    SaveNumForMovePart(0) = SaveNumForMovePart(0) + Length
                    SaveNumForMovePart(1) = SaveNumForMovePart(1) + Width
                    SaveNumForMovePart(2) = SaveNumForMovePart(2) + Height
                End If
            End If


        End While

        NumForMovePart = 0

    End Sub

    Public Function ChangeNamesForAxis(ByVal MoveTo As String) As String
        If MoveTo = "1" Then
            MoveTo = "+x"
        End If
        If MoveTo = "-1" Then
            MoveTo = "-x"
        End If
        If MoveTo = "2" Then
            MoveTo = "+y"
        End If
        If MoveTo = "-2" Then
            MoveTo = "-y"
        End If
        If MoveTo = "3" Then
            MoveTo = "+z"
        End If
        If MoveTo = "-3" Then
            MoveTo = "-z"
        End If
        Return MoveTo
    End Function

    Public Sub OutputCompXform(ByVal X_move As Double, ByVal Y_move As Double, ByVal Z_move As Double, WhichSideToMove As String, ByVal MovingTo As Integer)
        Dim swComp As Component2
        Dim sPadStr As String
        Dim swCompXform As MathTransform
        Dim vXform As Object
        Dim swMathUtil As MathUtility
        Dim swModel As ModelDoc2
        Dim swModelDocExt As ModelDocExtension
        Dim swSelMgr As SelectionMgr
        Dim bRet As Boolean
        Dim TranArr(9) As Double
        Dim Length, Width, Height As Double

        swModel = swApp.ActiveDoc
        swModelDocExt = swModel.Extension
        bRet = swModelDocExt.SelectByID2(CheckInRod, "COMPONENT", 0, 0, 0, False, 0, Nothing, 0)

        swMathUtil = swApp.GetMathUtility
        swSelMgr = swModel.SelectionManager
        swComp = swSelMgr.GetSelectedObjectsComponent(1)



        ' Null for root component

        swCompXform = swComp.Transform2

        If Not swCompXform Is Nothing Then

            vXform = swCompXform.ArrayData

            ' Root component has no name

            'Debug.Print(sPadStr & "Component = " & swComp.Name2 & " (" & swComp.ReferencedConfiguration & ")")

            'Debug.Print(sPadStr & "  Suppr   = " & swComp.IsSuppressed)

            'Debug.Print(sPadStr & "  Hidden  = " & swComp.IsHidden(False))

            'Debug.Print(sPadStr & "  Rot1  = (" + Str(vXform(0)) + ", " + Str(vXform(1)) + ", " + Str(vXform(2)) + ")")
            'Debug.Print(sPadStr & "  Rot2  = (" + Str(vXform(3)) + ", " + Str(vXform(4)) + ", " + Str(vXform(5)) + ")")
            'Debug.Print(sPadStr & "  Rot3  = (" + Str(vXform(6)) + ", " + Str(vXform(7)) + ", " + Str(vXform(8)) + ")")
            'Debug.Print(sPadStr & "  Trans = (" + Str(vXform(9)) + ", " + Str(vXform(10)) + ", " + Str(vXform(11)) + ")")
            'Debug.Print(sPadStr & "  Scale = " + Str(vXform(12)))
            'Debug.Print("")

        End If

        Length = vXform(9)
        Width = vXform(10)
        Height = vXform(11)
        'Debug.Print("")
        'Debug.Print("The part that check Is " & CheckInRod)
        If WhichSideToMove = "+x" Then
            If MovingTo = 1 Then
                NumForMovePart = X_move
            Else
                NumForMovePart = StepSize
            End If
            CreateTransform(vXform, CheckInRod, Length + NumForMovePart, Width, Height) 'for Length +

        End If
        If WhichSideToMove = "-x" Then
            If MovingTo = 1 Then
                NumForMovePart = X_move
            Else
                NumForMovePart = StepSize

            End If
            CreateTransform(vXform, CheckInRod, Length - NumForMovePart, Width, Height) 'for Length - 
        End If

        If WhichSideToMove = "+y" Then
            If MovingTo = 1 Then
                NumForMovePart = Y_move
            Else
                NumForMovePart = StepSize

            End If
            CreateTransform(vXform, CheckInRod, Length, Width + NumForMovePart, Height) 'for Width +
        End If

        If WhichSideToMove = "-y" Then
            If MovingTo = 1 Then
                NumForMovePart = Y_move
            Else
                NumForMovePart = StepSize
            End If
            CreateTransform(vXform, CheckInRod, Length, Width - NumForMovePart, Height) 'for Width -
        End If

        If WhichSideToMove = "+z" Then
            If MovingTo = 1 Then
                NumForMovePart = Z_move
            Else
                NumForMovePart = StepSize
            End If
            CreateTransform(vXform, CheckInRod, Length, Width, Height + NumForMovePart) 'for Height +
        End If

        If WhichSideToMove = "-z" Then
            If MovingTo = 1 Then
                NumForMovePart = Z_move
            Else
                NumForMovePart = StepSize
            End If
            CreateTransform(vXform, CheckInRod, Length, Width, Height - NumForMovePart) 'for Height -
        End If
        GetTransform(CheckInRod)
        If MateFreeSafe < 5 Then
            InitLength = Length
            InitWidth = Width
            InitHeight = Height
        End If
        If MovingTo = 1 AndAlso MateFreeSafe = 5 Then
            CreateTransform(vXform, CheckInRod, InitLength, InitWidth, InitHeight)
        End If
    End Sub

    Public Sub CreateTransform(ByVal TranArr() As Double, ByVal CompName As String, ByVal Length As Double, ByVal Width As Double, ByVal Height As Double)
        Dim mathUtility As MathUtility
        Dim transform As MathTransform
        Dim boolstatus As Boolean = False
        Dim swModel As ModelDoc2
        Dim swModelDocExt As ModelDocExtension
        Dim swAssembly As AssemblyDoc
        Dim swComp As Object = Nothing


        swModel = swApp.ActiveDoc
        swModelDocExt = swModel.Extension
        boolstatus = swModelDocExt.SelectByID2(CompName, "COMPONENT", 0, 0, 0, False, 0, Nothing, 0)
        swAssembly = swModel
        mathUtility = swApp.GetMathUtility

        Dim transformArray(15) As Double
        'Rotation. Note that there are three assignment per line to shorten the macro
        transformArray(0) = TranArr(0) : transformArray(1) = TranArr(1) : transformArray(2) = TranArr(2)
        transformArray(3) = TranArr(3) : transformArray(4) = TranArr(4) : transformArray(5) = TranArr(5)
        transformArray(6) = TranArr(6) : transformArray(7) = TranArr(7) : transformArray(8) = TranArr(8)
        'Translation
        transformArray(9) = Length : transformArray(10) = Width : transformArray(11) = Height


        'Scale
        transformArray(12) = TranArr(12)


        transform = mathUtility.CreateTransform(transformArray)

        Dim transformData As Object
        transformData = transform.ArrayData
        swComp = swModel.ISelectionManager.GetSelectedObjectsComponent4(1, -1)
        boolstatus = swComp.SetTransformAndSolve2(transform)
        boolstatus = swAssembly.ForceRebuild3(False)
        'swModel.ClearSelection2(True)
    End Sub
    Dim MultiPlacMatrixMate(3, 3) As Double
    Dim MultiPlacMatrixFree(3, 3) As Double
    Dim MultiPlacMatrixSafe(3, 3) As Double
    Dim TransMatrixPart(3, 3) As Double
    Dim TransMatrixMate(3, 3) As Double
    Dim TransMatrixFree(3, 3) As Double
    Dim TransMatrixSafe(3, 3) As Double
    Dim quatPartCoordMateSide1(3) As Double
    Dim quatPartCoordMateSide2(3) As Double
    Dim quatPartCoordFreeSide1(3) As Double
    Dim quatPartCoordFreeSide2(3) As Double
    Dim quatPartCoordSafeSide1(3) As Double
    Dim quatPartCoordSafeSide2(3) As Double
    Dim quatPanelCoordMateSide1(3) As Double
    Dim quatPanelCoordMateSide2(3) As Double
    Dim quatPanelCoordFreeSide1(3) As Double
    Dim quatPanelCoordFreeSide2(3) As Double
    Dim quatPanelCoordSafeSide1(3) As Double
    Dim quatPanelCoordSafeSide2(3) As Double
    Dim MatrixForQuatTranSide1o(3, 3) As Double
    Dim MatrixForQuatTranSide2o(3, 3) As Double
    Public Sub GetTransform(ByVal PartName As String)
        Dim swComp As Component2
        Dim sPadStr As String
        Dim swCompXform As MathTransform
        Dim vXform As Object
        Dim swMathUtil As MathUtility
        Dim swModel As ModelDoc2
        Dim swModelDocExt As ModelDocExtension
        Dim swSelMgr As SelectionMgr
        Dim bRet As Boolean
        Dim NameToExcel As String
        Dim TransMatrix(3, 3) As Double
        Dim MultiMatrix(3, 3) As Double
        Dim IdentityMatrix(3, 3) As Double


        Dim OperationSide As Integer = 1
        For i = 0 To 3
            For j = 0 To 3
                MatrixForQuatTranSide1o(i, j) = 0
                MatrixForQuatTranSide2o(i, j) = 0
                IdentityMatrix(i, j) = 0
            Next
        Next
        'For SIDE 1 in BAR
        MatrixForQuatTranSide1o(0, 0) = 1
        MatrixForQuatTranSide1o(1, 2) = 1
        MatrixForQuatTranSide1o(2, 1) = -1
        MatrixForQuatTranSide1o(3, 3) = 1
        'For SIDE 2 out BAR
        MatrixForQuatTranSide2o(0, 0) = -1
        MatrixForQuatTranSide2o(1, 2) = 1
        MatrixForQuatTranSide2o(2, 1) = 1
        MatrixForQuatTranSide2o(3, 3) = 1

        'For identityMatrix 
        IdentityMatrix(0, 0) = 1
        IdentityMatrix(1, 1) = 1
        IdentityMatrix(2, 2) = 1
        IdentityMatrix(3, 3) = 1

        swModel = swApp.ActiveDoc
        swModelDocExt = swModel.Extension
        bRet = swModelDocExt.SelectByID2(PartName, "COMPONENT", 0, 0, 0, False, 0, Nothing, 0)

        swMathUtil = swApp.GetMathUtility
        swSelMgr = swModel.SelectionManager
        swComp = swSelMgr.GetSelectedObjectsComponent(1)


        ' Null for root component

        swCompXform = swComp.Transform2

        If Not swCompXform Is Nothing Then

            vXform = swCompXform.ArrayData

            'Root component has no name

            'Debug.Print(sPadStr & "Component = " & swComp.Name2 & " (" & swComp.ReferencedConfiguration & ")")

            'Debug.Print(sPadStr & "  Suppr   = " & swComp.IsSuppressed)

            ' Debug.Print(sPadStr & "  Hidden  = " & swComp.IsHidden(False))

            Debug.Print(sPadStr & FormatNumber(Str(vXform(0)), 4) + ", " + FormatNumber(Str(vXform(3)), 4) + ", " + FormatNumber(Str(vXform(6)), 4) + "," + FormatNumber(Str(vXform(9) * 1000), 4))
            Debug.Print(sPadStr & FormatNumber(Str(vXform(1)), 4) + ", " + FormatNumber(Str(vXform(4)), 4) + ", " + FormatNumber(Str(vXform(7)), 4) + "," + FormatNumber(Str(vXform(10) * 1000), 4))
            Debug.Print(sPadStr & FormatNumber(Str(vXform(2)), 4) + ", " + FormatNumber(Str(vXform(5)), 4) + ", " + FormatNumber(Str(vXform(8)), 4) + "," + FormatNumber(Str(vXform(11) * 1000), 4))
            Debug.Print(sPadStr & "  Scale = " + Str(vXform(12)))
            Debug.Print("")

        End If
        TransMatrix(0, 0) = FormatNumber(Str(vXform(0)), 4)
        TransMatrix(1, 0) = FormatNumber(Str(vXform(1)), 4)
        TransMatrix(2, 0) = FormatNumber(Str(vXform(2)), 4)
        TransMatrix(3, 0) = 0
        TransMatrix(0, 1) = FormatNumber(Str(vXform(3)), 4)
        TransMatrix(1, 1) = FormatNumber(Str(vXform(4)), 4)
        TransMatrix(2, 1) = FormatNumber(Str(vXform(5)), 4)
        TransMatrix(3, 1) = 0
        TransMatrix(0, 2) = FormatNumber(Str(vXform(6)), 4)
        TransMatrix(1, 2) = FormatNumber(Str(vXform(7)), 4)
        TransMatrix(2, 2) = FormatNumber(Str(vXform(8)), 4)
        TransMatrix(3, 2) = 0
        TransMatrix(0, 3) = FormatNumber(Str(vXform(9)), 4) * 1000
        TransMatrix(1, 3) = FormatNumber(Str(vXform(10)), 4) * 1000
        TransMatrix(2, 3) = FormatNumber(Str(vXform(11)), 4) * 1000
        TransMatrix(3, 3) = 1




        If MateFreeSafe = 1 AndAlso PartName <> CheckInRod Then
            NameToExcel = PartName

            Array.Copy(TransMatrix, TransMatrixPart, TransMatrixPart.Length)
            InverseMatrix = InverseTransMatrix(TransMatrix)
            If Construction IsNot Nothing Then
                TransMatrix = CalcConstruction(Construction)
                InverseMatrix = InverseTransMatrix(TransMatrix)
                NameToExcel = Construction
            End If
            WriteToArrForSaveInexcel(NameToExcel, quatPartCoordMateSide1, CalcQuat(TransMatrix, IdentityMatrix, NameToExcel), 1)

        End If

        If MateFreeSafe = 2 Then
            NameToExcel = "Mate"
            Array.Copy(TransMatrix, TransMatrixMate, TransMatrixMate.Length)
            MultiPlacMatrixMate = MultiplcationMatrixs(InverseMatrix, TransMatrixMate, NameToExcel)
            If Worientation = "o" Then
                quatPartCoordMateSide1 = CalcQuat(MultiPlacMatrixMate, MatrixForQuatTranSide1o, NameToExcel)
                quatPanelCoordMateSide1 = CalcQuat(TransMatrixMate, MatrixForQuatTranSide1o, NameToExcel)
                OperationSide = 2
                quatPartCoordMateSide2 = CalcQuat(MultiPlacMatrixMate, MatrixForQuatTranSide2o, NameToExcel)
                quatPanelCoordMateSide2 = CalcQuat(TransMatrixMate, MatrixForQuatTranSide2o, NameToExcel)
            Else
                quatPartCoordMateSide1 = CalcQuat(MultiPlacMatrixMate, IdentityMatrix, NameToExcel)
                quatPanelCoordMateSide1 = CalcQuat(TransMatrixMate, IdentityMatrix, NameToExcel)
            End If

        End If
        If MateFreeSafe = 3 Then
            NameToExcel = "Free"

            If TransMatrixFree(2, 3) > 0 Then
                TransMatrixFree(2, 3) = TransMatrixFree(2, 3)
            Else
                TransMatrixFree(2, 3) = TransMatrixFree(2, 3)
            End If

            Array.Copy(TransMatrix, TransMatrixFree, TransMatrixFree.Length)
            MultiPlacMatrixFree = MultiplcationMatrixs(InverseMatrix, TransMatrixFree, NameToExcel)
            If Worientation = "o" Then
                quatPartCoordFreeSide1 = CalcQuat(MultiPlacMatrixFree, MatrixForQuatTranSide1o, NameToExcel)
                quatPanelCoordFreeSide1 = CalcQuat(TransMatrixFree, MatrixForQuatTranSide1o, NameToExcel)
                OperationSide = 2
                quatPartCoordFreeSide2 = CalcQuat(MultiPlacMatrixFree, MatrixForQuatTranSide2o, NameToExcel)
                quatPanelCoordFreeSide2 = CalcQuat(TransMatrixFree, MatrixForQuatTranSide2o, NameToExcel)
            Else
                quatPartCoordFreeSide1 = CalcQuat(MultiPlacMatrixFree, IdentityMatrix, NameToExcel)
                quatPanelCoordFreeSide1 = CalcQuat(TransMatrixFree, IdentityMatrix, NameToExcel)
            End If

        End If
        If MateFreeSafe = 4 Then
            NameToExcel = "Safe"
            Array.Copy(TransMatrix, TransMatrixSafe, TransMatrixSafe.Length)
            If TransMatrixSafe(1, 3) >= 0 Then
                TransMatrixSafe(1, 3) = TransMatrixSafe(1, 3) + SafeDisUP 'YSafe - add 30mm for siemens
            Else
                TransMatrixSafe(1, 3) = TransMatrixSafe(1, 3) - SafeDisUP 'YSafe - add 30mm for siemens
            End If

            ThirdAxisAss = ChangeNamesForAxis(ThirdAxisAss)
            MultiPlacMatrixSafe = MultiplcationMatrixs(InverseMatrix, TransMatrixSafe, NameToExcel)
            If Worientation = "o" Then
                quatPartCoordSafeSide1 = CalcQuat(MultiPlacMatrixSafe, MatrixForQuatTranSide1o, NameToExcel)
                quatPanelCoordSafeSide1 = CalcQuat(TransMatrixSafe, MatrixForQuatTranSide1o, NameToExcel)
                OperationSide = 2
                quatPartCoordSafeSide2 = CalcQuat(MultiPlacMatrixSafe, MatrixForQuatTranSide2o, NameToExcel)
                quatPanelCoordSafeSide2 = CalcQuat(TransMatrixSafe, MatrixForQuatTranSide2o, NameToExcel)
            Else
                quatPartCoordSafeSide1 = CalcQuat(MultiPlacMatrixSafe, IdentityMatrix, NameToExcel)
                quatPanelCoordSafeSide1 = CalcQuat(TransMatrixSafe, IdentityMatrix, NameToExcel)
            End If
            WriteToArrForSaveInexcel("Mate", quatPartCoordMateSide1, quatPanelCoordMateSide1, 1)
            If Worientation = "o" Then
                WriteToArrForSaveInexcel("Mate", quatPartCoordMateSide2, quatPanelCoordMateSide2, 2)
            End If
            WriteToArrForSaveInexcel("Free", quatPartCoordFreeSide1, quatPanelCoordFreeSide1, 1)
            If Worientation = "o" Then
                WriteToArrForSaveInexcel("Free", quatPartCoordFreeSide2, quatPanelCoordFreeSide2, 2)
            End If
            WriteToArrForSaveInexcel("Safe", quatPartCoordSafeSide1, quatPanelCoordSafeSide1, 1)
            If Worientation = "o" Then
                WriteToArrForSaveInexcel("Safe", quatPartCoordSafeSide2, quatPanelCoordSafeSide2, 2)
            End If

        End If
        If Short_name <> "Panel" Then
            MateFreeSafe += 1
        End If

    End Sub
    Public Function CalcConstruction(ByVal PartName As String) As Double(,)
        Dim swComp As Component2
        Dim sPadStr As String
        Dim swCompXform As MathTransform
        Dim vXform As Object
        Dim swMathUtil As MathUtility
        Dim swModel As ModelDoc2
        Dim swModelDocExt As ModelDocExtension
        Dim swSelMgr As SelectionMgr
        Dim bRet As Boolean

        Dim IdentityMatrix(3, 3) As Double
        Dim HelpMatrixForConstruction(3, 3) As Double
        swModel = swApp.ActiveDoc
        swModelDocExt = swModel.Extension
        bRet = swModelDocExt.SelectByID2(PartName, "COMPONENT", 0, 0, 0, False, 0, Nothing, 0)

        swMathUtil = swApp.GetMathUtility
        swSelMgr = swModel.SelectionManager
        swComp = swSelMgr.GetSelectedObjectsComponent(1)

        swCompXform = swComp.Transform2

        If Not swCompXform Is Nothing Then

            vXform = swCompXform.ArrayData
        End If
        HelpMatrixForConstruction(0, 0) = FormatNumber(Str(vXform(0)), 4)
        HelpMatrixForConstruction(1, 0) = FormatNumber(Str(vXform(1)), 4)
        HelpMatrixForConstruction(2, 0) = FormatNumber(Str(vXform(2)), 4)
        HelpMatrixForConstruction(3, 0) = 0
        HelpMatrixForConstruction(0, 1) = FormatNumber(Str(vXform(3)), 4)
        HelpMatrixForConstruction(1, 1) = FormatNumber(Str(vXform(4)), 4)
        HelpMatrixForConstruction(2, 1) = FormatNumber(Str(vXform(5)), 4)
        HelpMatrixForConstruction(3, 1) = 0
        HelpMatrixForConstruction(0, 2) = FormatNumber(Str(vXform(6)), 4)
        HelpMatrixForConstruction(1, 2) = FormatNumber(Str(vXform(7)), 4)
        HelpMatrixForConstruction(2, 2) = FormatNumber(Str(vXform(8)), 4)
        HelpMatrixForConstruction(3, 2) = 0
        HelpMatrixForConstruction(0, 3) = FormatNumber(Str(vXform(9)), 4) * 1000
        HelpMatrixForConstruction(1, 3) = FormatNumber(Str(vXform(10)), 4) * 1000
        HelpMatrixForConstruction(2, 3) = FormatNumber(Str(vXform(11)), 4) * 1000
        HelpMatrixForConstruction(3, 3) = 1

        Return HelpMatrixForConstruction
    End Function
    Function CalcQuat(ByVal MartrixA(,) As Double, ByVal MatrixB(,) As Double, ByVal NameToExcel As String) As Double()
        Dim MatrixForQuat(3, 3) As Double
        Dim quatCoord(3) As Double
        Dim quatPanelCoord(3) As Double

        MatrixForQuat = MultiplcationMatrixs(MartrixA, MatrixB, NameToExcel)
        quatCoord = QUAT(MatrixForQuat)


        Return quatCoord
    End Function
    Sub SaveToExcel()
        Dim i As Integer
        Dim j As Integer
        oSheet.cells(1, 1).Value = "Catalog_No"
        oSheet.cells(1, 2).Value = "Compnent_ID"
        oSheet.cells(1, 3).Value = "Short_Name"
        oSheet.cells(1, 4).Value = "Type"
        oSheet.cells(1, 5).Value = "Identifier"
        oSheet.cells(1, 6).Value = "Point_name"
        oSheet.cells(1, 7).Value = "Designation"
        oSheet.cells(1, 8).Value = "Via (zone)"
        oSheet.cells(1, 9).Value = "Collision"
        oSheet.cells(1, 10).Value = "Fixed"
        oSheet.cells(1, 11).Value = "Coord1"
        oSheet.cells(1, 12).Value = "X [mm]"
        oSheet.cells(1, 13).Value = "Y [mm]"
        oSheet.cells(1, 14).Value = "Z [mm]"
        oSheet.cells(1, 15).Value = "W"
        oSheet.cells(1, 16).Value = "qX"
        oSheet.cells(1, 17).Value = "qY"
        oSheet.cells(1, 18).Value = "qZ"
        oSheet.cells(1, 19).Value = "Coord2"
        oSheet.cells(1, 20).Value = "X [mm]"
        oSheet.cells(1, 21).Value = "Y [mm]"
        oSheet.cells(1, 22).Value = "Z [mm]"
        oSheet.cells(1, 23).Value = "W"
        oSheet.cells(1, 24).Value = "qX"
        oSheet.cells(1, 25).Value = "qY"
        oSheet.cells(1, 26).Value = "qZ"
        oSheet.cells(1, 27).Value = "Main_Axis_Panel"
        oSheet.cells(1, 28).Value = "Secondary_Axis_Panel"
        oSheet.cells(1, 29).Value = "Height_ID"
        oSheet.cells(1, 30).Value = "Worientation"

        For i = 0 To 199
            For j = 0 To 29
                oSheet.cells(i + 2, j + 1).Value = ArrForExcel(i + 1, j)
            Next
        Next


    End Sub
    Sub WriteToArrForSaveInexcel(ByVal NameToExcel As String, ByVal quatPartCoord() As Double, quatPanelCoord() As Double, OperationSide As Integer)
        Dim Type As String
        If Construction IsNot Nothing Then
            Catalog_No = Construction.Remove(Construction.Length - 2, 2)
            Short_name = Catalog_No
        End If
        ArrForExcel(Exceli, Excelj) = Catalog_No 'Catalog name
        ArrForExcel(Exceli, Excelj + 1) = Component_ID 'id
        ArrForExcel(Exceli, Excelj + 2) = Short_name  'short name

        If NameToExcel = PartName OrElse NameToExcel = Construction Then
            Type = "CCP"
            ArrForExcel(Exceli, Excelj + 2) = Short_name
            ArrForExcel(Exceli, Excelj + 3) = Type 'port or ccp
            ArrForExcel(Exceli, Excelj + 5) = Short_name & "_" + "CCP" 'Point_name
            ArrForExcel(Exceli, Excelj + 6) = "Ref" 'Designation
            ArrForExcel(Exceli, Excelj + 7) = "Flyby" 'Via (zone)
            ArrForExcel(Exceli, Excelj + 8) = "Free" 'Collision
            ArrForExcel(Exceli, Excelj + 9) = "Keep" 'Fixed
            ArrForExcel(Exceli, Excelj + 14) = Nothing 'W
            ArrForExcel(Exceli, Excelj + 15) = Nothing 'qX
            ArrForExcel(Exceli, Excelj + 16) = Nothing 'qY
            ArrForExcel(Exceli, Excelj + 17) = Nothing 'qZ
            ArrForExcel(Exceli, Excelj + 18) = "Panel" 'Coord2
            ArrForExcel(Exceli, Excelj + 22) = quatPanelCoord(0) 'W
            ArrForExcel(Exceli, Excelj + 23) = quatPanelCoord(1) 'qX
            ArrForExcel(Exceli, Excelj + 24) = quatPanelCoord(2) 'qY
            ArrForExcel(Exceli, Excelj + 25) = quatPanelCoord(3) 'qZ
            ArrForExcel(Exceli, Excelj + 19) = TransMatrixPart(0, 3) 'X
            ArrForExcel(Exceli, Excelj + 20) = TransMatrixPart(1, 3) 'Y
            ArrForExcel(Exceli, Excelj + 21) = TransMatrixPart(2, 3) 'Z
            Exceli = Exceli + 1
        Else
            Type = "port"
            ArrForExcel(Exceli, Excelj + 3) = Type 'port or ccp
            ArrForExcel(Exceli, Excelj + 4) = PORT_num 'identifier
            ArrForExcel(Exceli, Excelj + 10) = Short_name 'Coord1
            'part
            ArrForExcel(Exceli, Excelj + 14) = quatPartCoord(0) 'W
            ArrForExcel(Exceli, Excelj + 15) = quatPartCoord(1) 'qX
            ArrForExcel(Exceli, Excelj + 16) = quatPartCoord(2) 'qY
            ArrForExcel(Exceli, Excelj + 17) = quatPartCoord(3) 'qZ
            'panel
            ArrForExcel(Exceli, Excelj + 18) = "Panel" 'Coord2 - Panel/in
            ArrForExcel(Exceli, Excelj + 22) = quatPanelCoord(0) 'W
            ArrForExcel(Exceli, Excelj + 23) = quatPanelCoord(1) 'qX
            ArrForExcel(Exceli, Excelj + 24) = quatPanelCoord(2) 'qY
            ArrForExcel(Exceli, Excelj + 25) = quatPanelCoord(3) 'qZ

            ArrForExcel(Exceli, Excelj + 26) = MainAxisAss '/Main_AXIS PANEL
                ArrForExcel(Exceli, Excelj + 27) = SecondaryAxisAss '/Secondary PANEL
                ArrForExcel(Exceli, Excelj + 28) = Height_ID
                ArrForExcel(Exceli, Excelj + 29) = Worientation 'Gripper orientation

            End If
            If NameToExcel = "Mate" Then
            ArrForExcel(Exceli, Excelj + 7) = "Fine" 'Via (zone)
            ArrForExcel(Exceli, Excelj + 8) = "Collision" 'Collision
            ArrForExcel(Exceli, Excelj + 9) = "Keep" 'Fixed
            If OperationSide = 1 AndAlso ChangeorientationForParallal <> 1 Then
                ArrForExcel(Exceli, Excelj + 5) = Short_name & "_" + MainAxisAss & "_" + Component_ID & "_" + PORT_num & "_1_in"  'Point_name
                ArrForExcel(Exceli, Excelj + 6) = NameToExcel & "_in" 'Designation
            Else
                ArrForExcel(Exceli, Excelj + 5) = Short_name & "_" + MainAxisAss & "_" + Component_ID & "_" + PORT_num & "_1_out"  'Point_name
                ArrForExcel(Exceli, Excelj + 6) = NameToExcel & "_out" 'Designation
            End If
            CoordSystem(NameToExcel)
            Exceli = Exceli + 1
        End If
        If NameToExcel = "Free" Then
            ArrForExcel(Exceli, Excelj + 7) = "Fine"
            ArrForExcel(Exceli, Excelj + 8) = "Collision"
            ArrForExcel(Exceli, Excelj + 9) = "Don't Keep"
            If OperationSide = 1 AndAlso ChangeorientationForParallal <> 1 Then
                ArrForExcel(Exceli, Excelj + 5) = Short_name & "_" + MainAxisAss & "_" + Component_ID & "_" + PORT_num & "_2_in"  'Point_name
                ArrForExcel(Exceli, Excelj + 6) = NameToExcel & "_in" 'Designation
            Else
                ArrForExcel(Exceli, Excelj + 5) = Short_name & "_" + MainAxisAss & "_" + Component_ID & "_" + PORT_num & "_2_out"  'Point_name
                ArrForExcel(Exceli, Excelj + 6) = NameToExcel & "_out" 'Designation
            End If
            CoordSystem(NameToExcel)
            Exceli = Exceli + 1
        End If
        If NameToExcel = "Safe" Then
            ArrForExcel(Exceli, Excelj + 7) = "Coarse"
            ArrForExcel(Exceli, Excelj + 8) = "Free"
            ArrForExcel(Exceli, Excelj + 9) = "Don't Keep"
            If OperationSide = 1 AndAlso ChangeorientationForParallal <> 1 Then
                ArrForExcel(Exceli, Excelj + 5) = Short_name & "_" + MainAxisAss & "_" + Component_ID & "_" + PORT_num & "_3_in"  'Point_name
                ArrForExcel(Exceli, Excelj + 6) = NameToExcel & "_in" 'Designation
            Else
                ArrForExcel(Exceli, Excelj + 5) = Short_name & "_" + MainAxisAss & "_" + Component_ID & "_" + PORT_num & "_3_out"  'Point_name
                ArrForExcel(Exceli, Excelj + 6) = NameToExcel & "_out" 'Designation
            End If
            CoordSystem(NameToExcel)
            Exceli = Exceli + 1
        End If






    End Sub
    Sub CoordSystem(ByVal NameToExcel As String)

        If NameToExcel = "Mate" Then
            ArrForExcel(Exceli, Excelj + 11) = MultiPlacMatrixMate(0, 3) 'X mate
            ArrForExcel(Exceli, Excelj + 12) = MultiPlacMatrixMate(1, 3) 'Y MATE 
            ArrForExcel(Exceli, Excelj + 13) = MultiPlacMatrixMate(2, 3) 'Z MATE

            ArrForExcel(Exceli, Excelj + 19) = TransMatrixMate(0, 3) 'X mate panel
            ArrForExcel(Exceli, Excelj + 20) = TransMatrixMate(1, 3) 'Ymate panel
            ArrForExcel(Exceli, Excelj + 21) = TransMatrixMate(2, 3) 'Zmate panel
        End If
        'X is the main axis in Part Coord 
        If NameToExcel = "Free" Then
            ArrForExcel(Exceli, Excelj + 11) = MultiPlacMatrixFree(0, 3) 'X free 
            ArrForExcel(Exceli, Excelj + 12) = MultiPlacMatrixFree(1, 3) 'Y free 
            ArrForExcel(Exceli, Excelj + 13) = MultiPlacMatrixFree(2, 3) 'Z free 

            ArrForExcel(Exceli, Excelj + 19) = TransMatrixFree(0, 3) 'X free panel
            ArrForExcel(Exceli, Excelj + 20) = TransMatrixFree(1, 3) 'Yfree panel
            ArrForExcel(Exceli, Excelj + 21) = TransMatrixFree(2, 3) 'Zfree panel

        End If
        If NameToExcel = "Safe" Then
            ArrForExcel(Exceli, Excelj + 11) = MultiPlacMatrixSafe(0, 3) 'X Safe
            ArrForExcel(Exceli, Excelj + 12) = MultiPlacMatrixSafe(1, 3) 'Y safe  
            ArrForExcel(Exceli, Excelj + 13) = MultiPlacMatrixSafe(2, 3) 'Z Safe 

            ArrForExcel(Exceli, Excelj + 19) = TransMatrixSafe(0, 3) 'X Safe panel
            ArrForExcel(Exceli, Excelj + 20) = TransMatrixSafe(1, 3) 'YSafe panel
            ArrForExcel(Exceli, Excelj + 21) = TransMatrixSafe(2, 3) 'Zsafe panel
        End If


    End Sub

    Function InverseTransMatrix(ByVal TransMatrix(,) As Double) As Double(,)
        Dim matrixInverse(,) As Double = MatrixHelper.Inverse(TransMatrix)
        Console.WriteLine(MatrixHelper.MakeDisplayable(TransMatrix) &
         vbNewLine & vbNewLine & "Inverse: " &
         vbNewLine & MatrixHelper.MakeDisplayable(matrixInverse))
        Return matrixInverse
    End Function
    Dim z As Integer = 0
    Function MultiplcationMatrixs(ByVal matrixInverse(,) As Double, ByVal TransMatrix(,) As Double, ByVal NameToExcel As String)
        Dim MultiPlacMatrix(3, 3) As Double
        Dim TempArr(15) As Double
        Dim temp As Double = 0
        For i = 0 To 3
            For j = 0 To 3
                For k = 0 To 3
                    temp = temp + (matrixInverse(i, k) * TransMatrix(k, j))
                Next
                TempArr(z) = temp
                z += 1
                temp = 0
            Next
        Next
        z = 0

        For i = 0 To 3
            For j = 0 To 3
                MultiPlacMatrix(i, j) = TempArr(z)
                z += 1
            Next
        Next
        z = 0
        Return MultiPlacMatrix
    End Function
    Function QUAT(ByVal Matrix(,) As Double) As Double()

        Dim tr As Double
        Dim S As Double
        Dim q(3) As Double
        Dim i As Integer = 0
        Dim j As Integer = 0

        tr = Matrix(0, 0) + Matrix(1, 1) + Matrix(2, 2)

        If (tr > 0) Then
            S = Math.Sqrt(tr + 1) * 2
            q(0) = FormatNumber(0.25 * S, 4)
            q(1) = FormatNumber((Matrix(2, 1) - Matrix(1, 2)) / S, 4)
            q(2) = FormatNumber((Matrix(0, 2) - Matrix(2, 0)) / S, 4)
            q(3) = FormatNumber((Matrix(1, 0) - Matrix(0, 1)) / S, 4)
        Else
            If ((Matrix(0, 0) > Matrix(1, 1)) And (Matrix(0, 0) > Matrix(2, 2))) Then
                S = Math.Sqrt(1 + Matrix(0, 0) - Matrix(1, 1) - Matrix(2, 2)) * 2
                q(0) = FormatNumber((Matrix(2, 1) - Matrix(1, 2)) / S, 4)
                q(1) = FormatNumber(0.25 * S, 4)
                q(2) = FormatNumber((Matrix(0, 1) + Matrix(1, 0)) / S, 4)
                q(3) = FormatNumber((Matrix(0, 2) + Matrix(2, 0)) / S, 4)
            Else
                If (Matrix(1, 1) > Matrix(2, 2)) Then
                    S = Math.Sqrt(1 + Matrix(1, 1) - Matrix(0, 0) - Matrix(2, 2)) * 2
                    q(0) = FormatNumber((Matrix(0, 2) - Matrix(2, 0)) / S, 4)
                    q(1) = FormatNumber(((Matrix(0, 1) + Matrix(1, 0)) / S), 4)
                    q(2) = FormatNumber(0.25 * S, 4)
                    q(3) = FormatNumber((Matrix(1, 2) + Matrix(2, 1)) / S, 4)
                Else
                    S = Math.Sqrt(1 + Matrix(2, 2) - Matrix(0, 0) - Matrix(1, 1)) * 2
                    q(0) = FormatNumber((Matrix(1, 0) - Matrix(0, 1)) / S, 4)
                    q(1) = FormatNumber((Matrix(0, 2) + Matrix(2, 0)) / S, 4)
                    q(2) = FormatNumber((Matrix(1, 2) + Matrix(2, 1)) / S, 4)
                    q(3) = FormatNumber(0.25 * S, 4)
                End If
            End If
        End If
        Return q
    End Function


    Public Sub deleteRod()
        Dim swModelDocExt As ModelDocExtension
        Dim swAssembly As AssemblyDoc
        Dim status As Boolean
        swModelDocExt = swModel.Extension
        swAssembly = swModel
        status = swModelDocExt.SelectByID2(strCompName, "COMPONENT", 0, 0, 0, False, 0, Nothing, 0)
        status = swAssembly.DeleteSelections(swAssemblyDeleteOptions_e.swDelete_SelectedComponents)
    End Sub

    ''' <summary>
    ''' The SldWorks swApp variable is pre-assigned for you.
    ''' </summary>
    Public swApp As SldWorks


End Class

