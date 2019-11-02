'Imports Autodesk.Revit
'Imports Autodesk.Revit.DB
'Imports Autodesk.Revit.UI
''Imports Microsoft.WindowsAPICodePack
'Partial Public Class clsCR
'    '
'    Public Shared colComandos As Dictionary(Of String, UI.RevitCommandId)
'    Public Shared arrComandos As String() = New String() {
'        "ID_REVIT_FILE_SAVE_AS", "ID_REVIT_SAVE_AS_FAMILY", "ID_REVIT_SAVE_AS_TEMPLATE",
'        "ID_REVIT_FILE_SAVE",
'        "ID_REVIT_FILE_OPEN",
'    "ID_FAMILY_LOAD"}
'    '"ID_REVIT_FILE_SAVE_AS",
'    '"ID_REVIT_FILE_OPEN", "ID_APPMENU_PROJECT_OPEN", "ID_FAMILY_OPEN", "ID_IMPORT_IFC", "ID_OPEN_ADSK",
'    '"ID_FAMILY_LOAD"}
'    'ID_APP_EXIT
'    Public Shared Sub Comandos_RellenaId()
'        colComandos = New Dictionary(Of String, UI.RevitCommandId)
'        For Each s_commandtodisable As String In arrComandos
'            Dim s_commandId As UI.RevitCommandId = Autodesk.Revit.UI.RevitCommandId.LookupCommandId(s_commandtodisable)
'            If s_commandId Is Nothing OrElse (s_commandId IsNot Nothing AndAlso s_commandId.CanHaveBinding = False) Then
'                colComandos.Add(s_commandtodisable, Nothing)
'            ElseIf s_commandId IsNot Nothing Then
'                colComandos.Add(s_commandtodisable, s_commandId)
'            End If
'        Next
'        ' Ahora reasignaremos cada comando (Comentarlo temporalmente para que no reasigne)
'        Comandos_Reasignar()
'    End Sub
'    Public Shared Sub Comandos_Reasignar()
'        For Each s_command As String In colComandos.Keys
'            If colComandos(s_command) Is Nothing Then Continue For
'            '
'            Try
'                Dim commandbinding As Autodesk.Revit.UI.AddInCommandBinding = crUIAppCont.CreateAddInCommandBinding(colComandos(s_command))
'                Select Case s_command
'                    Case "ID_REVIT_FILE_SAVE_AS"
'                        'saveFilter = New Microsoft.WindowsAPICodePack.Dialogs.CommonFileDialogFilter("Revit Project", "*.rvt")
'                        AddHandler commandbinding.Executed, AddressOf clsCR.Comando_GuardarComoRVT
'                    Case "ID_REVIT_SAVE_AS_FAMILY"
'                        'saveFilter = New Microsoft.WindowsAPICodePack.Dialogs.CommonFileDialogFilter("Revit Family", "*.rfa")
'                        AddHandler commandbinding.Executed, AddressOf clsCR.Comando_GuardarComoRFA
'                    Case "ID_REVIT_SAVE_AS_TEMPLATE"
'                        'saveFilter = New Microsoft.WindowsAPICodePack.Dialogs.CommonFileDialogFilter("Revit Project Template", "*.rte")
'                        AddHandler commandbinding.Executed, AddressOf clsCR.Comando_GuardarComoRTE
'                    Case "ID_REVIT_FILE_SAVE"
'                        'saveFilter = New Microsoft.WindowsAPICodePack.Dialogs.CommonFileDialogFilter("Revit Project Template", "*.rte")
'                        AddHandler commandbinding.Executed, AddressOf clsCR.Comando_GuardarRFA
'                    Case "ID_REVIT_FILE_OPEN"
'                        AddHandler commandbinding.Executed, AddressOf clsCR.Comando_AbrirRevitFile
'                    'Case "ID_APPMENU_PROJECT_OPEN"
'                    '    AddHandler commandbinding.Executed, AddressOf clsCR.Comando_AbrirProyecto
'                    'Case "ID_FAMILY_OPEN"
'                    '    AddHandler commandbinding.Executed, AddressOf clsCR.Comando_AbrirFamilia
'                    Case "ID_FAMILY_LOAD"
'                        AddHandler commandbinding.Executed, AddressOf clsCR.Comando_CargarFamilia
'                        'Case "ID_IMPORT_IFC"
'                        '    AddHandler commandbinding.Executed, AddressOf clsCR.Comando_AbrirIFC
'                        'Case "ID_OPEN_ADSK"
'                        '    AddHandler commandbinding.Executed, AddressOf clsCR.Comando_AbrirADSK
'                        '
'                End Select
'            Catch ex As Exception
'                Continue For
'            End Try
'        Next
'    End Sub
'    '
'    Public Shared Sub Comandos_QuitarAsignacion()
'        If colComandos Is Nothing Then Exit Sub
'        For Each s_command As String In colComandos.Keys
'            If colComandos(s_command) Is Nothing Then Continue For
'            '
'            Try
'                crUIAppCont.RemoveAddInCommandBinding(colComandos(s_command))
'            Catch ex As Exception

'            End Try
'        Next
'    End Sub
'    '
'    Public Shared Sub Comando_GuardarComoRVT()
'        clsCR.crPuedoborrar = False
'        Dim fullPath As String = crUIApp.ActiveUIDocument.Document.PathName
'        Dim directorio As String = ""
'        Dim nombre As String = ""
'        Dim extension As String = ".rvt"
'        '
'        If fullPath = "" Then
'            directorio = IIf(crfolder_ultimo = "", crfolder_MisDoc, crfolder_ultimo)
'            nombre = ""
'        ElseIf IO.File.Exists(fullPath) Then
'            fullPath = Fichero_DameDesdeKey(fullPath, False)
'            directorio = IO.Path.GetDirectoryName(fullPath)
'            nombre = IO.Path.GetFileNameWithoutExtension(fullPath)
'        End If
'        '
'REPITE:
'        Dim correcto As Microsoft.WindowsAPICodePack.Dialogs.CommonFileDialogResult = Microsoft.WindowsAPICodePack.Dialogs.CommonFileDialogResult.Cancel
'        Dim fiFinal As String = ""
'        If Microsoft.WindowsAPICodePack.Dialogs.CommonFileDialog.IsPlatformSupported Then
'            Using oSaveFileDialog As New Microsoft.WindowsAPICodePack.Dialogs.CommonSaveFileDialog
'                oSaveFileDialog.Title = "ULMA - Save as Revit Project"
'                oSaveFileDialog.ShowPlacesList = True
'                oSaveFileDialog.AddToMostRecentlyUsedList = True
'                oSaveFileDialog.Filters.Add(New Microsoft.WindowsAPICodePack.Dialogs.CommonFileDialogFilter("Revit Project", extension))
'                oSaveFileDialog.AlwaysAppendDefaultExtension = True
'                oSaveFileDialog.EnsureReadOnly = True
'                oSaveFileDialog.InitialDirectory = directorio
'                oSaveFileDialog.DefaultExtension = extension
'                oSaveFileDialog.DefaultFileName = nombre
'                oSaveFileDialog.DefaultDirectory = directorio
'                oSaveFileDialog.OverwritePrompt = False
'                If oSaveFileDialog.ShowDialog(Autodesk.Windows.ComponentManager.ApplicationWindow) = Microsoft.WindowsAPICodePack.Dialogs.CommonFileDialogResult.Ok Then
'                    fiFinal = oSaveFileDialog.FileName    ' IO.Path.GetDirectoryName(folderselectorDialog.FileName)
'                    If fiFinal.StartsWith(crfolder_TempRaiz) = True Then    ' Or crDicFicheros.ContainsValue(fiFinal) Then
'                        correcto = Microsoft.WindowsAPICodePack.Dialogs.CommonFileDialogResult.None
'                    Else
'                        correcto = Microsoft.WindowsAPICodePack.Dialogs.CommonFileDialogResult.Ok
'                    End If
'                Else
'                    Exit Sub
'                End If
'            End Using
'            '
'            If correcto = Microsoft.WindowsAPICodePack.Dialogs.CommonFileDialogResult.None Then
'                Autodesk.Revit.UI.TaskDialog.Show("ERROR", "Folder not allowed to save Revit files")    ' / Or File is the same, use Save")
'                GoTo REPITE
'            ElseIf correcto = Microsoft.WindowsAPICodePack.Dialogs.CommonFileDialogResult.Ok And fiFinal <> "" Then
'                If fiFinal.EndsWith(extension) = False Then fiFinal &= extension
'                'Dim oSOpt As New SaveAsOptions
'                'oSOpt.Compact = True
'                'oSOpt.MaximumBackups = 1
'                If IO.File.Exists(fiFinal) Then
'                    crMuestraDialogos = True
'                    If Autodesk.Revit.UI.TaskDialog.Show("NOTICE", "Overwrite existing file?", TaskDialogCommonButtons.Yes Or TaskDialogCommonButtons.No) = Autodesk.Revit.UI.TaskDialogResult.No Then
'                        GoTo REPITE
'                    End If
'                End If
'                clsCR.crEventoCerrar = False
'                clsCR.crEventoAbrir = False
'                '
'                IO.File.Copy(crUIApp.ActiveUIDocument.Document.PathName, fiFinal, True)
'                Dim optSaveAs As New SaveAsOptions
'                optSaveAs.MaximumBackups = 1
'                optSaveAs.OverwriteExistingFile = True
'                'crfiAbrir = Fichero_PonNuevo(fiFinal)
'                'crfiCerrar = crDicFicheros(crfiAbrir)
'                crUIApp.ActiveUIDocument.Document.SaveAs(fiFinal, optSaveAs)
'                '
'                clsCR.crEventoAbrir = True
'                clsCR.crEventoCerrar = True
'            End If
'        End If
'        clsCR.crPuedoborrar = True
'    End Sub
'    '
'    Public Shared Sub Comando_GuardarComoRFA()
'        Dim fullPath As String = crUIApp.ActiveUIDocument.Document.PathName
'        Dim directorio As String = crfolder_library
'        Dim nombre As String = crUIApp.ActiveUIDocument.Document.Title
'        Dim extension As String = ".rfa"
'        '
'        If IO.File.Exists(fullPath) Then
'            nombre = IO.Path.GetFileNameWithoutExtension(fullPath)
'            If IO.Directory.Exists(directorio) = False Then
'                directorio = IO.Path.GetDirectoryName(fullPath)
'            End If
'        Else
'            If IO.Directory.Exists(directorio) = False Then
'                directorio = crfolder_MisDoc
'            End If
'        End If
'        '
'REPITE:
'        Dim correcto As Microsoft.WindowsAPICodePack.Dialogs.CommonFileDialogResult = Nothing
'        Dim fiFinal As String = ""
'        If Microsoft.WindowsAPICodePack.Dialogs.CommonFileDialog.IsPlatformSupported Then
'            Using oSaveFileDialog As New Microsoft.WindowsAPICodePack.Dialogs.CommonSaveFileDialog
'                oSaveFileDialog.Title = "ULMA - Save as Revit Family"
'                oSaveFileDialog.ShowPlacesList = True
'                oSaveFileDialog.AddToMostRecentlyUsedList = True
'                oSaveFileDialog.Filters.Add(New Microsoft.WindowsAPICodePack.Dialogs.CommonFileDialogFilter("Revit Family", extension))
'                oSaveFileDialog.AlwaysAppendDefaultExtension = True
'                oSaveFileDialog.EnsureReadOnly = True
'                oSaveFileDialog.InitialDirectory = directorio
'                oSaveFileDialog.DefaultExtension = extension
'                oSaveFileDialog.DefaultFileName = nombre
'                oSaveFileDialog.DefaultDirectory = directorio
'                oSaveFileDialog.OverwritePrompt = False
'                If oSaveFileDialog.ShowDialog(Autodesk.Windows.ComponentManager.ApplicationWindow) = Microsoft.WindowsAPICodePack.Dialogs.CommonFileDialogResult.Ok Then
'                    fiFinal = oSaveFileDialog.FileName    ' IO.Path.GetDirectoryName(folderselectorDialog.FileName)
'                    If fiFinal.StartsWith(crfolder_TempRaiz) = True Then
'                        correcto = Microsoft.WindowsAPICodePack.Dialogs.CommonFileDialogResult.None
'                    Else
'                        correcto = Microsoft.WindowsAPICodePack.Dialogs.CommonFileDialogResult.Ok
'                    End If
'                Else
'                    Exit Sub
'                End If
'            End Using
'            '
'            If correcto = Microsoft.WindowsAPICodePack.Dialogs.CommonFileDialogResult.None Then
'                Autodesk.Revit.UI.TaskDialog.Show("ERROR", "Folder not allowed to save Revit files...")
'                GoTo REPITE
'            ElseIf correcto = Microsoft.WindowsAPICodePack.Dialogs.CommonFileDialogResult.Ok And fiFinal <> "" Then
'                If fiFinal.EndsWith(extension) = False Then fiFinal &= extension
'                Dim oSOpt As New SaveAsOptions
'                oSOpt.Compact = True
'                oSOpt.MaximumBackups = 1
'                If IO.File.Exists(fiFinal) Then
'                    'AddHandler uiContApp.DialogBoxShowing, AddressOf ApplicationUIEvent_DialogBoxShowing_Handler
'                    If Autodesk.Revit.UI.TaskDialog.Show("NOTICE", "Overwrite existing file?", TaskDialogCommonButtons.Yes Or TaskDialogCommonButtons.No) = Autodesk.Revit.UI.TaskDialogResult.Yes Then
'                        oSOpt.OverwriteExistingFile = True
'                        crUIApp.ActiveUIDocument.Document.SaveAs(fiFinal, oSOpt)
'                    Else
'                        GoTo REPITE
'                    End If
'                    'RemoveHandler uiContApp.DialogBoxShowing, AddressOf ApplicationUIEvent_DialogBoxShowing_Handler
'                Else
'                    crUIApp.ActiveUIDocument.Document.SaveAs(fiFinal, oSOpt)
'                End If
'                clsCR.crPuedoborrar = False
'                crip2aCAD.clsCR.Fichero_Encripta(fiFinal, crSaveImagePreview, True)
'                crip2aCAD.clsCR.crfiAbrir = Fichero_PonNuevo(fiFinal)
'                crip2aCAD.clsCR.crfiCerrar = fiFinal
'                'AddHandler crUIAppCont.Idling, AddressOf ApplicationUIEvent_Idling_AbreTrasSaveAs
'                clsCR.crPuedoborrar = True
'            End If
'        End If
'    End Sub
'    Public Shared Sub Comando_GuardarComoRTE()
'        Dim fullPath As String = crUIApp.ActiveUIDocument.Document.PathName
'        Dim directorio As String = crfolder_templates
'        Dim nombre As String = ""
'        Dim extension As String = ".rte"
'        '
'        If IO.File.Exists(fullPath) Then
'            nombre = IO.Path.GetFileNameWithoutExtension(fullPath)
'            If IO.Directory.Exists(directorio) = False Then
'                directorio = IO.Path.GetDirectoryName(fullPath)
'            End If
'        Else
'            nombre = ""
'            If IO.Directory.Exists(directorio) = False Then
'                directorio = crfolder_MisDoc
'            End If
'        End If
'        '
'REPITE:
'        Dim correcto As Microsoft.WindowsAPICodePack.Dialogs.CommonFileDialogResult = Nothing
'        Dim fiFinal As String = ""
'        If Microsoft.WindowsAPICodePack.Dialogs.CommonFileDialog.IsPlatformSupported Then
'            Using oSaveFileDialog As New Microsoft.WindowsAPICodePack.Dialogs.CommonSaveFileDialog
'                oSaveFileDialog.Title = "ULMA - Save as Revit Project Template"
'                oSaveFileDialog.ShowPlacesList = True
'                oSaveFileDialog.AddToMostRecentlyUsedList = True
'                oSaveFileDialog.Filters.Add(New Microsoft.WindowsAPICodePack.Dialogs.CommonFileDialogFilter("Revit Project Template", extension))
'                oSaveFileDialog.AlwaysAppendDefaultExtension = True
'                oSaveFileDialog.EnsureReadOnly = True
'                oSaveFileDialog.InitialDirectory = directorio
'                oSaveFileDialog.DefaultExtension = extension
'                oSaveFileDialog.DefaultFileName = nombre
'                oSaveFileDialog.DefaultDirectory = directorio
'                oSaveFileDialog.OverwritePrompt = False
'                If oSaveFileDialog.ShowDialog(Autodesk.Windows.ComponentManager.ApplicationWindow) = Microsoft.WindowsAPICodePack.Dialogs.CommonFileDialogResult.Ok Then
'                    fiFinal = oSaveFileDialog.FileName    ' IO.Path.GetDirectoryName(folderselectorDialog.FileName)
'                    If fiFinal.StartsWith(crfolder_TempRaiz) = True Then
'                        correcto = Microsoft.WindowsAPICodePack.Dialogs.CommonFileDialogResult.None
'                    Else
'                        correcto = Microsoft.WindowsAPICodePack.Dialogs.CommonFileDialogResult.Ok
'                    End If
'                Else
'                    Exit Sub
'                End If
'            End Using
'            '
'            If correcto = Microsoft.WindowsAPICodePack.Dialogs.CommonFileDialogResult.None Then
'                Autodesk.Revit.UI.TaskDialog.Show("ERROR", "Folder not allowed to save Revit files...")
'                GoTo REPITE
'            ElseIf correcto = Microsoft.WindowsAPICodePack.Dialogs.CommonFileDialogResult.Ok And fiFinal <> "" Then
'                If fiFinal.EndsWith(extension) = False Then fiFinal &= extension
'                Dim oSOpt As New SaveAsOptions
'                oSOpt.Compact = True
'                oSOpt.MaximumBackups = 1
'                If IO.File.Exists(fiFinal) Then
'                    If Autodesk.Revit.UI.TaskDialog.Show("NOTICE", "Overwrite existing file?", TaskDialogCommonButtons.Yes Or TaskDialogCommonButtons.No) = Autodesk.Revit.UI.TaskDialogResult.Yes Then
'                        oSOpt.OverwriteExistingFile = True
'                        crUIApp.ActiveUIDocument.Document.SaveAs(fiFinal, oSOpt)
'                    Else
'                        GoTo REPITE
'                    End If
'                Else
'                    crUIApp.ActiveUIDocument.Document.SaveAs(fiFinal, oSOpt)
'                End If
'                crip2aCAD.clsCR.crfiAbrir = Fichero_PonNuevo(fiFinal)
'                crip2aCAD.clsCR.crfiCerrar = fiFinal
'                'AddHandler crUIAppCont.Idling, AddressOf ApplicationUIEvent_Idling_AbreTrasSaveAs
'            End If
'        End If
'    End Sub
'    '
'    Public Shared Sub Comando_GuardarRFA()
'        Dim fullPath As String = crUIApp.ActiveUIDocument.Document.PathName
'        ' Solo para ficheros de familia
'        If crUIApp.ActiveUIDocument.Document.IsFamilyDocument = False Then
'            crUIApp.ActiveUIDocument.Document.Save()
'            Exit Sub
'        End If
'        If fullPath.StartsWith(crfolder_Temp) = False Then
'            crUIApp.ActiveUIDocument.Document.Save()
'            Exit Sub
'        End If
'        'If fullPath = "" Then
'        '    fullPath = IO.Path.Combine(folder_TempRevit, uiApp.ActiveUIDocument.Document.Title)
'        'End If
'        '
'        Dim directorio As String = ""
'        Dim nombre As String = ""
'        Dim extension As String = ".rfa"
'        ' Esta en TEMP. Coger el que tengamos en la clase
'        fullPath = Fichero_DameDesdeKey(fullPath, False)
'        '
'        If fullPath = "" Then
'            directorio = IIf(crfolder_library = "", FolderLibrary_BuscaDame(crUIAppCont), crfolder_library)
'            nombre = crUIApp.ActiveUIDocument.Document.Title
'        ElseIf IO.File.Exists(fullPath) Then
'            fullPath = Fichero_DameDesdeKey(fullPath, False)
'            directorio = IO.Path.GetDirectoryName(fullPath)
'            nombre = IO.Path.GetFileNameWithoutExtension(fullPath)
'        End If
'        '
'REPITE:
'        Dim correcto As Microsoft.WindowsAPICodePack.Dialogs.CommonFileDialogResult = Microsoft.WindowsAPICodePack.Dialogs.CommonFileDialogResult.None
'        Dim fiFinal As String = ""
'        If Microsoft.WindowsAPICodePack.Dialogs.CommonFileDialog.IsPlatformSupported Then
'            Using oSaveFileDialog As New Microsoft.WindowsAPICodePack.Dialogs.CommonSaveFileDialog
'                oSaveFileDialog.Title = "ULMA - Save as Revit Family"
'                oSaveFileDialog.ShowPlacesList = True
'                oSaveFileDialog.AddToMostRecentlyUsedList = True
'                oSaveFileDialog.Filters.Add(New Microsoft.WindowsAPICodePack.Dialogs.CommonFileDialogFilter("Revit Family", extension))
'                oSaveFileDialog.AlwaysAppendDefaultExtension = True
'                oSaveFileDialog.EnsureReadOnly = True
'                oSaveFileDialog.InitialDirectory = directorio
'                oSaveFileDialog.DefaultExtension = extension
'                oSaveFileDialog.DefaultFileName = nombre
'                oSaveFileDialog.DefaultDirectory = directorio
'                oSaveFileDialog.OverwritePrompt = False
'                If oSaveFileDialog.ShowDialog(Autodesk.Windows.ComponentManager.ApplicationWindow) = Microsoft.WindowsAPICodePack.Dialogs.CommonFileDialogResult.Ok Then
'                    fiFinal = oSaveFileDialog.FileName    ' IO.Path.GetDirectoryName(folderselectorDialog.FileName)
'                    If fiFinal.StartsWith(crfolder_Temp) = True Then
'                        correcto = Microsoft.WindowsAPICodePack.Dialogs.CommonFileDialogResult.None
'                    Else
'                        correcto = Microsoft.WindowsAPICodePack.Dialogs.CommonFileDialogResult.Ok
'                    End If
'                Else
'                    Exit Sub
'                End If
'            End Using
'            '
'            If correcto = Microsoft.WindowsAPICodePack.Dialogs.CommonFileDialogResult.None Then
'                Autodesk.Revit.UI.TaskDialog.Show("ERROR", "Folder not allowed to save Revit files...")
'                GoTo REPITE
'            ElseIf correcto = Microsoft.WindowsAPICodePack.Dialogs.CommonFileDialogResult.Ok And fiFinal <> "" Then
'                If fiFinal.EndsWith(extension) = False Then fiFinal &= extension
'                Dim oSOpt As New SaveAsOptions
'                oSOpt.Compact = True
'                oSOpt.MaximumBackups = 1
'                If IO.File.Exists(fiFinal) Then
'                    If Autodesk.Revit.UI.TaskDialog.Show("NOTICE", "Overwrite existing file?", TaskDialogCommonButtons.Yes Or TaskDialogCommonButtons.No) = Autodesk.Revit.UI.TaskDialogResult.Yes Then
'                        Try
'                            IO.File.Delete(fiFinal)
'                        Catch ex As Exception

'                        End Try
'                        '
'                        oSOpt.OverwriteExistingFile = True
'                        crUIApp.ActiveUIDocument.Document.SaveAs(fiFinal, oSOpt)
'                    Else
'                        GoTo REPITE
'                    End If
'                Else
'                    crUIApp.ActiveUIDocument.Document.SaveAs(fiFinal, oSOpt)
'                End If
'                clsCR.crPuedoborrar = False
'                ' Cerrar el documento actual y abrir el de temp.
'                'Fichero_EncriptaConBackups(fiFinal)
'                crfiAbrir = Fichero_PonNuevo(fiFinal, True)
'                Fichero_Encripta(fiFinal, False)
'                'FicherosDesarrollo_BorraTempTASK()
'                clsCR.crPuedoborrar = True
'            End If
'        End If
'    End Sub
'    '
'    Public Shared Sub Comando_AbrirRevitFile(ByVal sender As Object, ByVal e As EventArgs)
'        If crUIApp Is Nothing Then crUIApp = sender
'        Dim directorio As String = crfolder_ultimo ' My.Computer.FileSystem.SpecialDirectories.MyDocuments
'        If IO.Directory.Exists(directorio) = False Then
'            directorio = My.Computer.FileSystem.SpecialDirectories.MyDocuments
'        End If
'        Dim nombre As String = ""       ' Inicialmente, sin nombre
'        Dim extension1 As String = ".rvt"       ' Extension1 que podemos cargar
'        Dim extension2 As String = ".rfa"       ' Extension1 que podemos cargar
'        Dim extension3 As String = ".rte"       ' Extension1 que podemos cargar
'        Dim extension4 As String = ".rft"       ' Extension1 que podemos cargar
'        Dim extension5 As String = ".adsk"      ' Extension1 que podemos cargar
'        Dim extension6 As String = ".ifc"       ' Extension1 que podemos cargar
'        '
'REPITE:
'        Dim correcto As Microsoft.WindowsAPICodePack.Dialogs.CommonFileDialogResult
'        Dim fiFinal As String = ""
'        Dim bolInserta As Boolean = False
'        If Microsoft.WindowsAPICodePack.Dialogs.CommonFileDialog.IsPlatformSupported Then
'            Using oOpenFileDialog As New Microsoft.WindowsAPICodePack.Dialogs.CommonOpenFileDialog
'                oOpenFileDialog.Title = "ULMA - Open Revit File"
'                oOpenFileDialog.ShowPlacesList = True
'                oOpenFileDialog.AddToMostRecentlyUsedList = True
'                oOpenFileDialog.Filters.Add(New Microsoft.WindowsAPICodePack.Dialogs.CommonFileDialogFilter("Revit Project File", extension1))
'                oOpenFileDialog.Filters.Add(New Microsoft.WindowsAPICodePack.Dialogs.CommonFileDialogFilter("Revit Family File", extension2))
'                oOpenFileDialog.Filters.Add(New Microsoft.WindowsAPICodePack.Dialogs.CommonFileDialogFilter("Revit Template File", extension3))
'                oOpenFileDialog.Filters.Add(New Microsoft.WindowsAPICodePack.Dialogs.CommonFileDialogFilter("Revit Template Family File", extension4))
'                oOpenFileDialog.Filters.Add(New Microsoft.WindowsAPICodePack.Dialogs.CommonFileDialogFilter("Autodesk adsk file", extension5))
'                oOpenFileDialog.Filters.Add(New Microsoft.WindowsAPICodePack.Dialogs.CommonFileDialogFilter("Autodesk ifc file", extension6))
'                'oOpenFileDialog.DefaultExtension = extension1
'                oOpenFileDialog.EnsureFileExists = True
'                oOpenFileDialog.EnsurePathExists = True
'                oOpenFileDialog.EnsureReadOnly = True
'                oOpenFileDialog.InitialDirectory = directorio
'                oOpenFileDialog.DefaultFileName = nombre
'                oOpenFileDialog.DefaultDirectory = directorio
'                ' CheckBox para insertar o no la familia cargada / A plano o cara
'                'Dim cbInsert As New Microsoft.WindowsAPICodePack.Dialogs.Controls.CommonFileDialogCheckBox("Insert First FamilySymbol", True)
'                'oOpenFileDialog.Controls.Add(cbInsert)
'                '
'                If oOpenFileDialog.ShowDialog(Autodesk.Windows.ComponentManager.ApplicationWindow) = Microsoft.WindowsAPICodePack.Dialogs.CommonFileDialogResult.Ok Then
'                    fiFinal = oOpenFileDialog.FileName    ' IO.Path.GetDirectoryName(folderselectorDialog.FileName)
'                    If fiFinal.StartsWith(crfolder_Temp) = True Then
'                        correcto = Microsoft.WindowsAPICodePack.Dialogs.CommonFileDialogResult.None
'                        crMuestraDialogos = True
'                        TaskDialog.Show("NOTICE TO THE USER", "Folder not allowed to work with Revit files", TaskDialogCommonButtons.Close)
'                        crMuestraDialogos = False
'                        GoTo REPITE
'                    Else
'                        correcto = Microsoft.WindowsAPICodePack.Dialogs.CommonFileDialogResult.Ok
'                    End If
'                End If
'            End Using
'        End If
'        ' Abrir el documento.
'        If IO.File.Exists(fiFinal) And correcto = Microsoft.WindowsAPICodePack.Dialogs.CommonFileDialogResult.Ok Then
'            clsCR.crfiAbrir = fiFinal
'            ' Siempre desencriptar el fichero.
'            crip2aCAD.clsCR.Fichero_Desencripta(crfiAbrir)
'            If crip2aCAD.clsCR.crEncrypt = True Then
'                crfiAbrir = crip2aCAD.clsCR.Fichero_PonNuevo(crfiAbrir, True)
'                crip2aCAD.clsCR.Fichero_Desencripta(crfiAbrir)
'                'crip2aCAD.clsCR.Fichero_DesencriptaAbreTemp(fiAbrir)
'            End If
'            clsCR.crfolder_ultimo = IO.Path.GetDirectoryName(Fichero_DameDesdeKey(clsCR.crfiAbrir))
'            '
'            'Try
'            '
'            Fichero_AbreYActiva(crfiAbrir)
'            '
'            'Catch ex As Exception
'            '    oDoc = oApp.OpenDocumentFile(fiAbrir)
'            '    If oDoc IsNot Nothing Then
'            '        Dim v3D As View3D = View3D_Dame(oDoc)
'            '        Dim uiDoc As UIDocument = New UIDocument(oDoc)
'            '        Dim collector As New FilteredElementCollector(oDoc, v3D.Id) ', uiDoc.ActiveGraphicalView.Id)
'            '        collector.OfClass(GetType(FamilyInstance))
'            '        Dim lstElement As List(Of ElementId) = collector.ToList.Select(Function(x) x.Id).ToList()
'            '        '
'            '        uiDoc.ShowElements(lstElement)
'            '        uiDoc.RefreshActiveView()
'            '        'uiDoc.ActiveView = v3D
'            '        View3D_ActivaCierraOtras(oDoc)
'            '    End If
'            'End Try
'            crfiAbrir = ""
'            crfiCerrar = ""
'            'If encrypt = True Then
'            '    fiAbrir = Fichero_DesencriptaEnTempDame(fiFinal)
'            'Else
'            '    Fichero_Desencripta(fiAbrir, False)
'            'End If
'            'Try
'            '    uiDoc = uiApp.OpenAndActivateDocument(fiAbrir)
'            '    If uiDoc IsNot Nothing Then oDoc = uiDoc.Document
'            'Catch ex As Exception
'            '    Debug.Print(ex.ToString)
'            'End Try
'        End If
'    End Sub

'    Public Shared Sub Comando_CargarFamilia()  'ByVal sender As Object, ByVal e As EventArgs)
'        Dim fiActual As String = crUIApp.ActiveUIDocument.Document.PathName   ' Fichero abierto actualmente. Para cargar la familia en él.
'        Dim directorio As String = crfolder_library                         ' Carpeta de libreria inicial para mostrar.
'        If IO.Directory.Exists(directorio) = False Then
'            directorio = clsCR.Fichero_DameDesdeKey(fiActual, True)
'        End If
'        Dim nombre As String = ""       ' Inicialmente, sin nombre
'        Dim extension1 As String = ".rfa"       ' Extension1 que podemos cargar RFA
'        Dim extension2 As String = ".adsk"      ' Extensión2 que podemos cargar ADSK
'        '
'REPITE:
'        Dim correcto As Boolean = False
'        Dim fiFinal As String = ""
'        Dim bolInserta As Boolean = False
'        If Microsoft.WindowsAPICodePack.Dialogs.CommonFileDialog.IsPlatformSupported Then
'            Using oOpenFileDialog As New Microsoft.WindowsAPICodePack.Dialogs.CommonOpenFileDialog
'                oOpenFileDialog.Title = "ULMA - Load Revit Family"
'                oOpenFileDialog.ShowPlacesList = True
'                oOpenFileDialog.AddToMostRecentlyUsedList = True
'                oOpenFileDialog.Filters.Add(New Microsoft.WindowsAPICodePack.Dialogs.CommonFileDialogFilter("Revit Family", extension1))
'                oOpenFileDialog.Filters.Add(New Microsoft.WindowsAPICodePack.Dialogs.CommonFileDialogFilter("Autodesk adsk file", extension2))
'                oOpenFileDialog.DefaultExtension = extension1
'                oOpenFileDialog.EnsureFileExists = True
'                oOpenFileDialog.EnsurePathExists = True
'                oOpenFileDialog.EnsureReadOnly = True
'                oOpenFileDialog.InitialDirectory = directorio
'                oOpenFileDialog.DefaultFileName = nombre
'                oOpenFileDialog.DefaultDirectory = directorio
'                ' CheckBox para insertar o no la familia cargada / A plano o cara
'                Dim cbInsert As New Microsoft.WindowsAPICodePack.Dialogs.Controls.CommonFileDialogCheckBox("Insert First FamilySymbol", True)
'                oOpenFileDialog.Controls.Add(cbInsert)
'                '
'                If oOpenFileDialog.ShowDialog(Autodesk.Windows.ComponentManager.ApplicationWindow) = Microsoft.WindowsAPICodePack.Dialogs.CommonFileDialogResult.Ok Then
'                    fiFinal = oOpenFileDialog.FileName    ' IO.Path.GetDirectoryName(folderselectorDialog.FileName)
'                    bolInserta = cbInsert.IsChecked
'                End If
'            End Using
'        End If
'        ' Cargar e insertar la familia elegida
'        If IO.File.Exists(fiFinal) Then
'            If crip2aCAD.clsCR.crEncrypt = True Then
'                crip2aCAD.clsCR.crfiAbrir = crip2aCAD.clsCR.Fichero_DesencriptaEnTempDame(fiFinal)
'                If crip2aCAD.clsCR.crfiAbrir = "" OrElse IO.File.Exists(crip2aCAD.clsCR.crfiAbrir) = False Then
'                    crip2aCAD.clsCR.crfiAbrir = ""
'                Else
'                    Call Family_Cargar(pathFamily:=crfiAbrir, insertar:=bolInserta, borrarDespues:=True)
'                End If
'            End If
'        End If
'    End Sub

'    Public Shared Sub Comando_AbrirIFC()  'ByVal sender As Object, ByVal e As EventArgs)
'        'Dim oOpen As New System.Windows.Forms.OpenFileDialog
'        'oOpen.Filter = "Open IFC (*.ifc)|*.ifc"
'        'oOpen.Title = "UCREVIT - Open IFC (*.ifc)"
'        'oOpen.FilterIndex = 1
'        'If oOpen.ShowDialog = System.Windows.Forms.DialogResult.OK Then
'        '    MostrarMensaje("CORRECT", "Abriría un fichero --> " + oOpen.DefaultExt)
'        'End If
'    End Sub


'    Public Shared Sub Comando_AbrirADSK()  'ByVal sender As Object, ByVal e As EventArgs)
'        'Dim oOpen As New System.Windows.Forms.OpenFileDialog
'        'oOpen.Filter = "Open ADSKC (*.adsk)|*.adsk"
'        'oOpen.Title = "UCREVIT - Open ADSK"
'        'oOpen.FilterIndex = 1
'        'If oOpen.ShowDialog = System.Windows.Forms.DialogResult.OK Then
'        '    MostrarMensaje("CORRECT", "Abriría un fichero --> " + oOpen.DefaultExt)
'        'End If
'    End Sub
'End Class
