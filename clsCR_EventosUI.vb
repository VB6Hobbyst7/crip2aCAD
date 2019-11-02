Imports System
Imports System.Timers
Imports System.Text
Imports System.Drawing
Imports System.Security
Imports System.Security.Cryptography
Imports System.Xml
Imports System.Linq
Imports System.Reflection
Imports System.ComponentModel
Imports System.Collections
Imports System.Collections.Generic
Imports System.Windows
Imports System.Windows.Forms
Imports System.Windows.Media.Imaging
Imports System.IO
'
Imports Autodesk.Revit
Imports Autodesk.Revit.ApplicationServices
Imports Autodesk.Revit.Attributes
Imports Autodesk.Revit.DB
Imports Autodesk.Revit.DB.Events
Imports Autodesk.Revit.DB.Architecture
Imports Autodesk.Revit.DB.Structure
Imports Autodesk.Revit.DB.Mechanical
Imports Autodesk.Revit.DB.Electrical
Imports Autodesk.Revit.DB.Plumbing
Imports Autodesk.Revit.UI
Imports Autodesk.Revit.UI.Selection
Imports Autodesk.Revit.UI.Events
'Imports Autodesk.Revit.Collections
Imports Autodesk.Revit.Exceptions
Imports RvtApplication = Autodesk.Revit.ApplicationServices.Application
Imports RvtDocument = Autodesk.Revit.DB.Document
'
Imports RevitPreview.StructuredS
'
Partial Public Class clsCR
    'Public Shared Sub ApplicationUIEvent_DialogBoxShowing_Handler(sender As Object, e As Autodesk.Revit.UI.Events.DialogBoxShowingEventArgs)
    '    ' Get the string id of the showing dialog
    '    Dim dialogId As String = e.DialogId
    '    If clsCR.crMuestraDialogos = False Then
    '        e.OverrideResult(DialogResult.OK)
    '    End If
    'End Sub
    '' ***** OBLIGATORIOS
    'Public Shared Sub AppUI_ViewActivated_Handler(ByVal sender As [Object], ByVal e As Autodesk.Revit.UI.Events.ViewActivatedEventArgs)
    '    'Dim specArgs As Autodesk.Revit.UI.Events.ViewActivatedEventArgs = TryCast(args, Autodesk.Revit.UI.Events.ViewActivatedEventArgs)
    '    If Not _registered Then
    '        ' Only subscribing once!
    '        crApp = e.Document.Application
    '        SubscribeForApplication(crApp)
    '        _registered = True
    '    End If
    'End Sub

    'Public Shared Sub ApplicationUIEvent_Idling_AbreTrasSaveAs(sender As Object, e As Autodesk.Revit.UI.Events.IdlingEventArgs)
    '    ' Abrir fichero después de Guardar Como.
    '    If (clsCR.crfiAbrir <> "" AndAlso IO.File.Exists(clsCR.crfiAbrir)) = True And clsCR.crfiCerrar = "" Then
    '        clsCR.Fichero_DesencriptaAbreTemp(clsCR.crfiAbrir)
    '        'clsCR.oTimer_Abre.Enabled = True
    '        RemoveHandler crUIAppCont.Idling, AddressOf ApplicationUIEvent_Idling_AbreTrasSaveAs
    '    End If
    'End Sub

    'Public Shared Sub AppUI_ApplicationClosing_Handler(ByVal sender As [Object], ByVal e As Autodesk.Revit.UI.Events.ApplicationClosingEventArgs)
    '    'Dim specArgs As Autodesk.Revit.UI.Events.ApplicationClosingEventArgs = TryCast(args, Autodesk.Revit.UI.Events.ApplicationClosingEventArgs)
    '    If crApp IsNot Nothing AndAlso _registered Then
    '        UnsubscribeForApplication(crApp)
    '        Dim oTask As New System.Threading.Tasks.Task(AddressOf BorraTempAlSalir)
    '        oTask.Start()
    '        _registered = False
    '    End If
    'End Sub
    '' ***** OPCIONALES
    'Public Shared Sub ApplicationUIEvent_DialogBoxShowing_Handler(sender As Object, e As Autodesk.Revit.UI.Events.DialogBoxShowingEventArgs)
    '    ' Get the string id of the showing dialog
    '    Dim dialogId As String = e.DialogId
    '    e.OverrideResult(DialogResult.OK)
    'End Sub

    'Public Shared Sub AppUI_DisplayOptionsDialog_Handler(sender As Object, e As Autodesk.Revit.UI.Events.DisplayingOptionsDialogEventArgs)
    '    'System.Windows.MessageBox.Show("DisplayOptionsDialog")
    'End Sub

    'Public Shared Sub AppUI_DockableFrameFocusChanged_Handler(sender As Object, e As Autodesk.Revit.UI.Events.DockableFrameFocusChangedEventArgs)
    '    'System.Windows.MessageBox.Show("DockableFrameFocusChanged")
    'End Sub
    'Public Shared Sub AppUI_DockableFrameVisibilityChanged_Handler(sender As Object, e As Autodesk.Revit.UI.Events.DockableFrameVisibilityChangedEventArgs)
    '    'System.Windows.MessageBox.Show("DockableFrameVisibilityChanged")
    'End Sub
    'Public Shared Sub AppUI_FabricationPartBrowserChanged_Handler(sender As Object, e As Autodesk.Revit.UI.Events.FabricationPartBrowserChangedEventArgs)
    '    'System.Windows.MessageBox.Show("FabricationPartBrowserChanged")
    'End Sub
    'Public Shared Sub AppUI_Idling_Handler(sender As Object, e As Autodesk.Revit.UI.Events.IdlingEventArgs)
    '    'System.Windows.MessageBox.Show("Idling")
    '    '' Rellenar ITEM_MARKET en las familias ULMA que se hayan insertado.
    '    'If clsCR.uiApp Is Nothing Then
    '    clsCR.crUIApp = TryCast(sender, UIApplication)
    '    clsCR.crApp = crUIApp.Application
    '    'If clsCR.oApp IsNot Nothing Then
    '    clsCR.SubscribeForApplication(crApp)
    '    'End If
    '    'End If
    '    '
    '    'AddHandler appUI.ControlledApplication.DocumentSavingAs, AddressOf appEv_DocumentSavingAs
    '    If clsCR.crUIApp Is Nothing Then
    '        clsCR.ObjetosClase_Pon(clsCR.crUIApp)
    '    End If
    '    '
    '    If clsCR.crUIApp.ActiveUIDocument Is Nothing Then Exit Sub
    '    If clsCR.crUIApp.ActiveUIDocument.Document.IsFamilyDocument Then Exit Sub
    '    ''
    '    clsCR.crDoc = clsCR.crUIApp.ActiveUIDocument.Document
    '    'If oDoc.IsModifiable = False Then Exit Sub
    '    If clsCR.crDoc.IsReadOnly Then Exit Sub
    '    ''
    '    clsCR.crApp.PurgeReleasedAPIObjects()
    'End Sub

    'Public Shared Sub ApplicationUIEvent_TransferredProjectStandards_Handler(sender As Object, e As Autodesk.Revit.UI.Events.TransferredProjectStandardsEventArgs)
    '    'System.Windows.MessageBox.Show("TransferredProjectStandards")
    'End Sub
    'Public Shared Sub ApplicationUIEvent_TransferringProjectStandards_Handler(sender As Object, e As Autodesk.Revit.UI.Events.TransferringProjectStandardsEventArgs)
    '    'System.Windows.MessageBox.Show("TransferringProjectStandards")
    'End Sub
    'Public Shared Sub ApplicationUIEvent_ViewActivating_Handler(sender As Object, e As Autodesk.Revit.UI.Events.ViewActivatingEventArgs)
    '    'System.Windows.MessageBox.Show("ViewActivating")
    'End Sub
End Class
