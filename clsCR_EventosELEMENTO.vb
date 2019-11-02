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
Imports adWin = Autodesk.Windows
Partial Public Class clsCR
    'in OnStartup 
    'AddHandler() Autodesk.Windows.ComponentManager.UIElementActivated, AddressOf MyUiElementActivated
    'in OnShutdown
    'RemoveHandler() Autodesk.Windows.ComponentManager.UIElementActivated, AddressOf MyUiElementActivated
    'Public Shared Sub MyUiElementActivated(ByVal sender As Object, ByVal e As Autodesk.Windows.UIElementActivatedEventArgs)
    '    'do your thing
    '    Dim mensaje As String = ""
    '    If TypeOf e.Item Is adWin.RibbonItem Then
    '        On Error Resume Next
    '        mensaje &= "AutomationName = " & e.Item.AutomationName & vbCrLf
    '        mensaje &= "Name = " & e.Item.Name & vbCrLf
    '        mensaje &= "Text = " & e.Item.Text & vbCrLf
    '        mensaje &= "Description = " & e.Item.Description & vbCrLf
    '        mensaje &= "Id = " & e.Item.Id & vbCrLf
    '        mensaje &= "Cookie=" & e.Item.Cookie & vbCrLf
    '        mensaje &= "UId = " & e.Item.UID & vbCrLf
    '        mensaje &= "Description = " & e.Item.Description & vbCrLf
    '        mensaje &= "Image = " & e.Item.Image.ToString & vbCrLf
    '        'e.UiElement.CommandBindings.Add()
    '    End If
    '    If mensaje <> "" Then
    '        'TaskDialog.Show("Datos del botón pulsado (UIElement)", mensaje)
    '        IO.File.AppendAllText(IO.Path.Combine("H:\DESARROLLO\REVIT_DOC\COMMANDS-EXECUTE-IEXTERNAL", "_comandos.txt"), mensaje & vbCrLf)
    '    End If
    'End Sub
End Class
