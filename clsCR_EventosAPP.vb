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
    '
    ' ***** OPCIONALES
    'Public Shared Sub ApplicationEvent_ApplicationInitialized(sender As Object, e As Autodesk.Revit.DB.Events.ApplicationInitializedEventArgs)
    '    '' ***** Aquí pondremos lo que queremos que carge, después de tener Revit iniciacizado y
    '    'System.Windows.MessageBox.Show("ApplicationInitialized")
    '    ' SubscribeForApplication(oApp)
    'End Sub
    '
    'Public Shared Sub App_DocumentChanged_Handler(ByVal sender As Object, e As Autodesk.Revit.DB.Events.DocumentChangedEventArgs)
    '    '
    'End Sub

    'Public Shared Sub App_DocumentCreating_Handler(ByVal sender As Object, ByVal e As Autodesk.Revit.DB.Events.DocumentCreatingEventArgs)
    '    'System.Windows.MessageBox.Show("DocumentCreating")
    '    crUltimaplantilla = e.Template
    '    '
    '    If IO.File.Exists(crUltimaplantilla) Then
    '        If (IO.Path.GetDirectoryName(crUltimaplantilla).ToUpper.Contains("UCREVIT") And
    '        IO.Path.GetDirectoryName(crUltimaplantilla).ToUpper.EndsWith("\TEMPLATES")) Then
    '            'Dim fiTemp As String = clsCr.Fichero_PonNuevo(ultimaplantilla)
    '            'clsCr.FicheroDesencripta(fiTemp)
    '            clsCR.Fichero_Desencripta(e.Template)
    '        Else
    '            crUltimaplantilla = ""
    '        End If
    '    End If
    'End Sub
    'Public Shared Sub App_DocumentCreated_Handler(ByVal sender As Object, ByVal e As Autodesk.Revit.DB.Events.DocumentCreatedEventArgs)
    '    If IO.File.Exists(crUltimaplantilla) Then
    '        If (IO.Path.GetDirectoryName(crUltimaplantilla).ToUpper.Contains("UCREVIT") And
    '        IO.Path.GetDirectoryName(crUltimaplantilla).ToUpper.EndsWith("\TEMPLATES")) Then
    '            clsCR.Fichero_Encripta(crUltimaplantilla, False)
    '        Else
    '            crUltimaplantilla = ""
    '        End If
    '    End If
    'End Sub

    'Public Shared Sub App_DocumentOpening_Handler(ByVal sender As Object, ByVal e As Autodesk.Revit.DB.Events.DocumentOpeningEventArgs)
    '    clsCR.crPuedoborrar = False
    '    '
    '    If crip2aCAD.clsCR.crDicFicheros.ContainsKey(e.PathName) = False Then
    '        If crip2aCAD.clsCR.crEncrypt = False Then
    '            crip2aCAD.clsCR.Fichero_Desencripta(e.PathName)
    '        ElseIf crip2aCAD.clsCR.crEncrypt = True Then
    '            crip2aCAD.clsCR.Fichero_PonNuevo(e.PathName, True)
    '            crip2aCAD.clsCR.Fichero_DesencriptaAbreTemp(e.PathName)
    '            If e.Cancellable Then e.Cancel()
    '            'AddHandler appContUI.Idling, AddressOf ApplicationUIEvent_Idling_AbreTrasSaveAs
    '        End If
    '    End If
    '    RemoveHandler crApp.DocumentOpening, AddressOf App_DocumentOpening_Handler
    '    clsCR.crPuedoborrar = True
    'End Sub

    'Public Shared Sub App_DocumentOpened_Handler(ByVal sender As [Object], ByVal e As Autodesk.Revit.DB.Events.DocumentOpenedEventArgs)
    '    If e.IsCancelled = True Then
    '        AddHandler crApp.DocumentOpening, AddressOf App_DocumentOpening_Handler
    '        Exit Sub
    '    End If
    '    'clsCR.puedoborrar = False
    '    'Dim specArgs As Autodesk.Revit.DB.Events.DocumentOpenedEventArgs = TryCast(args, Autodesk.Revit.DB.Events.DocumentOpenedEventArgs)
    '    'SubscribeToDoc(e.Document)
    '    ''
    '    'If e.Document Is Nothing Then Exit Sub
    '    'If crDoc IsNot Nothing AndAlso crDoc.PathName = e.Document.PathName Then Exit Sub
    '    ''
    '    'crDoc = e.Document
    '    'If e.Document.IsFamilyDocument Or e.Document.PathName = "" Then Exit Sub
    '    'If IO.Path.GetExtension(e.Document.PathName).ToUpper <> ".RVT" Then Exit Sub
    '    ' Si es un RVT. Pero ya estaba procesado, también salimos.
    '    'clsCR.puedoborrar = True
    'End Sub
    ''
    'Public Shared Sub App_DocumentSavingAs_Handler(ByVal sender As Object, ByVal e As Autodesk.Revit.DB.Events.DocumentSavingAsEventArgs)
    '    clsCR.crPuedoborrar = False
    '    'System.Windows.MessageBox.Show("DocumentSavingAs")
    '    crUltimodoccerrado = e.PathName   ' Nombre del nuevo documento
    '    'e.Document.PathName 'Nombre del documento que se está guardardo.
    '    '
    '    If e.Document.IsFamilyDocument = False Then
    '        '' Si es un documento de proyecto y no lleva alguna familia ULMA no lo codificaremos.
    '        Dim colUlmaFam As IList(Of ElementId) = clsCR.FamilySymbol_DameULMA_ID(e.Document, BuiltInCategory.OST_GenericModel)
    '        '' No lleva familias ulma. No lo codificaremos.
    '        If colUlmaFam Is Nothing OrElse (colUlmaFam IsNot Nothing AndAlso colUlmaFam.Count = 0) Then
    '            crUltimodoccerrado = ""
    '        Else
    '            crUltimodoccerrado = e.PathName
    '            'If e.Cancellable Then e.Cancel()
    '        End If
    '    ElseIf e.Document.IsFamilyDocument = True Then
    '        Dim manu As String = e.Document.FamilyManager.Parameter(BuiltInParameter.ALL_MODEL_MANUFACTURER).Formula
    '        ' Si Manufacturer no contiene ULMA
    '        ' O si el tipo de familia no es Modelo Genérico (typeFamily)
    '        If manu Is Nothing Then
    '            crUltimodoccerrado = ""
    '        ElseIf manu.ToUpper.Contains("ULMA") And e.Document.OwnerFamily.FamilyCategoryId.IntegerValue.Equals(CInt(typeFamily)) = True Then
    '            crUltimodoccerrado = e.PathName
    '        End If
    '    End If
    '    'If ultimodoccerrado <> "" And encrypt = True Then
    '    '    clsCR.oTimer_Cierra.Enabled = True
    '    'End If
    '    clsCR.crPuedoborrar = True
    'End Sub
    '
    'Public Shared Sub App_DocumentSaved_Handler(ByVal sender As Object, ByVal e As Autodesk.Revit.DB.Events.DocumentSavedEventArgs)
    '    clsCR.crPuedoborrar = False
    '    'System.Windows.MessageBox.Show("DocumentSaved")
    '    crUltimoguardado = e.Document.PathName
    '    ' Si no estaba en Temp, salir.
    '    If crUltimoguardado.StartsWith(crfolder_TempRaiz) = False Then Exit Sub
    '    '
    '    Dim ultimoguardadoOriginal As String = crDicFicheros(crUltimoguardado)
    '    ''
    '    'If e.Document.IsFamilyDocument = False Then
    '    Dim colUlmaFam As IList(Of ElementId) = clsCR.FamilySymbol_DameULMA_ID(e.Document, BuiltInCategory.OST_GenericModel)
    '    '' No lleva familias ulma. No lo codificaremos.
    '    If colUlmaFam IsNot Nothing AndAlso colUlmaFam.Count = 0 Then
    '        crUltimoguardado = ""
    '    End If
    '    'ElseIf e.Document.IsFamilyDocument = True Then
    '    'ultimoguardado = ""
    '    'End If
    '    '
    '    If e.Document.IsFamilyDocument = True Then

    '    End If
    '    If crUltimoguardado <> "" And crUltimoguardado <> ultimoguardadoOriginal Then
    '        If crip2aCAD.clsCR.crEncrypt = True Then
    '            Try
    '                'IO.File.Copy(ultimoguardado, ultimoguardadoOriginal, True)
    '                crip2aCAD.clsCR.Fichero_Encripta(ultimoguardadoOriginal, False)
    '            Catch ex As Exception
    '                Console.WriteLine(ex)
    '            End Try
    '        End If
    '    End If
    '    clsCR.crPuedoborrar = True
    'End Sub
    '
    'Public Shared Sub App_DocumentSavedAs_Handler(ByVal sender As Object, ByVal e As Autodesk.Revit.DB.Events.DocumentSavedAsEventArgs)
    '    clsCR.crPuedoborrar = False
    '    'System.Windows.MessageBox.Show("DocumentSavedAs")
    '    ''
    '    Dim fiActual As String = e.Document.PathName
    '    If e.Document.IsModified And e.Document.PathName <> "" Then
    '        Try
    '            e.Document.Save()
    '        Catch ex As Exception
    '            Debug.Print(ex.ToString)
    '        End Try
    '    End If
    '    '
    '    If crUltimodoccerrado <> "" AndAlso IO.File.Exists(crUltimodoccerrado) Then
    '        ' Es un fichero de ulma
    '        If fiActual <> "" AndAlso IO.File.Exists(fiActual) Then
    '            ' Encryptar el que cerramos. copiarlo al originales y borrarlo de Temp
    '            clsCR.Fichero_Quita(e.OriginalPath)
    '            'Dim fiTemp As String = clsCR.Fichero_PonNuevo(fiActual)
    '            clsCR.crfiAbrir = clsCR.Fichero_PonNuevo(fiActual)
    '            clsCR.crfiCerrar = fiActual
    '            'clsCr.FormularioOpen(fiTemp)
    '            ' Cerrar documento, si estaba abierto.
    '            'Call utilesRevit.DocumentoEstaAbierto(e.Document.Application, e.Document.PathName, True)
    '            ' oDoc.Close(False)
    '            ' clsCr.FicheroEncripta(e.Document.PathName, False)
    '        End If
    '        '
    '        If crip2aCAD.clsCR.crEncrypt = True Then
    '            crip2aCAD.clsCR.oTimer_Cierra.Enabled = True
    '        End If
    '        '
    '        AddHandler crUIAppCont.Idling, AddressOf ApplicationUIEvent_Idling_AbreTrasSaveAs
    '    End If
    '    clsCR.crPuedoborrar = True
    'End Sub
End Class
