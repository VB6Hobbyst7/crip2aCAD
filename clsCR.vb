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
'Imports Autodesk.Revit
'Imports Autodesk.Revit.ApplicationServices
'Imports Autodesk.Revit.Attributes
'Imports Autodesk.Revit.DB
'Imports Autodesk.Revit.DB.Events
'Imports Autodesk.Revit.DB.Architecture
'Imports Autodesk.Revit.DB.Structure
'Imports Autodesk.Revit.DB.Mechanical
'Imports Autodesk.Revit.DB.Electrical
'Imports Autodesk.Revit.DB.Plumbing
'Imports Autodesk.Revit.UI
'Imports Autodesk.Revit.UI.Selection
'Imports Autodesk.Revit.UI.Events
'Imports Autodesk.Revit.Collections
'Imports Autodesk.Revit.Exceptions
'Imports Autodesk.Revit.ApplicationServices.Application
'Imports Autodesk.Revit.DB.Document
'
'Imports RevitPreview.StructuredS
Partial Public Class clsCR
    '
    'Public Shared crAppContUI As Autodesk.Revit.UI.UIControlledApplication
    Public Shared crAppUI As Autodesk.Revit.UI.UIApplication
    '
    ' COLECCIONES
    Public Shared crDicFicheros As Dictionary(Of String, String)     'Key=fullPath sin codificar en TEMP, Value=fulPath fichero Revit codificado
    Public Shared crDicFicherosNO As List(Of String)                 ' Ficheros que no codificaremos.
    Public Shared crFicherosEnUso As List(Of String)                ' Fichero que no borraremos de Temp, mientras estén en uso.
    ' (fullPath es padres, fullPath|fullPath en hijos)
    Public Shared crColExtBIN As List(Of String) = Nothing
    Public Shared crColExtTXT As List(Of String) = Nothing
    Public Shared crColExtTXTc As List(Of String) = Nothing
    Public Shared crColExtEnc As List(Of String) = Nothing
    Public Shared crColExtDes As List(Of String) = Nothing
    '
    ' VARIABLES
    Public Const crFabricante As String = "ULMA"
    Public Shared crfiAbrir As String = ""
    Public Shared crfiCerrar As String = ""
    Public Shared crfiLinkOld As String = ""
    Public Shared crfiLinkNew As String = ""
    Public Shared crEncrypt As Boolean = True               ' 0 (False) = No encripta ficheros Revit, 1 (True) = Si encrypta ficheros Revit (Siempre desencripta, si esta encriptado)
    Public Shared crUltimodoccerrado As String = ""
    Public Shared crUltimodoccerradoEnTemp As String = ""
    Public Shared crUltimaplantilla As String = ""
    Public Shared crUltimoguardado As String = ""
    Public Shared crSaveImagePreview As Boolean = False         '' Guardar .jpg de las familias que se guarden.
    Public Shared cr_user As String = "123098!·$%&/"         '' Usuario autorizado. Ponemos un valor erroneo temporal
    Public Shared cr_password As String = "123098!·$%&/"     '' Clave autorizada del usuaario. Ponemos en valore erroneo temporal
    ' CAMINOS
    Public Shared crfolder_library As String = ""
    Public Shared crfolder_templates As String = ""
    Public Shared crfolder_ultimo As String = ""
    Public Shared crfolder_TempRaiz As String = My.Computer.FileSystem.SpecialDirectories.Temp ' IO.Path.GetTempPath
    Public Const crfolder_TempRaizRevit As String = "tmplog"
    Public Shared cr_ficherolog As String = ""                    '' FulPath del fichero .log (Para auditar cosas)
    '
    Public Shared crfolder_Temp As String = IO.Path.Combine(crfolder_TempRaiz, crfolder_TempRaizRevit)
    '
    Public Shared crfolder_MisDoc As String = My.Computer.FileSystem.SpecialDirectories.MyDocuments
    Public Shared crfolder_Personal As String = Environment.GetFolderPath(Environment.SpecialFolder.Personal)
    ' CONTROL EVENTOS
    Public Shared crPuedoborrar As Boolean = True
    Public Shared crEventoAbrir As Boolean = True
    Public Shared crEventoCerrar As Boolean = True
    ' Clave inicial para activar funciones DLL
    Private ReadOnly key As String = "aiiao2K19"
    Protected Friend Shared activado As Boolean = False

    Public Sub New(k As String)
        If String.Equals(k, Me.key) Then ', StringComparison.CurrentCulture) Then
            activado = True
        Else
            activado = False
        End If
    End Sub
    'Public Shared crMuestraDialogos As Boolean = False

    Public Shared Sub Inicio_clsCR()
        ' Si no está activado, salir ****
        If activado = False Then Exit Sub
        ' ********************************
        crDicFicheros = New Dictionary(Of String, String)
        crDicFicherosNO = New List(Of String)
        crFicherosEnUso = New List(Of String)
        'crAppContUI = app
        ' Timer para abrir documentos, después de eventos
        oTimer_Abre = New System.Timers.Timer
        oTimer_Abre.Interval = 1000 * 4         ' Cada 2 sengundos.
        ' Timer para borrar fichero de folderTemp
        oTimer_Borra = New System.Timers.Timer
        'oTimer_Borra.Interval = 1000 * 60 * 2   ' Cada 2 minutos
        oTimer_Borra.Interval = 1000 * 30   ' Cada 30 segundos
        'oTimer_Borra.Enabled = True
        Configuracion_Cargar()
        'Dim dirTemp As String = IO.Path.Combine(folder_TempRevit, queGuid)
        'Dim fiTemp As String = IO.Path.Combine(folder_TempRevit, queGuid, IO.Path.GetFileName(fiRevit))
        If IO.Directory.Exists(crfolder_Temp) = False Then
            IO.Directory.CreateDirectory(crfolder_Temp)
        End If
        'Poner oculto el directorio "crfolder_Temp"
        Directorio_Hidden(crfolder_Temp, ocultar:=True)
        '
        crPuedoborrar = True
        FicherosDesarrollo_BorraTempTASK()
    End Sub
    '
    Public Shared Sub dicFicherosNo_Pon(queFi As String)
        ' Si no está activado, salir ****
        If activado = False Then Exit Sub
        ' ********************************
        If crDicFicherosNO Is Nothing Then crDicFicherosNO = New List(Of String)
        If crDicFicherosNO.Contains(queFi) = False Then crDicFicherosNO.Add(queFi)
    End Sub

    Public Shared Sub Configuracion_Cargar(Optional encriptar As Boolean = True)
        ' Si no está activado, salir ****
        If activado = False Then Exit Sub
        ' ********************************
        ' Colecciones de extensiones (Binarios y Texto)
        crColExtBIN = New List(Of String)
        crColExtTXT = New List(Of String)
        crColExtTXTc = New List(Of String)
        ' Colecion de extensiones para Encriptar/Desencriptar (En las de texto, poner c al final)
        crColExtEnc = New List(Of String)
        crColExtDes = New List(Of String)
        For Each queExt As String In My.Settings.extEncBIN.Split(","c)
            If queExt.StartsWith(".") Then
                crColExtBIN.Add(queExt)
                crColExtEnc.Add(queExt)
                crColExtDes.Add(queExt)
            Else
                crColExtBIN.Add("." & queExt)
                crColExtEnc.Add("." & queExt)
                crColExtDes.Add("." & queExt)
            End If
        Next
        ''
        For Each queExt As String In My.Settings.extEncTXT.Split(","c)
            If queExt.StartsWith(".") Then
                crColExtTXT.Add(queExt)
                crColExtTXTc.Add(queExt + "c")
                crColExtEnc.Add(queExt)
                crColExtDes.Add(queExt + "c")
            Else
                crColExtTXT.Add("." & queExt)
                crColExtTXTc.Add("." & queExt + "c")
                crColExtEnc.Add("." & queExt)
                crColExtDes.Add("." & queExt + "c")
            End If
        Next
        '
        crColExtBIN.Sort() : crColExtTXT.Sort() : crColExtEnc.Sort() : crColExtDes.Sort()
        '
        Dim _data As String = My.Settings.data
        Try
            _data = Texto_Desencriptar(_data)
            cr_user = _data.Split("|"c)(0)
            cr_password = _data.Split("|"c)(1)
        Catch ex As Exception
            cr_user = "123098!·$%&/"
            cr_password = "123098!·$%&/"
        End Try
    End Sub
    '
    Public Shared Function Fichero_PonNuevo(fiRevit As String, Optional copiar As Boolean = True) As String
        ' Si no está activado, salir ****
        If activado = False Then Return "" : Exit Function
        ' ********************************
        If crDicFicheros Is Nothing Then crDicFicheros = New Dictionary(Of String, String)
        'If crDicFicherosNO.Contains(fiRevit) Then 'IO.File.Exists(fiRevit) = False Or 
        '    Return ""
        '    Exit Function
        'End If
        '
        Dim fiTemp As String = ""
        'If crDicFicheros.ContainsValue(fiRevit) Then
        '    fiTemp = Fichero_DameDesdeValue(fiRevit)
        'Else
        fiTemp = IO.Path.Combine(crfolder_Temp, IO.Path.GetFileName(fiRevit))
            If IO.File.Exists(fiTemp) Then
                Dim queGuid As String = Guid.NewGuid.ToString
                Dim dirTemp As String = IO.Path.Combine(crfolder_Temp, queGuid)
                If IO.Directory.Exists(dirTemp) = False Then
                    IO.Directory.CreateDirectory(dirTemp)
                End If
                fiTemp = IO.Path.Combine(dirTemp, IO.Path.GetFileName(fiRevit))
            End If
        '
        Fichero_Quita(fiTemp, False)
        Fichero_Quita(fiRevit, False)
        '
        ' Añadir a la colección. No tiene que existir. Lo hemos borrado si existía
        Try
            If IO.File.Exists(fiRevit) Then
                Fichero_ReadOnly(fiRevit, sololectura:=False)
                If copiar Then
                    Try
                        IO.File.Copy(fiRevit, fiTemp, True)
                    Catch ex As Exception
                        Debug.Print(ex.ToString)
                    End Try
                End If
                'crDicFicheros.Add(fiTemp, fiRevit)
                crfiAbrir = fiTemp
                'crfiCerrar = fiRevit
            End If
            ' Siempre añadirlo a la colección.
            If crDicFicheros.ContainsKey(fiTemp) = False Then crDicFicheros.Add(fiTemp, fiRevit)
        Catch ex As Exception
                Debug.Print(ex.ToString)
            End Try
        'End If
        '
        Return fiTemp
    End Function
    '
    Public Shared Function Fichero_DameDesdeValue(fullFiRevit As String) As String
        ' Si no está activado, salir ****
        If activado = False Then Return "" : Exit Function
        ' ********************************
        Dim fullFiTemp As String = ""
        ' Si existe en Values (Puede haber varios)
        If crDicFicheros.ContainsValue(fullFiRevit) Then
            Dim itemsKeys = (From pair In crDicFicheros
                             Where pair.Value.Equals(fullFiRevit)
                             Select pair.Key).ToArray()

            fullFiTemp = itemsKeys.First
        Else
            'fullFiTemp = "" ' Fichero_DameIgualEnTemp(fullFiRevit)
            'If dicFicheros.ContainsKey(fullFiTemp) = False Then
            '    dicFicheros.Add(fullFiTemp, fullFiRevit)
            'End If
            fullFiTemp = Fichero_PonNuevo(fullFiRevit)
        End If
        Return fullFiTemp
    End Function
    Public Shared Function Fichero_DameDesdeKey(tempFiRevit As String, Optional folder As Boolean = False) As String
        ' Si no está activado, salir ****
        If activado = False Then Return "" : Exit Function
        ' ********************************
        Dim resultado As String = ""
        Dim fialternativo = Fichero_DameDesdeKeySoloNombre(tempFiRevit)
        '
        If crDicFicheros.ContainsKey(tempFiRevit) Then
            resultado = crDicFicheros(tempFiRevit)
        ElseIf tempFiRevit.StartsWith(crfolder_TempRaiz) = False Then
            resultado = tempFiRevit
        ElseIf fialternativo <> "" And IO.File.Exists(fialternativo) Then
            resultado = fialternativo
        ElseIf IO.File.Exists(tempFiRevit) Then
            resultado = IO.Path.Combine(crfolder_MisDoc, IO.Path.GetFileName(tempFiRevit))
        End If
        '
        If folder = True Then
            resultado = IIf(IO.File.Exists(resultado) = True, IO.Path.GetDirectoryName(resultado), crfolder_MisDoc)
        End If
        '
        Return resultado
    End Function
    '
    Public Shared Function Fichero_DameDesdeKeySoloNombre(tempFiRevit As String, Optional folder As Boolean = False) As String
        ' Si no está activado, salir ****
        If activado = False Then Return "" : Exit Function
        ' ********************************
        Dim resultado As String = ""
        ' Si existe en Values (Puede haber varios)
        For Each queFi As String In crDicFicheros.Keys
            Dim nTemp As String = IO.Path.GetFileName(tempFiRevit)
            Dim nDic As String = IO.Path.GetFileName(queFi)
            If nTemp = nDic Then
                resultado = crDicFicheros(queFi)
                Exit For
            End If
        Next
        '
        Return resultado
    End Function
    Public Shared Sub Fichero_Quita(fullFi As String, Optional borrar As Boolean = True)
        ' Si no está activado, salir ****
        If activado = False Then Exit Sub
        ' ********************************
        ' Quitar de crFicherosEnUso
        'For Each queFi As String In crFicherosEnUso
        '    ' PathName (Padres) ó PathName|PatnName (Padre|link hijo)
        '    If queFi = fullFi Or queFi.StartsWith(fullFi) Then
        '        Call crFicherosEnUso.Remove(queFi)
        '    End If
        'Next
        If crFicherosEnUso.Contains(fullFi) Then
            Call crFicherosEnUso.Remove(fullFi)
        End If
        ' Si existe en Keys (Es unico, solo tiene que haber 1)
        If crDicFicheros.ContainsKey(fullFi) Then
            Call crDicFicheros.Remove(fullFi)
            Try
                If borrar And fullFi.StartsWith(crfolder_Temp) Then IO.File.Delete(fullFi)
            Catch ex As Exception
                Debug.Print(ex.ToString)
            End Try
        End If
        ' Si existe en Values (Puede haber varios)
        If crDicFicheros.ContainsValue(fullFi) Then
            Dim itemsToRemove As New List(Of String)
            For Each queItem As String In crDicFicheros.Keys
                If crDicFicheros(queItem) = fullFi Then
                    itemsToRemove.Add(queItem)
                End If
            Next
            'Dim itemsToRemove = (From pair In crDicFicheros
            '                     Where pair.Value.Equals(fullFi)
            '                     Select pair.Key).ToArray()

            For Each item As String In itemsToRemove
                crDicFicheros.Remove(item)
                Try
                    If borrar And item.StartsWith(crfolder_Temp) Then IO.File.Delete(item)
                Catch ex As Exception
                    Debug.Print(ex.ToString)
                End Try
            Next
        End If
    End Sub
    '
    'Public Shared Sub Cierra_DocumentosUlma(oApp As Autodesk.Revit.ApplicationServices.Application)
    '    Dim cerraractual As Boolean = False
    '    For Each queDoc As Autodesk.Revit.DB.Document In oApp.Documents
    '        Try
    '            'Dim colIds As List(Of DB.ElementId) = clsCR.FamilyInstance_DameULMA_ID(queDoc)
    '            Dim colIds As List(Of DB.ElementId) = clsCR.FamilySymbol_DameULMA_ID(queDoc, BuiltInCategory.OST_GenericModel)
    '            If colIds IsNot Nothing AndAlso colIds.Count > 0 Then
    '                If queDoc.PathName <> crAppUI.ActiveUIDocument.Document.PathName Then
    '                    queDoc.Close(True)
    '                Else
    '                    cerraractual = True
    '                End If
    '            End If
    '            colIds = Nothing
    '        Catch ex As Exception
    '            Continue For
    '        End Try
    '    Next
    '    '
    '    If cerraractual Then
    '        Dim rvtIntPtr As IntPtr = Autodesk.Windows.ComponentManager.ApplicationWindow
    '        clsAPI.SetForegroundWindow(rvtIntPtr)
    '        Try
    '            clsAPI.keybd_event(CByte(clsAPI.eMensajes.VK_ESCAPE), 0, 0, UIntPtr.Zero)
    '            clsAPI.keybd_event(CByte(clsAPI.eMensajes.VK_ESCAPE), 0, 2, UIntPtr.Zero)
    '            clsAPI.keybd_event(CByte(clsAPI.eMensajes.VK_ESCAPE), 0, 0, UIntPtr.Zero)
    '            clsAPI.keybd_event(CByte(clsAPI.eMensajes.VK_ESCAPE), 0, 2, UIntPtr.Zero)
    '            System.Windows.Forms.SendKeys.SendWait("^{F4}")
    '        Catch ex As Exception
    '            Debug.Print(ex.ToString)
    '        End Try
    '        clsCR.crfiCerrar = ""
    '    End If
    'End Sub
End Class
