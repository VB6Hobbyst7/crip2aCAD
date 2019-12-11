Imports System
Imports System.Drawing
Imports System.IO
Imports System.Security
Imports System.Security.Cryptography
Imports System.Text
''
Imports RevitPreview.StructuredS
'
Partial Public Class clsCR
    'Public Const codbinario As String = "2ACGG.,b"      '' Siempre 8 carácteres (Codificado inicial para Binario)
    'Public Shared codbinarioEnc As String = "11zMNZcc0puPeOFO/bgC2w==" '   "2ACGG.,b" Codificado
    '
    ' 2018/10/31 Utilizaremos codificador para encriptar.
    ' Pero también en la cabecera del fichero encriptado, para saber qué encriptación/desencriptación utilizar.
    '** El codificador siempre serán 24 carácteres (Antes era de 2ACAD_GG y antes 2ACGG.,.) para compatibilidad con anterior 11zMNZcc0puPeOFO/bgC2w==
    Public Shared codificador As String = "2ACAD_LU"
    Public Shared codificadorBinarioEnc() As Byte = Encoding.ASCII.GetBytes(codificador)
    Public Shared IV() As Byte = ASCIIEncoding.ASCII.GetBytes(codificador)    ''La clave debe ser de 8 caracteres
    Public Shared EncryptionKey() As Byte = Convert.FromBase64String("rpaSPvIvVLlrcmtzPU9/c67Gkj7yL1S5") 'No se puede alterar la cantidad de caracteres pero si la clave
    Public Shared CabeceraFicherosRevit As String = "쿐놡"      ' Todos los ficheros Revit sin codificar empiezan por esto (RVT y RFA)
    '
    'Todos los codificadores que hemos utilizado (El primero siempre será el actual)
    Public Shared arrCodificadores As String() = New String() {codificador, "2ACAD_GG", "2ACGG.,."}
    'Public Shared arrIV As Array = New Array() {ASCIIEncoding.ASCII.GetBytes(arrCodificador(0)), ASCIIEncoding.ASCII.GetBytes(arrCodificador(1)), ASCIIEncoding.ASCII.GetBytes(arrCodificador(2))}
    '
    '*************************
    '** Global Variables
    '*************************
    Public Shared fsInput As System.IO.FileStream
    Public Shared fsOutput As System.IO.FileStream
    '
    '' CONSTANTES
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    'Public Const preUnoNOEncriptado As String = "<?xml version="        ' StartWith
    'Public Const preUnoSIEncriptado As String = "CjEZhvvo0GSRkzsBWqLLMZMu9lQoleGmpbmXf02mlLBJmSS3/L0cGKCzObdXHC5Q3n6ZT0+n6YU="
    '
    'Public Shared Function codbinarioEnc() As String
    '    Return Texto_Encripta(codbinario)
    'End Function
    'Public Shared Function codtextoEnc() As String
    '    Return Texto_Encripta(codtexto)
    'End Function
    Public Shared Sub IV_Pon(codificadorNew As String)
        ' Si no está activado, salir ****
        If activado = False Then Exit Sub
        ' ********************************
        codificador = codificadorNew
        ' Según la cabecera del fichero encriptado, usaremos un codificador u otro para desencriptar el contenido
        ' codificador siempre debe ser de 8 carácteres.
        'Public Const codificador As String = "2ACAD_LU" '' El codificador siempre serán 8 carácteres (Antes era 2ACAD_GG y antes 2ACGG.,.)
        codificadorBinarioEnc = Encoding.ASCII.GetBytes(codificadorNew)
        IV = ASCIIEncoding.ASCII.GetBytes(codificadorNew)    ''La clave debe ser de 8 caracteres
        'Public Shared EncryptionKey() As Byte = Convert.FromBase64String("rpaSPvIvVLlrcmtzPU9/c67Gkj7yL1S5") 'No se puede alterar la cantidad de caracteres pero si la clave
    End Sub
    '
    Public Shared Function Fichero_DameTipoEncriptacion(FileIn As String) As TipoCod
        ' Si no está activado, salir ****
        If activado = False Then Return Nothing : Exit Function
        ' ********************************
        Dim resultado As TipoCod = TipoCod.sincodificarRevit
        If IO.File.Exists(FileIn) = False Then
            Return TipoCod.noexistefichero
            Exit Function
        End If
        ' Abre el archivo, lee todas las lineas y evalua si la primera empieza por CabeceraFicherosRevit. Cierra el archivo.
        'Dim esRevitSinCodificar As Boolean = False
        'Try
        '    esRevitSinCodificar = File.ReadLines(FileIn, Encoding.Unicode).FirstOrDefault.StartsWith(CabeceraFicherosRevit)
        'Catch ex As Exception
        '    ' Error al leer la cabecera.
        'End Try
        'If esRevitSinCodificar Then
        '    resultado = TipoCod.sincodificarRevit
        '    GoTo Final
        '    Exit Function
        'End If
        ''
        '' Comprobar si es un fichero codificado por la aplicación o no.
        '' Nos devolverá el tipo de codificación que tenga (sincodificar, esbinario, estexto o ficheroenuso)
        Try
            Fichero_ReadOnly(FileIn, sololectura:=False)
            fsInput = New System.IO.FileStream(FileIn, FileMode.Open,
                                                      FileAccess.ReadWrite)
        Catch ex As Exception
            '' El fichero está en uso y no se puede escribir en él.
            If fsInput IsNot Nothing Then fsInput.Close()
            fsInput = Nothing
            Return TipoCod.ficheroenuso
            Exit Function
        End Try
        ''
        '' Leer la cabecera para saber que tipo de codificación tiene (sincodificar o codificado (esbinario o estexto))
        ' Codificador de Binario (Siempre debe ser 8 carácteres)
        Dim codificadoBloqueB(codificadorBinarioEnc.Count - 1) As Byte
        Dim intBytesB As Integer = fsInput.Read(codificadoBloqueB, 0, codificadorBinarioEnc.Count)
        Dim valorUnicode As String = Encoding.Unicode.GetString(codificadoBloqueB)
        Dim valorB As String = Encoding.ASCII.GetString(codificadoBloqueB)
        '
        If valorUnicode.StartsWith(CabeceraFicherosRevit) Then
            ' Fichero Revit sin codificar.
            fsInput.Seek(0, SeekOrigin.Begin)
            resultado = TipoCod.sincodificarRevit
        ElseIf arrCodificadores.Contains(valorB) = False Then
            fsInput.Seek(0, SeekOrigin.Begin)
            resultado = TipoCod.sincodificarOtros       ' O tiene la codificación vieja.
        ElseIf arrCodificadores.Contains(valorB) And crColExtBIN.Contains(IO.Path.GetExtension(FileIn)) Then
            ' Es codificado binario
            fsInput.Seek(0, SeekOrigin.Begin)
            resultado = TipoCod.codificadoB
            IV_Pon(valorB)
        ElseIf arrCodificadores.Contains(valorB) And crColExtTXTc.Contains(IO.Path.GetExtension(FileIn)) Then
            ' Es codificado texto
            fsInput.Seek(0, SeekOrigin.Begin)
            resultado = TipoCod.codificadoT
            IV_Pon(valorB)
        End If
        ''
Final:
        If fsInput IsNot Nothing Then fsInput.Close()
        fsInput = Nothing
        ''
        Return resultado
    End Function
    '
    Public Shared Sub DocumentoActual_EncriptaLinks(Optional mDoc As Autodesk.Revit.DB.Document = Nothing)
        ' Si no está activado, salir ****
        If activado = False Then Exit Sub
        ' ********************************
        ' Documento activo y filtro para Link Files
        Dim mainDoc As Autodesk.Revit.DB.Document = Nothing
        If mDoc Is Nothing Then
            mainDoc = crAppUI.ActiveUIDocument.Document
        Else
            mainDoc = mDoc
        End If
        ' Primero procesamos sólo los que no tienen Padre. Los link directos
        Dim collI As Autodesk.Revit.DB.FilteredElementCollector = New Autodesk.Revit.DB.FilteredElementCollector(mainDoc)
        collI.OfClass(GetType(Autodesk.Revit.DB.RevitLinkInstance))
        '
        For Each rIns As Autodesk.Revit.DB.RevitLinkInstance In collI
            Dim rType As Autodesk.Revit.DB.RevitLinkType = CType(mainDoc.GetElement(rIns.GetTypeId), Autodesk.Revit.DB.RevitLinkType)
            ' Solo los que no tienen padre (GetParentID = ElementId.InvalidElementId)
            If rType.GetParentId.Equals(Autodesk.Revit.DB.ElementId.InvalidElementId) Then
                Dim docHijo As Autodesk.Revit.DB.Document = rIns.GetLinkDocument
                If docHijo IsNot Nothing AndAlso docHijo.IsModified = True And crDicFicheros.ContainsKey(docHijo.PathName) Then
                    If IO.File.Exists(docHijo.PathName) = False Then
                        docHijo.SaveAs(docHijo.PathName)
                    End If
                    Dim fileIn As String = docHijo.PathName
                    Dim fileOut As String = crDicFicheros(fileIn)
                    Fichero_EncriptaConBackupsTemp(fileIn, fileOut)
                    If crFicherosEnUso.Contains(fileIn) = False Then crFicherosEnUso.Add(fileIn)
                End If
            End If
        Next
        '
        mainDoc = Nothing
        clsCR.crPuedoborrar = True
    End Sub

#Region "Encripta en Temporal"
    ' 2018/11/05. Nuevas funciones para encriptar en tmplog y copiar en destino ya encriptado.
    ' Estas funciones siempre dan por hecho que el fichero desencriptado esta en tmplog y
    ' el fichero original está en cualquier otra carpeta del disco duro.
    ' Siempre encriptaremos una copia en tmplog y luego lo pasaremos a la ubicación original
    ' ** (Así nunca habrá un desencriptado en la ubicación original)
    ' FileIn = Fichero que tenemos en tmplog. Ej: tmplog\fichero.rvt
    ' FileInTmp = Fichero que copiamos de FileIn (En tmplog, con extensión .tempcrypto) y codificamos. Ej: tmplog\fichero.rvt.tempcrypto
    ' FileOut = Fichero original que cerramos, borramos y creamos con copia de FileInTmp (Poniendo extensión correcta) Ej: [ubicación original]\fichero.rvt
    Public Shared Sub Fichero_EncriptaConBackupsTemp(ByVal FileIn As String, FileOut As String, Optional conimagen As Boolean = False, Optional BorraBackups As Boolean = False)
        ' Si no está activado, salir ****
        If activado = False Then Exit Sub
        ' ********************************
        ' Solo encriptar si crEncrypt = True
        If crip2aCADUC.clsCR.crEncrypt = True Then
            ' Cerrar el documento original, si estaba abierto y no es el activo.
            Call Documento_EstaAbierto(crAppUI.Application, FileOut, cerrar:=True)
            ' Quitar sólo lectura del destino, ya que lo sobrescribiremos.
            clsCR.Fichero_ReadOnly(FileOut, sololectura:=False)
            'clsCR.Fichero_SoloLectura(FileOut, sololectura:=False)
            ' Ocultar temporalmente, para que lo lo vea el usuario
            'clsCR.Fichero_Hidden(FileOut, ocultar:=True)
            ' Crear el fichero .tempcrypto
            Dim suf As String = ".tempcrypto"
            Dim FileInTmp As String = FileIn & suf
            IO.File.Copy(FileIn, FileInTmp, True)
            ' Encriptamos copia en tmplog\[fichero.rvt].tempcrypto y la copiamos al destino original (Con StreamWrite)
            ' Al terminar la encriptación, borramos .tempcrypto
            Fichero_EncriptaTemp(FileInTmp, FileOut, conimagen)
            ' Borrar o Encriptar los Backups que tuviera el fichero destino.
            Fichero_BackupsEncriptaOBorraTemp(FileOut, BorraBackups)
            ' Mostrar el fichero final (Estaba oculto)
            'clsCR.Fichero_Hidden(FileOut, ocultar:=False)
            ' Borrar ficheros.
            'clsCR.crPuedoborrar = True
            'clsCR.FicherosDesarrollo_BorraTempTASK()
        End If
    End Sub
    ' No usar directamente. Usar mejor Fichero_EncriptaConBackupsTemp
    Private Shared Sub Fichero_EncriptaTemp(ByVal FileIn As String, FileOut As String, Optional conimagen As Boolean = False)
        ' Si no está activado, salir ****
        If activado = False Then Exit Sub
        ' ********************************
        If crip2aCADUC.clsCR.crEncrypt = False Then Exit Sub
        Dim yaencriptado As Boolean = False
        Dim esviejo As Boolean = False
        Dim enuso As Boolean = False
        Dim tipoEncript As TipoCod = Fichero_DameTipoEncriptacion(FileIn)
        ''
        Try
            fsInput = New System.IO.FileStream(FileIn, FileMode.Open,
                                                      FileAccess.ReadWrite)
            'FileOutImg = IO.Path.ChangeExtension(FileIn, ".png")
        Catch ex As Exception
            'MsgBox(FileIn & vbCrLf & vbCrLf & "Fichero en uso o sin permisos de lectura")
            enuso = True
            GoTo FINAL
            Exit Sub
        End Try
        Try
            fsOutput = New System.IO.FileStream(FileOut,
                                                   FileMode.Create,
                                                   FileAccess.Write)
        Catch ex As Exception
            'MsgBox(FileOut & vbCrLf & vbCrLf & "No se tiene permiso para borrar/crear fichero")
            enuso = True
            GoTo FINAL
            Exit Sub
        End Try
        ''
        '' Comprobar si es un fichero codificado por la aplicación o no
        Dim codificadoBloque(codificadorBinarioEnc.Count - 1) As Byte
        Dim intBytesIncodificado As Integer = fsInput.Read(codificadoBloque, 0, codificadorBinarioEnc.Count)
        '
        If tipoEncript = TipoCod.codificadoB Or tipoEncript = TipoCod.codificadoT Then
            yaencriptado = True
            GoTo FINAL
            Exit Sub
        ElseIf tipoEncript = TipoCod.ficheroenuso Or tipoEncript = TipoCod.noexistefichero Then
            yaencriptado = True
            GoTo FINAL
            Exit Sub
        Else
            'Fichero sin encriptar
            yaencriptado = False
            fsInput.Seek(0, SeekOrigin.Begin)
        End If
        '
        Try 'In case of errors.
            'Declare variables for encrypt/decrypt process.
            Dim bloque As Integer = 1024
            Dim bytBuffer(bloque) As Byte 'holds a block of bytes for processing (No bloques 0 ni 1)
            Dim bytBuffer0(bloque) As Byte 'holds a block of bytes for processing (Bloque 0 solo)
            Dim bytBuffer1(bloque) As Byte 'holds a block of bytes for processing (Bloque 1 solo)
            Dim bytBuffer2(bloque) As Byte 'holds a block of bytes for processing (Bloque 2 solo)
            Dim lngBytesProcessed As Long = 0 'running count of bytes processed
            Dim lngFileLength As Long = fsInput.Length 'the input file's length
            Do While (bloque * 3 + intBytesIncodificado) > lngFileLength
                bloque = CInt(bloque / 2)
            Loop
            Dim intBytesInCurrentBlock As Integer 'current bytes being processed
            Dim intBytesInCurrentBlock0 As Integer 'current bytes being processed
            Dim intBytesInCurrentBlock1 As Integer 'current bytes being processed
            Dim intBytesInCurrentBlock2 As Integer 'current bytes being processed
            ''
            'Use While hasta que se procese todo el fichero.
            Dim contador As Integer = 0
            While lngBytesProcessed < lngFileLength
                ''
                If contador = 0 Then
                    'Read file with the input filestream.
                    intBytesInCurrentBlock0 = fsInput.Read(bytBuffer0, 0, bloque)
                    intBytesInCurrentBlock = intBytesInCurrentBlock0
                ElseIf contador = 1 Then
                    'Read file with the input filestream.
                    intBytesInCurrentBlock1 = fsInput.Read(bytBuffer1, 0, bloque)
                    intBytesInCurrentBlock = intBytesInCurrentBlock1
                ElseIf contador = 2 Then
                    'Read file with the input filestream.
                    intBytesInCurrentBlock2 = fsInput.Read(bytBuffer2, 0, bloque)
                    intBytesInCurrentBlock = intBytesInCurrentBlock2
                ElseIf contador = 3 Then
                    fsOutput.Write(codificadorBinarioEnc, 0, intBytesIncodificado)
                    fsOutput.Write(bytBuffer2, 0, intBytesInCurrentBlock2)
                    fsOutput.Write(bytBuffer1, 0, intBytesInCurrentBlock1)
                    fsOutput.Write(bytBuffer0, 0, intBytesInCurrentBlock0)
                    'Read file with the input filestream.
                    intBytesInCurrentBlock = fsInput.Read(bytBuffer, 0, bloque)
                    'Write output file with the cryptostream.
                    fsOutput.Write(bytBuffer, 0, intBytesInCurrentBlock)
                    'intBytesInCurrentBlock += intBytesInCurrentBlock0 + intBytesInCurrentBlock1 + intBytesInCurrentBlock2
                    intBytesInCurrentBlock += intBytesIncodificado
                Else
                    'Read file with the input filestream.
                    intBytesInCurrentBlock = fsInput.Read(bytBuffer, 0, bloque)
                    'Write output file with the cryptostream.
                    fsOutput.Write(bytBuffer, 0, intBytesInCurrentBlock)
                End If
                ''
                'Update lngBytesProcessed
                If intBytesInCurrentBlock = 0 Then intBytesInCurrentBlock = bloque
                lngBytesProcessed = lngBytesProcessed + CLng(intBytesInCurrentBlock)
                ''
                contador += 1
            End While
        Catch ex As Exception
            ''
            'MsgBox(ex.Message)
        End Try
        ''
FINAL:
        'Close FileStreams and CryptoStream.
        If fsInput IsNot Nothing Then fsInput.Close()
        If fsOutput IsNot Nothing Then fsOutput.Close()
        fsInput = Nothing
        fsOutput = Nothing
        '
        ' Generar imagen previa
        Try
            If conimagen = True And IO.File.Exists(FileIn) And yaencriptado = False And esviejo = False Then
                Using clsS As New Storage(FileIn)
                    Dim queImagen As Bitmap = Nothing
                    queImagen = CType(clsS.ThumbnailImage.GetPreviewAsImage, Bitmap)
                    ''
                    If queImagen IsNot Nothing Then
                        Dim nuevoFi As String = IO.Path.ChangeExtension(FileIn, ".jpg")
                        If IO.File.Exists(nuevoFi) Then
                            IO.File.Delete(nuevoFi)
                        End If
                        ''
                        Using nBitmap As New Bitmap(queImagen)
                            nBitmap.Save(nuevoFi, Imaging.ImageFormat.Jpeg)
                        End Using
                    End If
                End Using
            End If
        Catch ex As Exception
            Debug.Print(ex.ToString)
        End Try
        ''
        Try
            '' Borrar el fichero original (.tmpcrypto)
            Fichero_ReadOnly(FileIn, sololectura:=False)
            IO.File.Delete(FileIn)
        Catch ex As Exception
            'MsgBox("Fichero_EncriptaTemp: Delete FileIn, error", MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    Private Shared Sub Fichero_BackupsEncriptaOBorraTemp(FileOut As String, borrar As Boolean)
        ' Si no está activado, salir ****
        If activado = False Then Exit Sub
        ' ********************************
        If FileOut = "" OrElse IO.File.Exists(FileOut) = False Then Exit Sub
        Dim queDir As String = IO.Path.GetDirectoryName(FileOut)
        Dim queNom As String = IO.Path.GetFileNameWithoutExtension(FileOut)
        Dim exten As String = IO.Path.GetExtension(FileOut)
        Dim mascara = queNom & ".????" & exten
        Dim colFi As String() = IO.Directory.GetFiles(queDir, mascara, SearchOption.TopDirectoryOnly)
        If colFi IsNot Nothing AndAlso colFi.Count > 0 Then
            For Each queFi As String In colFi
                If borrar Then
                    Try
                        IO.File.Delete(queFi)
                    Catch ex As Exception
                        Debug.Print(ex.ToString)
                    End Try
                Else
                    Fichero_Encripta(queFi, False)
                End If
            Next
        End If
    End Sub
#End Region
    '
#Region "Encripta en destino"
    Public Shared Sub Fichero_EncriptaConBackups(ByVal FileIn As String, Optional conimagen As Boolean = False, Optional BorraBackups As Boolean = False)
        ' Si no está activado, salir ****
        If activado = False Then Exit Sub
        ' ********************************
        ' Solo encriptar si crEncrypt = True
        If crip2aCADUC.clsCR.crEncrypt = True Then
            Fichero_Encripta(FileIn, conimagen)
            Fichero_BackupsEncriptaOBorra(FileIn, BorraBackups)
        End If
    End Sub
    ' No usar directamente. Usar mejor Fichero_EncriptaConBackups
    Private Shared Sub Fichero_Encripta(ByVal FileIn As String, Optional conimagen As Boolean = False)
        ' Si no está activado, salir ****
        If activado = False Then Exit Sub
        ' ********************************
        If crip2aCADUC.clsCR.crEncrypt = False Then Exit Sub
        Dim FileOut As String = ""
        Dim suf As String = ".tempcrypto"
        Dim yaencriptado As Boolean = False
        Dim esviejo As Boolean = False
        Dim enuso As Boolean = False
        Dim tipoEncript As TipoCod = Fichero_DameTipoEncriptacion(FileIn)
        ''
        Try
            fsInput = New System.IO.FileStream(FileIn, FileMode.Open,
                                                      FileAccess.ReadWrite)
            FileOut = FileIn & suf
            'FileOutImg = IO.Path.ChangeExtension(FileIn, ".png")
        Catch ex As Exception
            'MsgBox(FileIn & vbCrLf & vbCrLf & "Fichero en uso o sin permisos de lectura")
            enuso = True
            GoTo FINAL
            Exit Sub
        End Try
        Try
            fsOutput = New System.IO.FileStream(FileOut,
                                                   FileMode.Create,
                                                   FileAccess.Write)
        Catch ex As Exception
            'MsgBox(FileOut & vbCrLf & vbCrLf & "No se tiene permiso para borrar/crear fichero")
            enuso = True
            GoTo FINAL
            Exit Sub
        End Try
        ''
        '' Comprobar si es un fichero codificado por la aplicación o no
        Dim codificadoBloque(codificadorBinarioEnc.Count - 1) As Byte
        Dim intBytesIncodificado As Integer = fsInput.Read(codificadoBloque, 0, codificadorBinarioEnc.Count)
        '
        If tipoEncript = TipoCod.codificadoB Or tipoEncript = TipoCod.codificadoT Then
            yaencriptado = True
            GoTo FINAL
            Exit Sub
        ElseIf tipoEncript = TipoCod.ficheroenuso Or tipoEncript = TipoCod.noexistefichero Then
            yaencriptado = True
            GoTo FINAL
            Exit Sub
        Else
            'Fichero sin encriptar
            yaencriptado = False
            fsInput.Seek(0, SeekOrigin.Begin)
        End If
        '
        Try 'In case of errors.
            'Declare variables for encrypt/decrypt process.
            Dim bloque As Integer = 1024
            Dim bytBuffer(bloque) As Byte 'holds a block of bytes for processing (No bloques 0 ni 1)
            Dim bytBuffer0(bloque) As Byte 'holds a block of bytes for processing (Bloque 0 solo)
            Dim bytBuffer1(bloque) As Byte 'holds a block of bytes for processing (Bloque 1 solo)
            Dim bytBuffer2(bloque) As Byte 'holds a block of bytes for processing (Bloque 2 solo)
            Dim lngBytesProcessed As Long = 0 'running count of bytes processed
            Dim lngFileLength As Long = fsInput.Length 'the input file's length
            Do While (bloque * 3 + intBytesIncodificado) > lngFileLength
                bloque = CInt(bloque / 2)
            Loop
            Dim intBytesInCurrentBlock As Integer 'current bytes being processed
            Dim intBytesInCurrentBlock0 As Integer 'current bytes being processed
            Dim intBytesInCurrentBlock1 As Integer 'current bytes being processed
            Dim intBytesInCurrentBlock2 As Integer 'current bytes being processed
            ''
            'Use While hasta que se procese todo el fichero.
            Dim contador As Integer = 0
            While lngBytesProcessed < lngFileLength
                ''
                If contador = 0 Then
                    'Read file with the input filestream.
                    intBytesInCurrentBlock0 = fsInput.Read(bytBuffer0, 0, bloque)
                    intBytesInCurrentBlock = intBytesInCurrentBlock0
                ElseIf contador = 1 Then
                    'Read file with the input filestream.
                    intBytesInCurrentBlock1 = fsInput.Read(bytBuffer1, 0, bloque)
                    intBytesInCurrentBlock = intBytesInCurrentBlock1
                ElseIf contador = 2 Then
                    'Read file with the input filestream.
                    intBytesInCurrentBlock2 = fsInput.Read(bytBuffer2, 0, bloque)
                    intBytesInCurrentBlock = intBytesInCurrentBlock2
                ElseIf contador = 3 Then
                    fsOutput.Write(codificadorBinarioEnc, 0, intBytesIncodificado)
                    fsOutput.Write(bytBuffer2, 0, intBytesInCurrentBlock2)
                    fsOutput.Write(bytBuffer1, 0, intBytesInCurrentBlock1)
                    fsOutput.Write(bytBuffer0, 0, intBytesInCurrentBlock0)
                    'Read file with the input filestream.
                    intBytesInCurrentBlock = fsInput.Read(bytBuffer, 0, bloque)
                    'Write output file with the cryptostream.
                    fsOutput.Write(bytBuffer, 0, intBytesInCurrentBlock)
                    'intBytesInCurrentBlock += intBytesInCurrentBlock0 + intBytesInCurrentBlock1 + intBytesInCurrentBlock2
                    intBytesInCurrentBlock += intBytesIncodificado
                Else
                    'Read file with the input filestream.
                    intBytesInCurrentBlock = fsInput.Read(bytBuffer, 0, bloque)
                    'Write output file with the cryptostream.
                    fsOutput.Write(bytBuffer, 0, intBytesInCurrentBlock)
                End If
                ''
                'Update lngBytesProcessed
                If intBytesInCurrentBlock = 0 Then intBytesInCurrentBlock = bloque
                lngBytesProcessed = lngBytesProcessed + CLng(intBytesInCurrentBlock)
                ''
                contador += 1
            End While
        Catch ex As Exception
            ''
            'MsgBox(ex.Message)
        End Try
        ''
FINAL:
        'Close FileStreams and CryptoStream.
        If fsInput IsNot Nothing Then fsInput.Close()
        If fsOutput IsNot Nothing Then fsOutput.Close()
        ''
        '' Generar imagen previa
        Try
            If conimagen = True And IO.File.Exists(FileIn) And yaencriptado = False And esviejo = False Then
                Using clsS As New Storage(FileIn)
                    Dim queImagen As Bitmap = Nothing
                    queImagen = CType(clsS.ThumbnailImage.GetPreviewAsImage, Bitmap)
                    ''
                    If queImagen IsNot Nothing Then
                        Dim nuevoFi As String = IO.Path.ChangeExtension(FileIn, ".jpg")
                        If IO.File.Exists(nuevoFi) Then
                            IO.File.Delete(nuevoFi)
                        End If
                        ''
                        Using nBitmap As New Bitmap(queImagen)
                            nBitmap.Save(nuevoFi, Imaging.ImageFormat.Jpeg)
                        End Using
                    End If
                End Using
            End If
        Catch ex As Exception
            Debug.Print(ex.ToString)
        End Try
        ''
        Try
            If yaencriptado = False And IO.File.Exists(FileOut) Then
                '' Borrar el fichero original y poner en su lugar en encriptado
                Fichero_ReadOnly(FileIn, sololectura:=False)
                IO.File.Copy(FileOut, FileIn, True)
                IO.File.Delete(FileOut)
            ElseIf yaencriptado = True And IO.File.Exists(FileOut) Then
                '' Borrar el fichero encriptado
                IO.File.Delete(FileOut)
            End If
        Catch ex As Exception
            MsgBox("Fichero_Encripta: Error in Copy o Delete FileOut", MsgBoxStyle.Critical, "ERROR")
        End Try
        '
        fsInput = Nothing
        fsOutput = Nothing
    End Sub
    Private Shared Sub Fichero_BackupsEncriptaOBorra(FileIn As String, borrar As Boolean)
        ' Si no está activado, salir ****
        If activado = False Then Exit Sub
        ' ********************************
        If FileIn = "" OrElse IO.File.Exists(FileIn) = False Then Exit Sub
        Dim queDir As String = IO.Path.GetDirectoryName(FileIn)
        Dim queNom As String = IO.Path.GetFileNameWithoutExtension(FileIn)
        Dim exten As String = IO.Path.GetExtension(FileIn)
        Dim mascara = queNom & ".????" & exten
        Dim colFi As String() = IO.Directory.GetFiles(queDir, mascara, SearchOption.TopDirectoryOnly)
        If colFi IsNot Nothing AndAlso colFi.Count > 0 Then
            For Each queFi As String In colFi
                If borrar Then
                    Try
                        IO.File.Delete(queFi)
                    Catch ex As Exception
                        Debug.Print(ex.ToString)
                    End Try
                Else
                    Fichero_Encripta(queFi, False)
                End If
            Next
        End If
    End Sub
#End Region
    '
    Public Shared Sub Fichero_Desencripta(ByVal FileIn As String, Optional borrarimagen As Boolean = False)
        ' Si no está activado, salir ****
        If activado = False Then Exit Sub
        ' ********************************
        If crip2aCADUC.clsCR.crEncrypt = False Then Exit Sub
        Dim FileOut As String = ""
        Dim suf As String = ".tempcrypto"
        Dim yaencriptado As Boolean = False
        Dim esviejo As Boolean = False
        Dim tipoEncript As TipoCod = Fichero_DameTipoEncriptacion(FileIn)
        ''
        'FileStream para fichero de entrada (FileIn) y de salida (FileOut)
        Try
            fsInput = New System.IO.FileStream(FileIn, FileMode.Open,
                                                      FileAccess.ReadWrite)
            FileOut = FileIn & suf
            If borrarimagen Then
                Dim formatosimagen() As String = New String() {".png", ".jpg", ".bmp"}
                For Each queFormato As String In formatosimagen
                    Dim FileOutImg = IO.Path.ChangeExtension(FileIn, queFormato)
                    If IO.File.Exists(FileOutImg) Then
                        Try
                            IO.File.Delete(FileOutImg)
                        Catch ex As Exception
                            Continue For
                        End Try
                    End If
                Next
            End If
        Catch ex As Exception
            'MsgBox(FileIn & vbCrLf & vbCrLf & "Fichero en uso o sin permisos de lectura")
            Exit Sub
        End Try
        ''
        Try
            fsOutput = New System.IO.FileStream(FileOut,
                                                   FileMode.Create,
                                                   FileAccess.Write)
        Catch ex As Exception
            'MsgBox(FileOut & vbCrLf & vbCrLf & "No se tiene permiso para borrar/crear fichero")
            Exit Sub
        End Try
        '
        '' Comprobar si es un fichero codificado por la aplicación o no
        Dim codificadoBloque(codificadorBinarioEnc.Count - 1) As Byte
        Dim intBytesIncodificado As Integer = fsInput.Read(codificadoBloque, 0, codificadorBinarioEnc.Count)
        '
        If tipoEncript = TipoCod.codificadoB Or tipoEncript = TipoCod.codificadoT Then
            yaencriptado = True
            ' No hay que ponerse al principio del fichero. Leemos desde codificador en adelante.
            'fsInput.Seek(0, SeekOrigin.Begin)
        ElseIf tipoEncript = TipoCod.ficheroenuso Or tipoEncript = TipoCod.noexistefichero Then
            ' No hacer nada con el fichero
            yaencriptado = True
            GoTo FINAL
            Exit Sub
        ElseIf tipoEncript = TipoCod.noexistefichero Then
            ' Para que borre el encriptado temporal
            esviejo = True
            GoTo FINAL
            Exit Sub
        Else
            'Fichero sin encriptar
            yaencriptado = False
            ' Ponernos al principio del fichero
            fsInput.Seek(0, SeekOrigin.Begin)
            GoTo FINAL
            Exit Sub
        End If
        '
        '' Comprobar si es un fichero codificado por la aplicación o no
        'Dim codificadoBloque(codificadorBinarioEnc.Count - 1) As Byte
        'Dim intBytesIncodificado As Integer = fsInput.Read(codificadoBloque, 0, codificadorBinarioEnc.Count)
        'Dim valor As String = Encoding.ASCII.GetString(codificadoBloque)
        'If valor = codificador Then  ' Nueva encriptación.
        '    yaencriptado = True
        '    '' No hay que ponerse al principio del fichero. Leemos desde valor en adelante.
        '    'fsInput.Seek(0, SeekOrigin.Begin)
        'Else
        '    yaencriptado = False
        '    '' Ponernos al principio del fichero
        '    fsInput.Seek(0, SeekOrigin.Begin)
        '    GoTo FINAL
        '    Exit Sub
        'End If

        ''
        'fsOutput.SetLength(0) 'make sure fsOutput is empty
        Try 'In case of errors.
            'Declare variables for encrypt/decrypt process.
            Dim bloque As Integer = 1024
            Dim bytBuffer(bloque) As Byte 'holds a block of bytes for processing (No bloques 0 ni 1)
            Dim bytBuffer0(bloque) As Byte 'holds a block of bytes for processing (Bloque 0 solo)
            Dim bytBuffer1(bloque) As Byte 'holds a block of bytes for processing (Bloque 1 solo)
            Dim bytBuffer2(bloque) As Byte 'holds a block of bytes for processing (Bloque 1 solo)
            Dim lngBytesProcessed As Long = 0 'running count of bytes processed
            Dim lngFileLength As Long = fsInput.Length 'the input file's length
            Do While (bloque * 3 + intBytesIncodificado) > lngFileLength
                bloque = CInt(bloque / 2)
            Loop
            Dim intBytesInCurrentBlock As Integer 'current bytes being processed
            Dim intBytesInCurrentBlock0 As Integer 'current bytes being processed
            Dim intBytesInCurrentBlock1 As Integer 'current bytes being processed
            Dim intBytesInCurrentBlock2 As Integer 'current bytes being processed
            ''
            'Use While hasta que se procese todo el fichero.
            Dim contador As Integer = 0
            While lngBytesProcessed < lngFileLength
                ''
                If contador = 0 Then
                    'Read file with the input filestream.
                    intBytesInCurrentBlock0 = fsInput.Read(bytBuffer0, 0, bloque)
                    intBytesInCurrentBlock = intBytesInCurrentBlock0
                ElseIf contador = 1 Then
                    'Read file with the input filestream.
                    intBytesInCurrentBlock1 = fsInput.Read(bytBuffer1, 0, bloque)
                    intBytesInCurrentBlock = intBytesInCurrentBlock1
                ElseIf contador = 2 Then
                    'Read file with the input filestream.
                    intBytesInCurrentBlock2 = fsInput.Read(bytBuffer2, 0, bloque)
                    intBytesInCurrentBlock = intBytesInCurrentBlock2
                ElseIf contador = 3 Then
                    fsOutput.Write(bytBuffer2, 0, intBytesInCurrentBlock2)
                    fsOutput.Write(bytBuffer1, 0, intBytesInCurrentBlock1)
                    fsOutput.Write(bytBuffer0, 0, intBytesInCurrentBlock0)
                    'Read file with the input filestream.
                    intBytesInCurrentBlock = fsInput.Read(bytBuffer, 0, bloque)
                    'Write output file with the cryptostream.
                    fsOutput.Write(bytBuffer, 0, intBytesInCurrentBlock)
                    'intBytesInCurrentBlock += intBytesInCurrentBlock0 + intBytesInCurrentBlock1 + intBytesInCurrentBlock2
                Else
                    'Read file with the input filestream.
                    intBytesInCurrentBlock = fsInput.Read(bytBuffer, 0, bloque)
                    'Write output file with the cryptostream.
                    fsOutput.Write(bytBuffer, 0, intBytesInCurrentBlock)
                End If
                ''
                'Update lngBytesProcessed
                If intBytesInCurrentBlock = 0 Then intBytesInCurrentBlock = bloque
                lngBytesProcessed = lngBytesProcessed + CLng(intBytesInCurrentBlock)
                ''
                contador += 1
            End While
        Catch ex As Exception
            ''
            'MsgBox(ex.Message)
        End Try
        ''
FINAL:
        'Close FileStreams and CryptoStream.
        fsInput.Close()
        fsOutput.Close()
        '' Imagen previa
        Try
            If yaencriptado = True And IO.File.Exists(FileOut) Then
                '' Borrar el fichero original y poner en su lugar en encriptado
                Fichero_ReadOnly(FileIn, sololectura:=False)
                IO.File.Copy(FileOut, FileIn, True)
                IO.File.Delete(FileOut)
            ElseIf (yaencriptado = False And IO.File.Exists(FileOut)) Or esviejo = True Then
                '' Borrar el fichero encriptado
                IO.File.Delete(FileOut)
            End If
        Catch ex As Exception
            MsgBox("Fichero_Desencripta: Error in Copy o Delete FileOut", MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub
    ' Utilizar este para las familias revit a cargar.
    Public Shared Function Fichero_DesencriptaEnTempDame(ByVal FileIn As String) As String
        ' Si no está activado, salir ****
        If activado = False Then Return "" : Exit Function
        ' ********************************
        Dim FileOut As String = ""
        Dim yaencriptado As Boolean = False
        Dim esviejo As Boolean = False
        '
        If crip2aCADUC.clsCR.crEncrypt = False Then
            Return FileIn   ' FileOut
            Exit Function
        End If
        ' Si fichero en uso. Devolvemos el original, como esté.
        Dim tipoEncript As TipoCod = Fichero_DameTipoEncriptacion(FileIn)
        'If tipoEncript = TipoCod.ficheroenuso Then
        '    System.Windows.Forms.MessageBox.Show("NOTICE TO USER", "File in use or do not have permissions")
        '    Return FileIn
        '    Exit Function
        'End If
        '
        'FileStream para fichero de entrada (FileIn) y de salida (FileOut)
        Try
            fsInput = New System.IO.FileStream(FileIn, FileMode.Open,
                                                      FileAccess.Read)
            'FileOut = IO.Path.Combine(folder_TempRevit, IO.Path.GetFileName(FileIn))
            FileOut = Fichero_PonNuevo(FileIn, False)
        Catch ex As Exception
            System.Windows.Forms.MessageBox.Show("NOTICE TO USER", "You do not have permission to read the file" & vbCrLf & FileIn)
            'MsgBox(FileIn & vbCrLf & vbCrLf & "Fichero en uso o sin permisos de lectura")
            If fsInput IsNot Nothing Then fsInput.Close()
            Return ""
            Exit Function
        End Try
        ''
        Try
            fsOutput = New System.IO.FileStream(FileOut,
                                                   FileMode.Create,
                                                   FileAccess.Write)
        Catch ex As Exception
            Return ""   ' FileOut
            Exit Function
        End Try
        ''
        '' Comprobar si es un fichero codificado por la aplicación o no
        Dim codificadoBloque(codificadorBinarioEnc.Count - 1) As Byte
        Dim intBytesIncodificado As Integer = fsInput.Read(codificadoBloque, 0, codificadorBinarioEnc.Count)
        '
        If tipoEncript = TipoCod.codificadoB Or tipoEncript = TipoCod.codificadoT Then
            yaencriptado = True
        ElseIf tipoEncript = TipoCod.ficheroenuso Then
            yaencriptado = False
            GoTo FINAL
            Exit Function
        ElseIf tipoEncript = TipoCod.noexistefichero Then
            yaencriptado = True
            If fsInput IsNot Nothing Then fsInput.Close()
            If fsOutput IsNot Nothing Then fsOutput.Close()
            Return ""
            Exit Function
        Else
            'Fichero sin encriptar
            yaencriptado = False
            ' Ponernos al principio del fichero
            fsInput.Seek(0, SeekOrigin.Begin)
            GoTo FINAL
            Exit Function
        End If
        '
        '' Comprobar si es un fichero codificado por la aplicación o no
        'Dim codificadoBloque(codificadorBinarioEnc.Count - 1) As Byte
        'Dim intBytesIncodificado As Integer = fsInput.Read(codificadoBloque, 0, codificadorBinarioEnc.Count)
        'Dim valor As String = Encoding.ASCII.GetString(codificadoBloque)
        'If valor = codificador Then  ' Nueva encriptación. Or valor = codbinario1 Or valor = codtexto Then
        '    yaencriptado = True
        'Else
        '    yaencriptado = False
        '    '' Ponernos al principio del fichero
        '    fsInput.Seek(0, SeekOrigin.Begin)
        '    GoTo FINAL
        '    Exit Function
        'End If

        ''
        'fsOutput.SetLength(0) 'make sure fsOutput is empty
        Try 'In case of errors.
            'Declare variables for encrypt/decrypt process.
            Dim bloque As Integer = 1024
            Dim bytBuffer(bloque) As Byte 'holds a block of bytes for processing (No bloques 0 ni 1)
            Dim bytBuffer0(bloque) As Byte 'holds a block of bytes for processing (Bloque 0 solo)
            Dim bytBuffer1(bloque) As Byte 'holds a block of bytes for processing (Bloque 1 solo)
            Dim bytBuffer2(bloque) As Byte 'holds a block of bytes for processing (Bloque 1 solo)
            Dim lngBytesProcessed As Long = 0 'running count of bytes processed
            Dim lngFileLength As Long = fsInput.Length 'the input file's length
            Do While (bloque * 3 + intBytesIncodificado) > lngFileLength
                bloque = CInt(bloque / 2)
            Loop
            Dim intBytesInCurrentBlock As Integer 'current bytes being processed
            Dim intBytesInCurrentBlock0 As Integer 'current bytes being processed
            Dim intBytesInCurrentBlock1 As Integer 'current bytes being processed
            Dim intBytesInCurrentBlock2 As Integer 'current bytes being processed
            ''
            'Use While hasta que se procese todo el fichero.
            Dim contador As Integer = 0
            While lngBytesProcessed < lngFileLength
                ''
                If contador = 0 Then
                    'Read file with the input filestream.
                    intBytesInCurrentBlock0 = fsInput.Read(bytBuffer0, 0, bloque)
                    intBytesInCurrentBlock = intBytesInCurrentBlock0
                ElseIf contador = 1 Then
                    'Read file with the input filestream.
                    intBytesInCurrentBlock1 = fsInput.Read(bytBuffer1, 0, bloque)
                    intBytesInCurrentBlock = intBytesInCurrentBlock1
                ElseIf contador = 2 Then
                    'Read file with the input filestream.
                    intBytesInCurrentBlock2 = fsInput.Read(bytBuffer2, 0, bloque)
                    intBytesInCurrentBlock = intBytesInCurrentBlock2
                ElseIf contador = 3 Then
                    fsOutput.Write(bytBuffer2, 0, intBytesInCurrentBlock2)
                    fsOutput.Write(bytBuffer1, 0, intBytesInCurrentBlock1)
                    fsOutput.Write(bytBuffer0, 0, intBytesInCurrentBlock0)
                    'Read file with the input filestream.
                    intBytesInCurrentBlock = fsInput.Read(bytBuffer, 0, bloque)
                    'Write output file with the cryptostream.
                    fsOutput.Write(bytBuffer, 0, intBytesInCurrentBlock)
                    'intBytesInCurrentBlock += intBytesInCurrentBlock0 + intBytesInCurrentBlock1 + intBytesInCurrentBlock2
                Else
                    'Read file with the input filestream.
                    intBytesInCurrentBlock = fsInput.Read(bytBuffer, 0, bloque)
                    'Write output file with the cryptostream.
                    fsOutput.Write(bytBuffer, 0, intBytesInCurrentBlock)
                End If
                ''
                'Update lngBytesProcessed
                If intBytesInCurrentBlock = 0 Then intBytesInCurrentBlock = bloque
                lngBytesProcessed = lngBytesProcessed + CLng(intBytesInCurrentBlock)
                ''
                contador += 1
            End While
        Catch ex As Exception
            ''
            'MsgBox(ex.Message)
        End Try
        ''
FINAL:
        'Close FileStreams and CryptoStream.
        If fsInput IsNot Nothing Then fsInput.Close()
        If fsOutput IsNot Nothing Then fsOutput.Close()
        '
        If yaencriptado = False Then
            If IO.File.Exists(FileOut) Then Fichero_ReadOnly(FileOut, sololectura:=False)
            If IO.File.Exists(FileIn) Then IO.File.Copy(FileIn, FileOut, True)
            If IO.File.Exists(FileOut) Then Fichero_ReadOnly(FileOut, sololectura:=False)
            If tipoEncript = TipoCod.ficheroenuso Then Fichero_Desencripta(FileOut)
        End If
        'crfiAbrir = FileOut
        'crfiCerrar = FileIn
        fsInput = Nothing
        fsOutput = Nothing
        '
        Return FileOut
    End Function
    Public Shared Function Texto_Encripta(ByVal Input As String, Optional conCodTexto As Boolean = False) As String
        ' Si no está activado, salir ****
        If activado = False Then Return "" : Exit Function
        ' ********************************
        If Input = "" Or Input Is Nothing Then
            Return ""
            Exit Function
        End If
        'Dim IV() As Byte = ASCIIEncoding.ASCII.GetBytes(codificador)    ''La clave debe ser de 8 caracteres
        'Dim EncryptionKey() As Byte = Convert.FromBase64String("rpaSPvIvVLlrcmtzPU9/c67Gkj7yL1S5") 'No se puede alterar la cantidad de caracteres pero si la clave
        Dim buffer() As Byte = Encoding.UTF8.GetBytes(Input)
        Dim des As TripleDESCryptoServiceProvider = New TripleDESCryptoServiceProvider
        des.Key = EncryptionKey
        des.IV = IV
        ''
        If conCodTexto Then
            Return codificador & Convert.ToBase64String(des.CreateEncryptor().TransformFinalBlock(buffer, 0, buffer.Length()))
        Else
            Return Convert.ToBase64String(des.CreateEncryptor().TransformFinalBlock(buffer, 0, buffer.Length()))
        End If
    End Function
    ''
    Public Shared Function Texto_Desencriptar(ByVal Input As String, Optional conCodTexto As Boolean = False) As String
        ' Si no está activado, salir ****
        If activado = False Then Return "" : Exit Function
        ' ********************************
        Try
            'Dim IV() As Byte = ASCIIEncoding.ASCII.GetBytes(codificador)    ''La clave debe ser de 8 caracteres
            'Dim EncryptionKey() As Byte = Convert.FromBase64String("rpaSPvIvVLlrcmtzPU9/c67Gkj7yL1S5") 'No se puede alterar la cantidad de caracteres pero si la clave
            Dim buffer() As Byte = Nothing
            If conCodTexto Then
                buffer = Convert.FromBase64String(Input.Substring((codificador.Length), Input.Length - (codificador.Length)))
            Else
                buffer = Convert.FromBase64String(Input)
            End If
            ''
            Dim des As TripleDESCryptoServiceProvider = New TripleDESCryptoServiceProvider
            des.Key = EncryptionKey
            des.IV = IV
            ''
            Return Encoding.UTF8.GetString(des.CreateDecryptor().TransformFinalBlock(buffer, 0, buffer.Length()))
        Catch ex As Exception
            Return Input
        End Try
    End Function

    ''
    '' ***** Encriptar/Desencriptar desde un hacia/desde un fichero.
    Private Shared Sub Texto_EncriptaAFichero(ByVal Data As String, ByVal FileName As String, ByVal Key() As Byte, ByVal IV() As Byte)
        ' Si no está activado, salir ****
        If activado = False Then Exit Sub
        ' ********************************
        Try
            ' Create or open the specified file.
            Dim fStream As FileStream = File.Open(FileName, FileMode.OpenOrCreate)

            ' Create a CryptoStream using the FileStream 
            ' and the passed key and initialization vector (IV).
            Dim cStream As New CryptoStream(fStream,
                                            New TripleDESCryptoServiceProvider().CreateEncryptor(Key, IV),
                                            CryptoStreamMode.Write)

            ' Create a StreamWriter using the CryptoStream.
            Dim sWriter As New StreamWriter(cStream)

            ' Write the data to the stream 
            ' to encrypt it.
            sWriter.WriteLine(Data)

            ' Close the streams and
            ' close the file.
            sWriter.Close()
            cStream.Close()
            fStream.Close()
        Catch e As CryptographicException
            Console.WriteLine("A Cryptographic error occurred: {0}", e.Message)
        Catch e As UnauthorizedAccessException
            Console.WriteLine("A file error occurred: {0}", e.Message)
        End Try
    End Sub
    ''
    Private Shared Function Texto_DesencriptaDesdeFichero(ByVal FileName As String, ByVal Key() As Byte, ByVal IV() As Byte) As String
        ' Si no está activado, salir ****
        If activado = False Then Return "" : Exit Function
        ' ********************************
        Try
            ' Create or open the specified file. 
            Dim fStream As FileStream = File.Open(FileName, FileMode.OpenOrCreate)

            ' Create a CryptoStream using the FileStream 
            ' and the passed key and initialization vector (IV).
            Dim cStream As New CryptoStream(fStream,
                                            New TripleDESCryptoServiceProvider().CreateDecryptor(Key, IV),
                                            CryptoStreamMode.Read)

            ' Create a StreamReader using the CryptoStream.
            Dim sReader As New StreamReader(cStream)

            ' Read the data from the stream 
            ' to decrypt it.
            Dim val As String = sReader.ReadToEnd

            ' Close the streams and
            ' close the file.
            sReader.Close()
            cStream.Close()
            fStream.Close()

            ' Return the string. 
            Return val
        Catch e As CryptographicException
            Console.WriteLine("A Cryptographic error occurred: {0}", e.Message)
            Return Nothing
        Catch e As UnauthorizedAccessException
            Console.WriteLine("A file error occurred: {0}", e.Message)
            Return Nothing
        End Try
    End Function
    ''
    '' ***** Encriptar/Desencriptar desde la memoria.
    Public Shared Function Texto_EncriptaAMemoria(ByVal Data As String, ByVal Key() As Byte, ByVal IV() As Byte) As Byte()
        ' Si no está activado, salir ****
        If activado = False Then Return Nothing : Exit Function
        ' ********************************
        Try
            ' Create a MemoryStream.
            Dim mStream As New MemoryStream

            ' Create a CryptoStream using the MemoryStream 
            ' and the passed key and initialization vector (IV).
            Dim cStream As New CryptoStream(mStream,
                                            New TripleDESCryptoServiceProvider().CreateEncryptor(Key, IV),
                                            CryptoStreamMode.Write)

            ' Convert the passed string to a byte array.
            Dim toEncrypt As Byte() = New ASCIIEncoding().GetBytes(Data)

            ' Write the byte array to the crypto stream and flush it.
            cStream.Write(toEncrypt, 0, toEncrypt.Length)
            cStream.FlushFinalBlock()

            ' Get an array of bytes from the 
            ' MemoryStream that holds the 
            ' encrypted data.
            Dim ret As Byte() = mStream.ToArray()

            ' Close the streams.
            cStream.Close()
            mStream.Close()

            ' Return the encrypted buffer.
            Return ret
        Catch e As CryptographicException
            Console.WriteLine("A Cryptographic error occurred: {0}", e.Message)
            Return Nothing
        End Try
    End Function

    Public Shared Function Texto_DesencriptaDesdeMemoria(ByVal Data() As Byte, ByVal Key() As Byte, ByVal IV() As Byte) As String
        ' Si no está activado, salir ****
        If activado = False Then Return "" : Exit Function
        ' ********************************
        Try
            ' Create a new MemoryStream using the passed 
            ' array of encrypted data.
            Dim msDecrypt As New MemoryStream(Data)

            ' Create a CryptoStream using the MemoryStream 
            ' and the passed key and initialization vector (IV).
            Dim csDecrypt As New CryptoStream(msDecrypt,
                                              New TripleDESCryptoServiceProvider().CreateDecryptor(Key, IV),
                                              CryptoStreamMode.Read)

            ' Create buffer to hold the decrypted data.
            Dim fromEncrypt(Data.Length - 1) As Byte

            ' Read the decrypted data out of the crypto stream
            ' and place it into the temporary buffer.
            csDecrypt.Read(fromEncrypt, 0, fromEncrypt.Length)

            'Convert the buffer into a string and return it.
            Return New ASCIIEncoding().GetString(fromEncrypt)
        Catch e As CryptographicException
            Console.WriteLine("A Cryptographic error occurred: {0}", e.Message)
            Return Nothing
        End Try
    End Function
    ''
    '' ***** Encriptar/Desencriptar solo el valor del elemento XML
    Private Shared Function XML_DameDato(linea As String) As String
        ' Si no está activado, salir ****
        If activado = False Then Return "" : Exit Function
        ' ********************************
        ''    <description language="local">TRAVESSA BASE DE 1,00 Mts</description>
        ''    <description language="it-IT" />
        Dim resultado As String = ""
        If linea.EndsWith("/>") Or (linea.IndexOf(">") = linea.Length - 1) Then
            '' No hacemos nada, es un elemento sin texto.
            'Console.WriteLine(linea)
        ElseIf linea.Contains("</") And linea.EndsWith(">") Then
            Dim partes As String() = linea.Split(">"c)
            If partes IsNot Nothing AndAlso partes.Length > 2 Then
                Dim regex As System.Text.RegularExpressions.Regex = New System.Text.RegularExpressions.Regex("</")
                resultado = regex.Split(partes(1))(0)
            End If
        End If
        Return resultado
    End Function
    ''
    Public Shared Sub XML_Encripta(fiIn As String,
                            Optional borraIn As Boolean = True,
                            Optional ByRef pb1 As System.Windows.Forms.ProgressBar = Nothing)
        ' Si no está activado, salir ****
        If activado = False Then Exit Sub
        ' ********************************
        If fiIn = "" Or IO.File.Exists(fiIn) = False Then
            'MsgBox("El fichero " & fiIn & " no existe..." & vbCrLf & "Encriptación falla...")
            Exit Sub
        End If
        ''
        Dim fiOut As String = IO.Path.ChangeExtension(fiIn, ".xmlc")
        ''
        Dim objReader As New System.IO.StreamReader(fiIn)
        Dim objWriter As New System.IO.StreamWriter(fiOut)
        ''
        If pb1 IsNot Nothing Then
            pb1.Value = 0
            pb1.Maximum = CInt(objReader.BaseStream.Length)
        End If
        ''
        Dim sLine As String = ""
        ''
        While Not objReader.EndOfStream
            sLine = objReader.ReadLine()
            ''
            Dim tLinea As String = ""
            If sLine Is Nothing Or (sLine IsNot Nothing AndAlso sLine.Trim = "") Then
                Continue While
            Else
                Dim busco As String = XML_DameDato(sLine)
                If busco = "" Then
                    tLinea = sLine
                ElseIf busco <> "" Then
                    Dim cambio As String = Texto_Encripta(busco)
                    tLinea = sLine.Replace(busco, cambio)
                End If
            End If
            objWriter.WriteLine(tLinea)
            ''
            If pb1 IsNot Nothing Then
                pb1.Value = CInt(objReader.BaseStream.Position)
            End If
            System.Windows.Forms.Application.DoEvents()
        End While
        ''
        If pb1 IsNot Nothing Then
            pb1.Value = CInt(objReader.BaseStream.Length)
        End If
        'objWriter.Write(Chr(Asc(vbBack)))
        ''
        objReader.Close()
        objWriter.Close()
        objReader = Nothing
        objWriter = Nothing
        ''
        If borraIn = True Then
            Try
                IO.File.Delete(fiIn)
            Catch ex As Exception
                ''
            End Try
        End If
    End Sub
    ''
    Public Shared Sub XML_Desencripta(fiIn As String,
                               Optional borraIn As Boolean = True,
                               Optional ByRef pb1 As System.Windows.Forms.ProgressBar = Nothing)
        ' Si no está activado, salir ****
        If activado = False Then Exit Sub
        ' ********************************
        Dim yadesencriptado As Boolean = False
        Dim esviejo As Boolean = False
        If fiIn = "" Or IO.File.Exists(fiIn) = False Then
            MsgBox("El fichero " & fiIn & " no existe..." & vbCrLf & "Encriptación falla...")
            Exit Sub
        End If
        ''
        Dim fiOut As String = IO.Path.ChangeExtension(fiIn, ".xml")
        ''
        Dim objReader As New System.IO.StreamReader(fiIn)
        Dim objWriter As New System.IO.StreamWriter(fiOut)
        ''
        If pb1 IsNot Nothing Then
            pb1.Value = 0
            pb1.Maximum = CInt(objReader.BaseStream.Length)
        End If
        ''
        Dim sLine As String = ""
        Dim contador As Integer = 1
        ''
        While Not objReader.EndOfStream
            sLine = objReader.ReadLine()
            ''
            Dim tLinea As String = ""
            If sLine Is Nothing Or (sLine IsNot Nothing AndAlso sLine.Trim = "") Then
                Continue While
            Else
                '' Çomprobar si ya está desencriptado.
                If contador = 1 And sLine.StartsWith(codificador) = False Then 'sLine.StartsWith(preUnoNOEncriptado) And
                    yadesencriptado = True
                    Exit While
                Else
                    Dim busco As String = XML_DameDato(sLine)
                    If busco = "" Then
                        tLinea = sLine
                    ElseIf busco <> "" Then
                        If busco.StartsWith(codificador) = False Then
                            Dim cambio As String = Texto_Desencriptar(busco)
                            tLinea = sLine.Replace(busco, cambio)
                        End If
                    End If
                End If
            End If
            objWriter.WriteLine(tLinea)
            ''
            If pb1 IsNot Nothing Then
                pb1.Value = CInt(objReader.BaseStream.Position)
            End If
            System.Windows.Forms.Application.DoEvents()
            contador += 1
        End While
        ''
        If pb1 IsNot Nothing Then
            pb1.Value = CInt(objReader.BaseStream.Length)
        End If
        ''
        objReader.Close()
        objWriter.Close()
        objReader = Nothing
        objWriter = Nothing
        ''
        If borraIn = True And yadesencriptado = False Then
            Try
                IO.File.Delete(fiIn)
            Catch ex As Exception
                ''
            End Try
        End If
    End Sub
    ''
    Public Shared Sub XML_EncriptaLinea(fiIn As String,
                            Optional borraIn As Boolean = True,
                            Optional ByRef pb1 As System.Windows.Forms.ProgressBar = Nothing,
                               Optional fiOut As String = "")
        ' Si no está activado, salir ****
        If activado = False Then Exit Sub
        ' ********************************
        '
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>        (Primera línea de un fichero XML sin encriptar)
        'codtexto As String = "2ACGG.,t" (Primera linea de un fichero de texto)
        If fiIn = "" Or IO.File.Exists(fiIn) = False Then
            'MsgBox("El fichero " & fiIn & " no existe..." & vbCrLf & "Encriptación falla...")
            Exit Sub
        End If
        ''
        If fiOut = "" Then fiOut = IO.Path.ChangeExtension(fiIn, IO.Path.GetExtension(fiIn) + "c")
        ''
        Dim objReader As New System.IO.StreamReader(fiIn)
        Dim objWriter As New System.IO.StreamWriter(fiOut)
        ''
        If pb1 IsNot Nothing Then
            pb1.Value = 0
            pb1.Maximum = CInt(objReader.BaseStream.Length)
        End If
        ''
        Dim sLine As String = ""
        '
        Dim contador As Integer = 0
        While Not objReader.EndOfStream
            contador += 1
            sLine = objReader.ReadLine()
            ''
            Dim tLinea As String = ""
            If sLine Is Nothing Or (sLine IsNot Nothing AndAlso sLine.Trim = "") Then
                Continue While
            Else
                Dim busco As String = sLine.Trim
                If busco = "" Then
                    tLinea = sLine
                ElseIf busco <> "" Then
                    If contador = 1 Then
                        objWriter.WriteLine(codificador)
                    End If
                    tLinea = Texto_Encripta(busco)
                End If
            End If
            objWriter.WriteLine(tLinea)
            ''
            If pb1 IsNot Nothing Then
                pb1.Value = CInt(objReader.BaseStream.Position)
            End If
            System.Windows.Forms.Application.DoEvents()
        End While
        ''
        If pb1 IsNot Nothing Then
            pb1.Value = CInt(objReader.BaseStream.Length)
        End If
        'objWriter.Write(Chr(Asc(vbBack)))
        ''
        objReader.Close()
        objWriter.Close()
        objReader = Nothing
        objWriter = Nothing
        ''
        If borraIn = True Then
            Try
                IO.File.Delete(fiIn)
            Catch ex As Exception
                ''
            End Try
        End If
    End Sub
    '
    Public Shared Sub XML_DesencriptaLinea(fiIn As String,
                               Optional borraIn As Boolean = True,
                               Optional ByRef pb1 As System.Windows.Forms.ProgressBar = Nothing,
                               Optional fiOut As String = "")
        ' Si no está activado, salir ****
        If activado = False Then Exit Sub
        ' ********************************
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>        (Primera línea de un fichero XML sin encriptar)
        'codtexto As String = "2ACGG.,t" (Primera linea de un fichero de texto)
        If fiIn = "" Or IO.File.Exists(fiIn) = False Then
            MsgBox("El fichero " & fiIn & " no existe..." & vbCrLf & "Encriptación falla...")
            Exit Sub
        End If
        ''
        Dim extension As String = ""
        If fiOut = "" Then
            extension = IO.Path.GetExtension(fiIn)
            If extension.Length = 5 And extension.EndsWith("c") Then
                extension = extension.Substring(0, extension.Length - 1)
                fiOut = IO.Path.ChangeExtension(fiIn, extension)
            Else
                extension = ".crypto"
                fiOut = IO.Path.ChangeExtension(fiIn, extension)
            End If
        End If
        ''
        Dim objReader As New System.IO.StreamReader(fiIn)
        Dim objWriter As New System.IO.StreamWriter(fiOut)
        ''
        If pb1 IsNot Nothing Then
            pb1.Value = 0
            pb1.Maximum = CInt(objReader.BaseStream.Length)
        End If
        ''
        Dim sLine As String = ""
        Dim contador As Integer = 0
        While Not objReader.EndOfStream
            contador += 1
            sLine = objReader.ReadLine()
            ''
            Dim tLinea As String = ""
            If sLine Is Nothing Or (sLine IsNot Nothing AndAlso sLine.Trim = "") Or sLine.StartsWith(codificador) Then
                Continue While
            Else
                Dim busco As String = sLine
                If busco = "" Then
                    tLinea = sLine
                ElseIf busco <> "" Then
                    tLinea = Texto_Desencriptar(busco)
                End If
            End If
            objWriter.WriteLine(tLinea)
            ''
            If pb1 IsNot Nothing AndAlso CInt(objReader.BaseStream.Position) <= pb1.Maximum Then
                pb1.Value = CInt(objReader.BaseStream.Position)
            End If
            System.Windows.Forms.Application.DoEvents()
        End While
        ''
        If pb1 IsNot Nothing Then
            pb1.Value = CInt(objReader.BaseStream.Length)
        End If
        ''
        objReader.Close()
        objWriter.Close()
        objReader = Nothing
        objWriter = Nothing
        ''
        If borraIn = True Then
            Try
                IO.File.Delete(fiIn)
            Catch ex As Exception
                ''
            End Try
        End If
        '
        extension = IO.Path.GetExtension(fiOut)
        If extension = ".crypto" Then
            Fichero_ReadOnly(fiIn, sololectura:=False)
            IO.File.Copy(fiOut, fiIn, True)
            IO.File.Delete(fiOut)
        End If
    End Sub

    Public Shared Sub XML_EncriptaValue(fiIn As String,
                            Optional borraIn As Boolean = True,
                            Optional ByRef pb1 As System.Windows.Forms.ProgressBar = Nothing,
                               Optional fiOut As String = "")
        ' Si no está activado, salir ****
        If activado = False Then Exit Sub
        ' ********************************
        If fiIn = "" Or IO.File.Exists(fiIn) = False Then
            'MsgBox("El fichero " & fiIn & " no existe..." & vbCrLf & "Encriptación falla...")
            Exit Sub
        End If
        ''
        If fiOut = "" Then fiOut = IO.Path.ChangeExtension(fiIn, IO.Path.GetExtension(fiIn) + "c")
        ''
        Dim objReader As New System.IO.StreamReader(fiIn)
        Dim objWriter As New System.IO.StreamWriter(fiOut)
        ''
        If pb1 IsNot Nothing Then
            pb1.Value = 0
            pb1.Maximum = CInt(objReader.BaseStream.Length)
        End If
        ''
        Dim sLine As String = ""
        ''
        While Not objReader.EndOfStream
            sLine = objReader.ReadLine()
            ''
            Dim tLinea As String = ""
            If sLine Is Nothing Or (sLine IsNot Nothing AndAlso sLine.Trim = "") Then
                Continue While
            Else
                Dim busco As String = XML_DameDato(sLine).Trim
                If busco = "" Then
                    tLinea = sLine
                ElseIf busco <> "" Then
                    Dim cambio As String = Texto_Encripta(busco)
                    tLinea = sLine.Replace(">" & busco & "<", ">" & cambio & "<")
                End If
            End If
            objWriter.WriteLine(tLinea)
            ''
            If pb1 IsNot Nothing Then
                pb1.Value = CInt(objReader.BaseStream.Position)
            End If
            System.Windows.Forms.Application.DoEvents()
        End While
        ''
        If pb1 IsNot Nothing Then
            pb1.Value = CInt(objReader.BaseStream.Length)
        End If
        'objWriter.Write(Chr(Asc(vbBack)))
        ''
        objReader.Close()
        objWriter.Close()
        objReader = Nothing
        objWriter = Nothing
        ''
        If borraIn = True Then
            Try
                IO.File.Delete(fiIn)
            Catch ex As Exception
                ''
            End Try
        End If
    End Sub
    ''
    Public Shared Sub XML_DesencriptaValue(fiIn As String,
                               Optional borraIn As Boolean = True,
                               Optional ByRef pb1 As System.Windows.Forms.ProgressBar = Nothing,
                               Optional fiOut As String = "")
        ' Si no está activado, salir ****
        If activado = False Then Exit Sub
        ' ********************************
        If fiIn = "" Or IO.File.Exists(fiIn) = False Then
            MsgBox("El fichero " & fiIn & " no existe..." & vbCrLf & "Encriptación falla...")
            Exit Sub
        End If
        ''
        If fiOut = "" Then fiOut = IO.Path.ChangeExtension(fiIn, ".xml")
        ''
        Dim objReader As New System.IO.StreamReader(fiIn)
        Dim objWriter As New System.IO.StreamWriter(fiOut)
        ''
        If pb1 IsNot Nothing Then
            pb1.Value = 0
            pb1.Maximum = CInt(objReader.BaseStream.Length)
        End If
        ''
        Dim sLine As String = ""
        ''
        While Not objReader.EndOfStream
            sLine = objReader.ReadLine()
            ''
            Dim tLinea As String = ""
            If sLine Is Nothing Or (sLine IsNot Nothing AndAlso sLine.Trim = "") Or sLine.StartsWith(codificador) Then
                Continue While
            Else
                Dim busco As String = XML_DameDato(sLine)
                If busco = "" Then
                    tLinea = sLine
                ElseIf busco <> "" Then
                    Dim cambio As String = Texto_Desencriptar(busco)
                    tLinea = sLine.Replace(busco, cambio)
                End If
            End If
            objWriter.WriteLine(tLinea)
            ''
            If pb1 IsNot Nothing Then
                pb1.Value = CInt(objReader.BaseStream.Position)
            End If
            System.Windows.Forms.Application.DoEvents()
        End While
        ''
        If pb1 IsNot Nothing Then
            pb1.Value = CInt(objReader.BaseStream.Length)
        End If
        ''
        objReader.Close()
        objWriter.Close()
        objReader = Nothing
        objWriter = Nothing
        ''
        If borraIn = True Then
            Try
                IO.File.Delete(fiIn)
            Catch ex As Exception
                ''
            End Try
        End If
    End Sub

    Public Enum TipoCod
        sincodificarRevit
        sincodificarOtros
        codificadoB
        codificadoT
        ficheroenuso
        noexistefichero
    End Enum
End Class

