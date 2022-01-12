Imports System.Data.SqlClient
Imports System.IO
Imports EXO_Log
'Imports CrystalDecisions.CrystalReports.Engine
Imports System.Threading

Public Class EnvioFact

    Public Shared Sub EnvioFacturas(empresa As String, tipodoc As String)

        Try
            Dim oCompany As SAPbobsCOM.Company = Nothing
            Conexiones.Connect_Company(oCompany, empresa)

            Dim oDBHana As EXO_DIAPI.EXO_DIAPI = New EXO_DIAPI.EXO_DIAPI(oCompany)

            ListFiles(oDBHana, oCompany, empresa, tipodoc)

            EXO_CleanCOM.CLiberaCOM.liberaCOM(oDBHana)
            Conexiones.Disconnect_Company(oCompany)

        Catch ex As Exception

        End Try


    End Sub

    Public Shared Sub ListFiles(oDBHANA As EXO_DIAPI.EXO_DIAPI, oCompany As SAPbobsCOM.Company, empresa As String, tipodoc As String)


        Dim olog As EXO_Log.EXO_Log = Nothing
        Dim sRutaFicheros As String = ""
        Dim SRutaLog As String = ""
        Dim dtDocumentos As DataTable = New System.Data.DataTable()
        Dim sRuta As String = ""
        Dim sFicheroCrystal As String = ""
        Dim sTextoTipoDoc As String = ""
        Dim DocumentoPdf As String = ""
        Dim sCardCode As String = ""
        Dim fichlog As String = SRutaLog & "Log_" & Format(Now.Year, "0000") & Format(Now.Month, "00") & Format(Now.Day, "00") & "_" & Format(Now.Hour, "00") & Format(Now.Minute, "00") & Format(Now.Second, "00") & ".txt"

        Try

            sRuta = Conexiones._sRutaFicheros

            'Crear y validar carpetas, obtener rutas para los ficheros
            ValidarCarpetas(sRuta, empresa, tipodoc, SRutaLog)

            olog = New EXO_Log.EXO_Log(SRutaLog & fichlog, 50, EXO_Log.EXO_Log.GestionFichero.dia)

            'consulta para buscar documentos a realizar
            Dim sSQL As String = " select ""Code"",""U_EXO_DOCN"",""U_EXO_MAIL"",""U_EXO_NOMF"", (SELECT ""AttachPath"" from ""OADP"") as ""RutaFichero"",""CardCode"" " +
                " FROM ""@EXO_FACTJOB"" T0 inner join ""OINV"" T1 ON T0.""U_EXO_DOCN""=T1.""DocNum"" WHERE ""U_EXO_ENV""='N'"

            dtDocumentos = oDBHANA.SQL.executeQuery(sSQL)

            If dtDocumentos.Rows.Count > 0 Then

                olog.escribeMensaje("entramos en bucle", EXO_Log.EXO_Log.Tipo.informacion)

                For Each row As DataRow In dtDocumentos.Rows

                    Thread.Sleep(1000)
                    If EnviarEmail(row, empresa, oDBHANA, olog) = True Then
                        'todo ok, en la funcion ya hemos escrito en el log

                        sSQL = "UPDATE ""@EXO_FACTJOB"" SET ""U_EXO_ENV""='Y' WHERE ""Code"" = " & row.Item("Code").ToString()

                        oDBHANA.SQL.executeQuery(sSQL)


                    End If
                Next
            Else
                ' olog.escribeMensaje("gestinamos filas no hay registros pendientes.", EXO_Log.EXO_Log.Tipo.informacion)
            End If

        Catch ex As Exception

            olog.escribeMensaje("Exception ListFiles: " + ex.Message, EXO_Log.EXO_Log.Tipo.error)
        Finally

        End Try
    End Sub

    Private Shared Function EnviarEmail(FilaSeleccionada As DataRow, empresa As String, oDBHANA As EXO_DIAPI.EXO_DIAPI, olog As EXO_Log.EXO_Log) As Boolean

        Dim correo As New System.Net.Mail.MailMessage()
        Dim adjunto As System.Net.Mail.Attachment

        Dim StrFirma As String = ""
        Dim htmbody As New System.Text.StringBuilder()
        Dim cuerpo As String = ""

        Try
            If empresa = "SBOLANDE" Then
                correo.From = New System.Net.Mail.MailAddress("administracion@landesa.com", "Landesa")
            Else
                correo.From = New System.Net.Mail.MailAddress("administracion@landesa.com", "Landesa")
            End If

            'vuelvo a realizar la consulta y lo envio todo junto

            correo.To.Add(FilaSeleccionada.Item("U_EXO_MAIL").ToString.Replace(";", ""))
            olog.escribeMensaje("correo " + FilaSeleccionada.Item("U_EXO_MAIL").ToString.Replace(";", ""), EXO_Log.EXO_Log.Tipo.informacion)
            Dim dirad As String = FilaSeleccionada.Item("RutaFichero").ToString() & FilaSeleccionada.Item("U_EXO_NOMF").ToString().Replace(";", "")
            adjunto = New System.Net.Mail.Attachment(dirad)

            correo.Attachments.Add(adjunto)

            correo.Subject = "E-Factura " & FilaSeleccionada.Item("CardCode").ToString & " " + FilaSeleccionada.Item("U_EXO_DOCN").ToString

            cuerpo = "Estimado cliente, " + Chr(13) + Chr(13)

            cuerpo = cuerpo + "Adjunto factura emitida." + Chr(13)

            'If empresa = "SBOLANDE" Then
            '    cuerpo = cuerpo + "ESTAMOS MODIFICANDO LAS CUENTAS BANCARIAS, POR FAVOR REVISEN LA QUE APARECE EN LA FACTURA ANTES DEL REALIZAR EL PAGO. GRACIAS " + Chr(13)
            'End If

            cuerpo = cuerpo + "Saludos" + Chr(13) + Chr(13)

            cuerpo = cuerpo + "Silvia Aguirre " + Chr(13)
            cuerpo = cuerpo + "Departamento de Administración." + Chr(13) + Chr(13)

            cuerpo = cuerpo + "Tel +34 916 840 050" + Chr(13)
            cuerpo = cuerpo + "administracion@landesa.com" + Chr(13) + Chr(13)

            cuerpo = cuerpo + "www.landesa.com" + Chr(13) + Chr(13)

            cuerpo = cuerpo + "Antes de imprimir este mensaje, piensa en tu responsabilidad y compromiso con el Medio Ambiente." + Chr(13) + Chr(13)

            cuerpo = cuerpo + "AVISO LEGAL: Este mensaje y sus archivos adjuntos van dirigidos exclusivamente a su destinatario, pudiendo contener información confidencial sometida a secreto profesional. No está permitida su comunicación, reproducción o distribución sin la autorización expresa de LANDE, S.A.. Si usted no es el destinatario final, por favor elimínelo e infórmenos por esta vía." + Chr(13)

            cuerpo = cuerpo + "PROTECCIÓN DE DATOS: De conformidad con lo dispuesto en el Reglamento (UE) 2016/679 de 27 de abril (GDPR) y la Ley Orgánica 3/2018 de 5 de diciembre (LOPDGDD), le informamos que los datos personales y dirección de correo electrónico del interesado, serán tratados bajo la responsabilidad de LANDE, S.A. por un interés legítimo y para el envío de comunicaciones sobre nuestros productos y servicios y se conservarán mientras ninguna de las partes se oponga a ello. Los datos no serán comunicados a terceros, salvo obligación legal. Le informamos que puede ejercer los derechos de acceso, rectificación, portabilidad y supresión de sus datos y los de limitación y oposición a su tratamiento dirigiéndose a CALLE FUNDIDORES, 63 - 28906 GETAFE (Madrid). Si considera que el tratamiento no se ajusta a la normativa vigente, podrá presentar una reclamación ante la autoridad de control en www.aepd.es." + Chr(13)

            'cuerpo = cuerpo + "Lande S.A."

            correo.Body = cuerpo
            correo.IsBodyHtml = False
            correo.Priority = System.Net.Mail.MailPriority.Normal



            Dim smtp As New System.Net.Mail.SmtpClient



            smtp.Host = "smtp.office365.com"
            smtp.Port = 587
            smtp.UseDefaultCredentials = False
            smtp.Credentials = New System.Net.NetworkCredential(Conexiones._user.ToString, Conexiones._pass.ToString)
            smtp.EnableSsl = True


            'smtp.DeliveryMethod = Net.Mail.SmtpDeliveryMethod.Network



            smtp.Send(correo)
            correo.Dispose()

            olog.escribeMensaje("Correo enviado: " & FilaSeleccionada.Item("U_EXO_MAIL").ToString & " " & FilaSeleccionada.Item("U_EXO_DOCN").ToString, EXO_Log.EXO_Log.Tipo.informacion)

            Return True

        Catch ex As Exception

            EnviarEmail = False



            olog.escribeMensaje("Error enviando correo:  " & FilaSeleccionada.Item("U_EXO_MAIL").ToString & " " & FilaSeleccionada.Item("U_EXO_DOCN").ToString & ex.Message, EXO_Log.EXO_Log.Tipo.error)

        End Try

        Return False

    End Function

    Public Shared Sub GeneraAlerta(ByVal oCompany As SAPbobsCOM.Company, ByVal stitulo As String, ByVal sUsuario As String, ByVal strCliente As String, ByVal strDocNum As String, ByVal TipoDoc As String)


        Dim oDtDTW As System.Data.DataTable = Nothing
        Dim oDtSAP As System.Data.DataTable = Nothing
        Dim sSQL As String = ""
        Dim pMessageDataColumns As SAPbobsCOM.MessageDataColumns = Nothing
        Dim pMessageDataColumn As SAPbobsCOM.MessageDataColumn = Nothing

        Dim oLines As SAPbobsCOM.MessageDataLines = Nothing
        Dim oLine As SAPbobsCOM.MessageDataLine = Nothing
        Dim oRecipientCollection As SAPbobsCOM.RecipientCollection = Nothing
        Dim oCmpSrv As SAPbobsCOM.CompanyService = Nothing
        Dim oMessageService As SAPbobsCOM.MessagesService = Nothing
        Dim oMessage As SAPbobsCOM.Message = Nothing

        Try


            oCmpSrv = oCompany.GetCompanyService
            oMessageService = CType(oCmpSrv.GetBusinessService(SAPbobsCOM.ServiceTypes.MessagesService), SAPbobsCOM.MessagesService)
            oMessage = CType(oMessageService.GetDataInterface(SAPbobsCOM.MessagesServiceDataInterfaces.msdiMessage), SAPbobsCOM.Message)

            oMessage.Subject = stitulo

            oMessage.Text = "No se ha podido enviar correo electronico para el documento " + TipoDoc + " " + strDocNum

            oRecipientCollection = oMessage.RecipientCollection

            oRecipientCollection.Add()
            oRecipientCollection.Item(0).SendInternal = SAPbobsCOM.BoYesNoEnum.tYES
            oRecipientCollection.Item(0).UserCode = sUsuario

            'pMessageDataColumns = oMessage.MessageDataColumns

            'pMessageDataColumn = pMessageDataColumns.Add()
            'pMessageDataColumn.ColumnName = "Número interno"
            'pMessageDataColumn.Link = SAPbobsCOM.BoYesNoEnum.tYES
            'oLines = pMessageDataColumn.MessageDataLines
            'oLine = oLines.Add()
            'oLine.Value = strDocEntry
            'oLine.Object = "17"
            'oLine.ObjectKey = strDocEntry

            'pMessageDataColumn = pMessageDataColumns.Add()
            'pMessageDataColumn.ColumnName = "Número Pedido"
            'oLines = pMessageDataColumn.MessageDataLines
            'oLine = oLines.Add()
            'oLine.Value = strPedido

            oMessageService.SendMessage(oMessage)

        Catch exCOM As System.Runtime.InteropServices.COMException

        Catch ex As Exception

        Finally
            If oDtDTW IsNot Nothing Then oDtDTW.Dispose()
            If oDtSAP IsNot Nothing Then oDtSAP.Dispose()
            If pMessageDataColumns IsNot Nothing Then System.Runtime.InteropServices.Marshal.FinalReleaseComObject(pMessageDataColumns)
            If pMessageDataColumn IsNot Nothing Then System.Runtime.InteropServices.Marshal.FinalReleaseComObject(pMessageDataColumn)
            If oLines IsNot Nothing Then System.Runtime.InteropServices.Marshal.FinalReleaseComObject(oLines)
            If oLine IsNot Nothing Then System.Runtime.InteropServices.Marshal.FinalReleaseComObject(oLine)
            If oRecipientCollection IsNot Nothing Then System.Runtime.InteropServices.Marshal.FinalReleaseComObject(oRecipientCollection)
            If oCmpSrv IsNot Nothing Then System.Runtime.InteropServices.Marshal.FinalReleaseComObject(oCmpSrv)
            If oMessageService IsNot Nothing Then System.Runtime.InteropServices.Marshal.FinalReleaseComObject(oMessageService)
            If oMessage IsNot Nothing Then System.Runtime.InteropServices.Marshal.FinalReleaseComObject(oMessage)


        End Try
    End Sub

    Private Shared Sub ValidarCarpetas(sRuta As String, empresa As String, tipodoc As String, ByRef sRutaLog As String)

        Try

            Dim sRutaFinal As String = sRuta & "\" & empresa & "\"
            Dim srutafinal2 As String = ""

            srutafinal2 = sRutaFinal & "Log" & "\"
            If System.IO.Directory.Exists(srutafinal2) = False Then
                System.IO.Directory.CreateDirectory(srutafinal2)
            End If

            sRutaLog = srutafinal2


        Catch ex As Exception

        End Try


    End Sub


End Class


