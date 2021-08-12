Public Class TestForm

    Dim consultStatus As String

    Private Sub BtnGetToken_Click(sender As Object, e As EventArgs) Handles BtnGetToken.Click

        baseUrl = TxtBaseUrl.Text
        accessKey = TxtAccessKey.Text
        clientId = TxtClientId.Text

        If Authentication(accessKey, clientId) = True Then
            TxtAccessToken.Text = accessToken
        Else
            MsgBox("Falha ao obter token!" & vbCrLf & vbCrLf & returnWithError, vbCritical, "PenseAPI")
            Exit Sub
        End If

        MsgBox("Token obtido com sucesso", vbInformation, "PenseAPI")

    End Sub

    Private Sub BtnCreateStore_Click(sender As Object, e As EventArgs) Handles BtnCreateStore.Click

        Dim open As String = TxtOpen.Text
        Dim close As String = TxtClose.Text

        Dim companyName As String = TxtCompanyName.Text
        Dim document As String = TxtDocument.Text
        Dim externalReference As String = TxtExternalReference.Text

        Dim streetNumber As String = TxtStreetNumber.Text
        Dim streetName As String = TxtStreetName.Text
        Dim cityName As String = TxtCityName.Text
        Dim stateName As String = TxtStateName.Text
        Dim latitude As String = TxtLatitude.Text
        Dim longitude As String = TxtLongitude.Text
        Dim reference As String = txtReferene.Text

        Dim tradeName As String = TxtTradeName.Text

        If CreateStore(open, close, companyName, document, externalReference, streetNumber, streetName, cityName, stateName, latitude, longitude, reference, tradeName) = True Then
            MsgBox("Loja cadastrada com sucesso", vbInformation, "PenseAPI")
        Else
            MsgBox("Falha ao cadastrar loja!" & vbCrLf & vbCrLf & returnWithError, vbCritical, "PenseAPI")
            Exit Sub
        End If

    End Sub

    Private Sub BtnRegisterPos_Click(sender As Object, e As EventArgs) Handles BtnRegisterPos.Click

        Dim externalReference As String = TxtExternalReferencePos.Text
        Dim externalReferenceStore As String = TxtExternalReferenceStorePos.Text
        Dim name As String = TxtNamePos.Text

        If RegisterPOS(externalReference, externalReferenceStore, name) = True Then
            MsgBox("PDV cadastrado com sucesso", vbInformation, "PenseAPI")
        Else
            MsgBox("Falha ao cadastrar PDV!" & vbCrLf & vbCrLf & returnWithError, vbCritical, "PenseAPI")
            Exit Sub
        End If

    End Sub

    Private Sub TabControl1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles TabControl1.SelectedIndexChanged
        If TabControl1.SelectedIndex = 3 Then
            Me.Width = 862
        Else
            Me.Width = 506
        End If
    End Sub

    Private Function ItemList() As String

        'Construção da lista de itens que estará contida no body da requisição de pagamento
        Dim quotationMarks As String = """"
        Dim parameters As String = ""
        Dim itemDescription As String = ""
        Dim itemQuantity As String = ""
        Dim itemValue As String = ""
        Dim lines As Integer = 0

        parameters = quotationMarks & "items" & quotationMarks & ": " & "["

        If DgvItems.Rows.Count - 1 = 1 Then

            For i = 0 To DgvItems.Rows.Count - 2
                itemDescription = DgvItems.Rows(i).Cells(0).Value.ToString
                itemQuantity = DgvItems.Rows(i).Cells(1).Value.ToString
                itemValue = DgvItems.Rows(i).Cells(2).Value.ToString
                itemValue = itemValue.Replace(",", ".")
            Next

            parameters = parameters & "
{" & "
" & quotationMarks & "description" & quotationMarks & ": " & quotationMarks & itemDescription & quotationMarks & ",
" & quotationMarks & "quantity" & quotationMarks & ": " & quotationMarks & itemQuantity & quotationMarks & ",
" & quotationMarks & "value" & quotationMarks & ": " & quotationMarks & itemValue & quotationMarks & "
}"

        Else

            lines = DgvItems.Rows.Count - 2

            For i = 0 To DgvItems.Rows.Count - 2

                itemDescription = DgvItems.Rows(i).Cells(0).Value.ToString
                itemQuantity = DgvItems.Rows(i).Cells(1).Value.ToString
                itemValue = DgvItems.Rows(i).Cells(2).Value.ToString

                itemValue = itemValue.Replace(",", ".")

                If i = lines Then
                    parameters = parameters & "
{" & "
" & quotationMarks & "description" & quotationMarks & ": " & quotationMarks & itemDescription & quotationMarks & ",
" & quotationMarks & "quantity" & quotationMarks & ": " & quotationMarks & itemQuantity & quotationMarks & ",
" & quotationMarks & "value" & quotationMarks & ": " & quotationMarks & itemValue & quotationMarks & "
}"
                Else
                    parameters = parameters & "
{" & "
" & quotationMarks & "description" & quotationMarks & ": " & quotationMarks & itemDescription & quotationMarks & ",
" & quotationMarks & "quantity" & quotationMarks & ": " & quotationMarks & itemQuantity & quotationMarks & ",
" & quotationMarks & "value" & quotationMarks & ": " & quotationMarks & itemValue & quotationMarks & "
},
"
                End If

            Next

        End If

        parameters = parameters & "],"

        parameters = parameters.Replace(vbCrLf, "")

        Return parameters

    End Function

    Private Sub BtnGenerateSale_Click(sender As Object, e As EventArgs) Handles BtnGenerateSale.Click

        Dim description As String = TxtDescriptionPayment.Text
        Dim externalReference As String = TxtExternalReferencePayment.Text
        Dim value As String = TxtValuePayment.Text
        value = value.Replace(",", ".")
        Dim wallet As String = TxtWalletPayment.Text
        Dim store As String = TxtStorePayment.Text
        Dim pdv As String = TxtPdvPayment.Text
        Dim items As String = ItemList()

        If Payment(description, externalReference, value, wallet, store, pdv, items) = True Then
            MsgBox("Venda de número: " & paymentId & " foi registrada com sucesso", vbInformation, "PenseAPI")
            TxtPaymentId.Text = paymentId
            ImgQrCodeUrl.Load(qrCodeUrl)
            LblPaymentStatus.Text = ""
            LblUpdateAt.Text = ""
            consultStatus = ""
            TimerPaymentStatus.Start()
        Else
            MsgBox("Falha ao registrar venda!" & vbCrLf & vbCrLf & returnWithError, vbCritical, "PenseAPI")
            Exit Sub
        End If

    End Sub

    Private Sub TestForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        If TabControl1.SelectedIndex = 3 Then
            Me.Width = 862
        Else
            Me.Width = 506
        End If
    End Sub

    Private Sub TimerPaymentStatus_Tick(sender As Object, e As EventArgs) Handles TimerPaymentStatus.Tick

        If LblPaymentStatus.Text = "" Or LblPaymentStatus.Text = "Waiting" Then
            Try
                BwPaymentStatus.RunWorkerAsync()

                System.Threading.Thread.Sleep(100)
                Application.DoEvents()

                If consultStatus = "OK" Then
                    LblPaymentStatus.Text = paymentStatus
                    LblUpdateAt.Text = updateAt.AddHours(-3)
                End If

                If LblPaymentStatus.Text = "Waiting" Then
                    LblPaymentStatus.ForeColor = Color.Gold
                ElseIf LblPaymentStatus.Text = "Paid" Then
                    LblPaymentStatus.ForeColor = Color.Green
                Else
                    LblPaymentStatus.ForeColor = Color.Red
                End If
            Catch ex As Exception
                MsgBox(ex.Message.ToString)
            End Try
        Else
            TimerPaymentStatus.Stop()
        End If

        TimerPaymentStatus.Interval = 5000

    End Sub

    Private Sub BwPaymentStatus_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BwPaymentStatus.DoWork
        If GetPaymentStatus(paymentId) = True Then
            consultStatus = "OK"
        Else
            consultStatus = "ERROR"
        End If
    End Sub

    Private Sub BtnAddItem_Click(sender As Object, e As EventArgs) Handles BtnAddItem.Click
        DgvItems.Rows.Add(TxtItemDescriptionPayment.Text, TxtItemQuantityPayment.Text, TxtItemValuePayment.Text, Convert.ToDecimal(TxtItemQuantityPayment.Text) * Convert.ToDecimal(TxtItemValuePayment.Text) / 100)
        CalculateTotal()
    End Sub

    Private Sub CalculateTotal()

        Dim total As Double = 0

        For i = 0 To DgvItems.Rows.Count - 1
            total += Convert.ToDecimal(DgvItems.Rows(i).Cells(3).Value)
        Next

        TxtValuePayment.Text = total

    End Sub

    Private Sub BtmRemoveItem_Click(sender As Object, e As EventArgs) Handles BtmRemoveItem.Click
        DgvItems.Rows.RemoveAt(DgvItems.CurrentCell.RowIndex)
        CalculateTotal()
    End Sub

    Private Sub BtnConsultStatus_Click(sender As Object, e As EventArgs) Handles BtnConsultStatus.Click

        ConsultTransactionStatus(TxtPaymentIdConsult.Text, "paymentId")

    End Sub

    Private Sub ConsultTransactionStatus(ByVal information As Integer, ByVal type As String)

        LblPaymentStatusConsult.Text = ""
        LblUpdateAtConsult.Text = ""

        If type = "paymentId" Then

            If GetPaymentStatus(information) = True Then
                ImgQrCodeUrlConsult.Load(qrCodeUrl)
                LblPaymentStatusConsult.Text = paymentStatus
                LblUpdateAtConsult.Text = updateAt.AddHours(-3)
                If LblPaymentStatusConsult.Text = "Waiting" Then
                    LblPaymentStatusConsult.ForeColor = Color.Gold
                ElseIf LblPaymentStatusConsult.Text = "Paid" Then
                    LblPaymentStatusConsult.ForeColor = Color.Green
                Else
                    LblPaymentStatusConsult.ForeColor = Color.Red
                End If
            End If

        ElseIf type = "externalReference" Then

            If GetPaymentStatusByExternalReference(information) = True Then
                ImgQrCodeUrlConsult.Load(qrCodeUrl)
                LblPaymentStatusConsult.Text = paymentStatus
                LblUpdateAtConsult.Text = updateAt.AddHours(-3)
                If LblPaymentStatusConsult.Text = "Waiting" Then
                    LblPaymentStatusConsult.ForeColor = Color.Gold
                ElseIf LblPaymentStatusConsult.Text = "Paid" Then
                    LblPaymentStatusConsult.ForeColor = Color.Green
                Else
                    LblPaymentStatusConsult.ForeColor = Color.Red
                End If
            End If

        End If

    End Sub

    Private Sub BtnCancelSale_Click(sender As Object, e As EventArgs) Handles BtnCancelSale.Click
        If CancelPayment(TxtPaymentIdConsult.Text) = True Then
            MsgBox("Venda de número: " & TxtPaymentIdConsult.Text & " cancelada com sucesso", vbInformation, "PenseAPI")
            ConsultTransactionStatus(TxtPaymentIdConsult.Text, "paymentId")
        Else
            MsgBox("Falha ao cancelar venda!" & vbCrLf & vbCrLf & returnWithError, vbCritical, "PenseAPI")
            Exit Sub
        End If
    End Sub

    Private Sub BtnConsultByReference_Click(sender As Object, e As EventArgs) Handles BtnConsultByReference.Click
        ConsultTransactionStatus(TxtExternalReferencePaymentConsult.Text, "externalReference")
    End Sub
End Class
