Namespace PenseAPI

    Public Class ReturnAuthentication
        Public Property authenticated As Boolean
        Public Property message As String
        Public Property create As DateTime
        Public Property expiration As DateTime
        Public Property accessToken As String
        Public Property accessKey As String
        Public Property clientId As String
    End Class

    Public Class ReturnPayment
        Public Property id As Integer
        Public Property description As String
        Public Property externalReference As String
        Public Property urlCallback As String
        Public Property qrCodeUrl As String
        Public Property qrCodeData As String
        Public Property status As String
        Public Property value As Double
        Public Property wallet As String
        Public Property updateAt As DateTime
    End Class

    Public Class ReturnPaymentStatus
        Public Property id As Integer
        Public Property description As String
        Public Property externalReference As String
        Public Property urlCallback As String
        Public Property qrCodeUrl As String
        Public Property qrCodeData As String
        Public Property status As String
        Public Property value As Double
        Public Property wallet As String
        Public Property updateAt As DateTime
    End Class

End Namespace