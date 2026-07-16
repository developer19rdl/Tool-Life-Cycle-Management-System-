Imports System.Windows.Data
Imports System.Globalization
Imports System.Windows.Media.Imaging
Imports System.IO

Public Class Base64ToImageConverter
    Implements IValueConverter

    Public Function Convert(value As Object, targetType As Type, parameter As Object, culture As CultureInfo) As Object Implements IValueConverter.Convert
        If value Is Nothing OrElse String.IsNullOrEmpty(value.ToString()) Then
            Return Nothing
        End If

        Try
            Dim base64String As String = value.ToString()
            Dim binaryData() As Byte = System.Convert.FromBase64String(base64String)
            Dim bi As New BitmapImage()
            bi.BeginInit()
            bi.StreamSource = New MemoryStream(binaryData)
            bi.CacheOption = BitmapCacheOption.OnLoad
            bi.EndInit()
            Return bi
        Catch ex As Exception
            Return Nothing
        End Try
    End Function

    Public Function ConvertBack(value As Object, targetType As Type, parameter As Object, culture As CultureInfo) As Object Implements IValueConverter.ConvertBack
        Throw New NotImplementedException()
    End Function
End Class

Public Class Base64ToVisibilityConverter
    Implements IValueConverter

    Public Function Convert(value As Object, targetType As Type, parameter As Object, culture As CultureInfo) As Object Implements IValueConverter.Convert
        Dim hasValue = Not (value Is Nothing OrElse String.IsNullOrEmpty(value.ToString()))
        
        If parameter IsNot Nothing AndAlso parameter.ToString() = "Inverse" Then
            Return If(hasValue, System.Windows.Visibility.Collapsed, System.Windows.Visibility.Visible)
        Else
            Return If(hasValue, System.Windows.Visibility.Visible, System.Windows.Visibility.Collapsed)
        End If
    End Function

    Public Function ConvertBack(value As Object, targetType As Type, parameter As Object, culture As CultureInfo) As Object Implements IValueConverter.ConvertBack
        Throw New NotImplementedException()
    End Function
End Class
