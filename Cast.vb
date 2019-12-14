Imports System.Globalization

Public Class Cast
    Public Function stringDoDouble(ByVal str As String) As Double
        str = str.Replace(",", ".")
        Dim value As Double
        Double.TryParse(str, NumberStyles.Any, CultureInfo.InvariantCulture, value)
        '(10.6 * 18000 / 1000) * (1.125 / 1000)
        Return value
    End Function
End Class
