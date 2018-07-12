Imports Common
Imports CommonEXT

Public Class CommonDataRiyobiCompareter

    Implements System.Collections.IComparer

    Public Function Compare(ByVal x As Object, ByVal y As Object) As Integer _
        Implements System.Collections.IComparer.Compare
        'String型以外の比較はエラー
        If Not (TypeOf x Is CommonDataRiyobi) Then
            Return -1
        ElseIf Not (TypeOf y Is CommonDataRiyobi) Then
            Return 1
        End If

        Dim dataRiyobi1 As CommonDataRiyobi = x
        Dim dataRiyobi2 As CommonDataRiyobi = y
        If dataRiyobi1.PropStrYoyakuDt Is Nothing Then
            Return -1
        End If
        If dataRiyobi2.PropStrYoyakuDt Is Nothing Then
            Return 1
        End If
        Dim str1 As String = dataRiyobi1.PropStrYoyakuDt
        Dim str2 As String = dataRiyobi2.PropStrYoyakuDt

        Dim dt1 As DateTime = DateTime.Parse(str1)
        Dim dt2 As DateTime = DateTime.Parse(str2)
        Return dt1.CompareTo(dt2)
    End Function
End Class