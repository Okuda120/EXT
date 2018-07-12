Imports Common
Imports CommonEXT
Public Class CommonLogicEXTB

    ''' <summary>
    ''' ステータス名称取得
    ''' </summary>
    ''' <param name="strStatus"></param>
    ''' <returns>ステータス名称</returns>
    ''' <remarks></remarks>
    Public Function GetYoyakuSts(ByVal strStatus As String) As String
        '---2016.1.15 START ↓y.ozawa ステータス名称変更
        If strStatus = YOYAKU_STS_KARI_MI Then
            Return "仮予約"
        ElseIf strStatus = YOYAKU_STS_KARI Then
            Return "決定"
        ElseIf strStatus = YOYAKU_STS_SEISHIKI Then
            Return "申請受諾済"
        ElseIf strStatus = YOYAKU_STS_SEISHIKI_COMP Then
            Return "精算完了"
        ElseIf strStatus = YOYAKU_STS_CANCEL_KARI Then
            Return "仮予約キャンセル"
        ElseIf strStatus = YOYAKU_STS_CANCEL_SEISHIKI Then
            Return "正式予約キャンセル"
        End If
        Return ""
        '---2016.1.15 END ↑y.ozawa ステータス名称変更
    End Function

    ''' <summary>
    ''' ステータスCOLOR取得
    ''' </summary>
    ''' <param name="strStatus"></param>
    ''' <returns>ステータスCOLOR</returns>
    ''' <remarks></remarks>
    Public Function GetYoyakuStsColor(ByVal strStatus As String) As Drawing.Color
        '---2016.1.15 START ↓y.ozawa ステータスCOLOR変更
        If strStatus = YOYAKU_STS_KARI_MI Then
            Return Color.White
        ElseIf strStatus = YOYAKU_STS_KARI Then
            Return Color.FromArgb(127, 255, 255)
        ElseIf strStatus = YOYAKU_STS_SEISHIKI Then
            Return Color.FromArgb(191, 255, 127)
        ElseIf strStatus = YOYAKU_STS_SEISHIKI_COMP Then
            Return Color.FromArgb(255, 255, 127)
        End If
        Return Color.FromArgb(224, 224, 224)
        '---2016.1.15 END ↑y.ozawa ステータスCOLOR変更
    End Function

    ''' <summary>
    ''' ステータス名称取得
    ''' </summary>
    ''' <param name="strStatus"></param>
    ''' <returns>ステータス名称</returns>
    ''' <remarks></remarks>
    Public Function GetCancelSts(ByVal strStatus As String) As String
        If strStatus = CANCEL_STS_MATI Then
            Return "キャンセル待ち"
        ElseIf strStatus = CANCEL_STS_KARI Then
            Return "（仮）予約昇格"
        ElseIf strStatus = CANCEL_STS_TORIKESI Then
            Return "キャンセル取消"
        End If
        Return ""
    End Function

End Class
