Imports FarPoint.Win.Spread
Public Class DataEXTM0103
    'フォームオブジェクト
    Private ppVwList As FpSpread                    '一覧シート
    'フッタ
    Private ppBtnReg As Button                      'フッタ：登録ボタン
    Private ppBtnBack As Button                     'フッタ：戻るボタン
    'データ
    Private ppDtExtTaxMasta As DataTable            '検索結果を格納するデータテーブル

    ''' <summary>
    ''' プロパティセット【一覧シート】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppVwList</returns>
    ''' <remarks><para>作成情報：2015/08/26 hayabuchi
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropVwList() As FpSpread
        Get
            Return ppVwList
        End Get
        Set(ByVal value As FpSpread)
            ppVwList = value
        End Set
    End Property
    ''' <summary>
    ''' プロパティセット【フッタ：登録ボタン】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppBtnReg</returns>
    ''' <remarks><para>作成情報：2015/08/26 hayabuchi
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropBtnReg() As Button
        Get
            Return ppBtnReg
        End Get
        Set(ByVal value As Button)
            ppBtnReg = value
        End Set
    End Property
    ''' <summary>
    ''' プロパティセット【フッタ：戻るボタン】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppBtnBack</returns>
    ''' <remarks><para>作成情報：2015/08/26 hayabuchi
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropBtnBack() As Button
        Get
            Return ppBtnBack
        End Get
        Set(ByVal value As Button)
            ppBtnBack = value
        End Set
    End Property
    ''' <summary>
    ''' プロパティセット【データテーブル】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppDtExtTaxMasta</returns>
    ''' <remarks><para>作成情報：2015/08/26 hayabuchi
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropDtExtTaxMasta() As DataTable
        Get
            Return ppDtExtTaxMasta
        End Get
        Set(ByVal value As DataTable)
            ppDtExtTaxMasta = value
        End Set
    End Property
End Class
