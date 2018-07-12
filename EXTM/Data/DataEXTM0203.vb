Imports FarPoint.Win.Spread

Public Class DataEXTM0203

    '画面情報
    Private ppTxtExasCode As TextBox '検索条件：EXAS相手先コードテキストボックス
    Private ppTxtExasName As TextBox '検索条件：EXAS相手先名テキストボックス
    Private ppVwResult As FpSpread '検索結果表示スプレッドシート

    '検索条件
    Private ppStrExasCode_Search As String '検索条件：EXAS相手先コード
    Private ppStrExasName_Search As String '検索条件：EXAS相手先名

    '検索結果
    Private ppDtResult As DataTable 'EXAS相手先検索結果データテーブル
    Private ppIntCheckRow As Integer 'チェックされた行インデックス
    Private ppStrMenuID As String '遷移前の画面ID


    ''' <summary>
    ''' プロパティセット【検索条件：EXAS相手先コードテキストボックス】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppTxtExasCode</returns>
    ''' <remarks><para>作成情報：2015/08/26 d.sonoda
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropTxtExasCode() As TextBox
        Get
            Return ppTxtExasCode
        End Get
        Set(ByVal value As TextBox)
            ppTxtExasCode = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【検索条件：EXAS相手先名テキストボックス】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppTxtExasName</returns>
    ''' <remarks><para>作成情報：2015/08/26 d.sonoda
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropTxtExasName() As TextBox
        Get
            Return ppTxtExasName
        End Get
        Set(ByVal value As TextBox)
            ppTxtExasName = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【検索結果表示スプレッドシート】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppVwResult</returns>
    ''' <remarks><para>作成情報：2015/08/26 d.sonoda
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropVwResult() As FpSpread
        Get
            Return ppVwResult
        End Get
        Set(ByVal value As FpSpread)
            ppVwResult = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【検索条件：EXAS相手先コード】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrExasCode_Search</returns>
    ''' <remarks><para>作成情報：2015/08/26 d.sonoda
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropStrExasCode_Search() As String
        Get
            Return ppStrExasCode_Search
        End Get
        Set(ByVal value As String)
            ppStrExasCode_Search = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【検索条件：EXAS相手先名】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrExasName_Search</returns>
    ''' <remarks><para>作成情報：2015/08/26 d.sonoda
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropStrExasName_Search() As String
        Get
            Return ppStrExasName_Search
        End Get
        Set(ByVal value As String)
            ppStrExasName_Search = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【EXAS相手先検索結果データテーブル】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppDtResult</returns>
    ''' <remarks><para>作成情報：2015/08/26 d.sonoda
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropDtResult() As DataTable
        Get
            Return ppDtResult
        End Get
        Set(ByVal value As DataTable)
            ppDtResult = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【チェックされた行インデックス】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppIntCheckRow</returns>
    ''' <remarks><para>作成情報：2015/08/26 d.sonoda
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropIntCheckRow() As Integer
        Get
            Return ppIntCheckRow
        End Get
        Set(ByVal value As Integer)
            ppIntCheckRow = value
        End Set
    End Property

    ''' <summary>
    ''' プロパティセット【遷移前の画面ID】
    ''' </summary>
    ''' <value></value>
    ''' <returns>ppStrMenuID</returns>
    ''' <remarks><para>作成情報：2015/08/26 d.sonoda
    ''' <p>改訂情報:</p>
    ''' </para></remarks>
    Public Property PropStrMenuID() As String
        Get
            Return ppStrMenuID
        End Get
        Set(ByVal value As String)
            ppStrMenuID = value
        End Set
    End Property

End Class
