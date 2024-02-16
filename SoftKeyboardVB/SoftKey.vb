Option Strict On
Imports System.Runtime.InteropServices

'--------------------------------------------------------------------------------
' ソフトキーボード画面クラス
'--------------------------------------------------------------------------------
Public Class SoftKey
    '--------------------------------------------------------
    ' メンバ変数
    '--------------------------------------------------------
    Private MousePoint As Point                     ' マウスのクリック位置を記憶 （画面ドラッグ移動用）
    Private FormLoaded As Boolean = False           ' フォームロード済みフラグ （初期位置サイズ確定前動作抑制用）
    Private AllControls As Control()                ' 全コントロール
    Private Buttons(0 To 64) As ButtonEx            ' 五十音ボタン配列 13×5＝65個
    Private OrgWidth As Integer                     ' 元幅

    ' メイン関数
    <STAThread()> Shared Sub Main()
        ' 二重起動をチェックする
        If Diagnostics.Process.GetProcessesByName(
            Diagnostics.Process.GetCurrentProcess.ProcessName).Length > 1 Then
            Return
        End If
        Application.EnableVisualStyles()
        Application.SetCompatibleTextRenderingDefault(False)
        Application.Run(New SoftKey())
    End Sub

    ' すべてのコントロールを再帰的に取得する。
    Private Function GetAllControls(ByVal top As Control) As Control()
        Dim buf As ArrayList = New ArrayList
        For Each c As Control In top.Controls
            buf.Add(c)
            buf.AddRange(GetAllControls(c))
        Next
        Return CType(buf.ToArray(GetType(Control)), Control())
    End Function

    ' フォームがアクティブにならないようにする。
    Protected Overrides ReadOnly Property CreateParams() As CreateParams
        Get
            Const WS_EX_NOACTIVATE As Integer = &H8000000
            Dim p As CreateParams = MyBase.CreateParams
            If Not MyBase.DesignMode Then
                p.ExStyle = p.ExStyle Or WS_EX_NOACTIVATE ' 非アクティブ化 拡張スタイル指定
            End If
            Return p
        End Get
    End Property

    '  フォームロード
    Private Sub SoftKeyForm_Load(sender As Object, e As System.EventArgs) Handles Me.Load
        ' すべてのコントロールを再帰的に取得する。
        AllControls = GetAllControls(Me)
        ' 初期位置サイズを確定する
        For Each c As Control In AllControls
            If c.GetType().Equals(GetType(ButtonEx)) Then               ' ボタン
                DirectCast(c, ButtonEx).Cmn.GetOrigins()
                ' 五十音ボタン配列を取得する
                If c.TabIndex < Buttons.Length Then
                    Buttons(c.TabIndex) = DirectCast(c, ButtonEx)
                End If
            ElseIf c.GetType().Equals(GetType(RadioButtonEx)) Then      ' ラジオボタン
                DirectCast(c, RadioButtonEx).Cmn.GetOrigins()
            ElseIf c.GetType().Equals(GetType(LabelEx)) Then            ' ラベル
                DirectCast(c, LabelEx).Cmn.GetOrigins()
            End If
        Next
        ' 元幅取得
        OrgWidth = ClientSize.Width

        ' フォームロード済みフラグON （初期位置サイズ確定前動作抑制用）
        FormLoaded = True

        RadioMode1.Checked = True   ' モード
    End Sub

    ' Windowsメッセージ処理
    ' ・矩形サイズ型構造体 WndProcハンドラ WM_SIZING パラメータ用
    Private Structure RECT
        Public left As Integer
        Public top As Integer
        Public right As Integer
        Public bottom As Integer
    End Structure

    Protected Overrides Sub WndProc(ByRef m As System.Windows.Forms.Message)
        Const WM_SIZING As Integer = &H214
        Const WM_NCHITTEST As Integer = &H84

        If m.Msg = WM_NCHITTEST And m.HWnd.Equals(Me.Handle) Then
            ' 隅リサイズ有効化
            Const HTTOPLEFT As Integer = 13
            Const HTTOPRIGHT As Integer = 14
            Const HTBOTTOMLEFT As Integer = 16
            Const HTBOTTOMRIGHT As Integer = 17
            Dim p As Point = Me.PointToClient(New Point(m.LParam.ToInt32 Mod 65536, m.LParam.ToInt32 \ 65536))
            If p.X < Me.ClientRectangle.Left + 15 Then
                If p.Y < Me.ClientRectangle.Top + 15 Then
                    m.Result = CType(HTTOPLEFT, IntPtr)     ' 左上
                    Exit Sub
                End If
                If p.Y > Me.ClientRectangle.Bottom - 15 Then
                    m.Result = CType(HTBOTTOMLEFT, IntPtr)  ' 左下
                    Exit Sub
                End If
            End If
            If p.X > Me.ClientRectangle.Right - 15 Then
                If p.Y < Me.ClientRectangle.Top + 15 Then
                    m.Result = CType(HTTOPRIGHT, IntPtr)    ' 右上
                    Exit Sub
                End If
                If p.Y > Me.ClientRectangle.Bottom - 15 Then
                    m.Result = CType(HTBOTTOMRIGHT, IntPtr) ' 右下
                    Exit Sub
                End If
            End If

        ElseIf m.Msg = WM_SIZING And m.HWnd.Equals(Me.Handle) Then
            ' 位置・サイズの縦横比率補正と下限・上限補正
            Static aspect_ratio As Double = 0   ' 縦横比率
            Const WMSZ_LEFT As Integer = 1
            Const WMSZ_TOP As Integer = 3
            Const WMSZ_TOPLEFT As Integer = 4
            Const WMSZ_TOPRIGHT As Integer = 5
            Const WMSZ_BOTTOMLEFT As Integer = 7
            ' lParamを矩形サイズ型に変換
            Dim r As RECT = DirectCast(Marshal.PtrToStructure(m.LParam, GetType(RECT)), RECT)
            ' 現在の幅と高さを取得
            Dim width As Double = r.right - r.left
            Dim height As Double = r.bottom - r.top
            ' サイズなし状態スキップ
            If width = 0 Or height = 0 Then
                MyBase.WndProc(m)   ' 標準処理を実行する
                Exit Sub
            End If

            ' 初回に縦横比率を計算保持
            If aspect_ratio = 0 Then
                aspect_ratio = height / width
            End If

            ' 縦横比率による幅と高さを計算
            If height / width > aspect_ratio Then
                width = height / aspect_ratio   ' 高く細い場合、幅を補正
            Else
                height = width * aspect_ratio   ' 低く広い場合、高さを補正
            End If

            ' 下限・上限補正
            If width < 400 Then
                width = 400
                height = width * aspect_ratio
            ElseIf width > 950 Then
                width = 950
                height = width * aspect_ratio
            End If

            ' 横リサイズによる左右位置補正
            If m.WParam.ToInt32 = WMSZ_LEFT Or m.WParam.ToInt32 = WMSZ_TOPLEFT Or m.WParam.ToInt32 = WMSZ_BOTTOMLEFT Then
                r.left = r.right - CInt(width)  ' 左位置補正
            Else
                r.right = r.left + CInt(width)  ' 右位置補正
            End If

            ' 縦リサイズによる上下位置補正
            If m.WParam.ToInt32 = WMSZ_TOP Or m.WParam.ToInt32 = WMSZ_TOPLEFT Or m.WParam.ToInt32 = WMSZ_TOPRIGHT Then
                r.top = r.bottom - CInt(height) ' 上位置補正
            Else
                r.bottom = r.top + CInt(height) ' 下位置補正
            End If

            ' メッセージのLParamを更新する
            Marshal.StructureToPtr(r, m.LParam, True)
        End If

        MyBase.WndProc(m)   ' 標準処理を実行する
    End Sub

    ' リサイズ
    Private Sub SoftKeyForm_Resize(sender As Object, e As System.EventArgs) Handles Me.Resize
        ' 初期位置サイズ確定前スキップ（フォームロード前）
        If FormLoaded = False Then
            Exit Sub
        End If
        Dim ratio As Double = ClientSize.Width / OrgWidth
        ' コントロールを対ClientSize比率でサイズ変更、フォント変更する
        For Each c As Control In AllControls
            If c.GetType().Equals(GetType(ButtonEx)) Then               ' ボタン
                DirectCast(c, ButtonEx).Cmn.ResizeByOrigins(ratio)
            ElseIf c.GetType().Equals(GetType(RadioButtonEx)) Then      ' ラジオボタン
                DirectCast(c, RadioButtonEx).Cmn.ResizeByOrigins(ratio)
            ElseIf c.GetType().Equals(GetType(LabelEx)) Then            ' ラベル
                DirectCast(c, LabelEx).Cmn.ResizeByOrigins(ratio)
            End If
        Next
    End Sub

    ' マウスダウンイベントハンドラ 画面のマウスドラッグ移動用
    Private Sub SoftKeyForm_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles LabelTop.MouseDown, MyBase.MouseDown
        'マウスのボタンが押されたとき
        If (e.Button And MouseButtons.Left) = MouseButtons.Left Then
            '位置を記憶する
            MousePoint = New Point(e.X, e.Y)
        End If
    End Sub

    ' マウス移動イベントハンドラ 画面のマウスドラッグ移動用
    Private Sub SoftKeyForm_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles LabelTop.MouseMove, MyBase.MouseMove
        'マウスが動いたとき
        If (e.Button And MouseButtons.Left) = MouseButtons.Left Then
            Location = New Point(Location.X + e.X - MousePoint.X, Location.Y + e.Y - MousePoint.Y)
        End If
    End Sub

    ' モード選択変更
    Private Sub RadioMode_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles RadioMode1.CheckedChanged, RadioMode2.CheckedChanged, RadioMode3.CheckedChanged, RadioButtonEx1.CheckedChanged
        Static fonts() As String = {"メイリオ", "メイリオ", "ＭＳ ゴシック", "メイリオ"}
        Static words() As String = {
        "あいうえおかきくけこさしすせそたちつてとなにぬねのはひふへほまみむめもやゆよ  らりるれろわをん  ぁぃぅぇぉゃゅょっ ー゛゜、。",
        "アイウエオカキクケコサシスセソタチツテトナニヌネノハヒフヘホマミムメモヤユヨ  ラリルレロワヲン  ァィゥェォャュョッ ー゛゜、。",
        "ｱｲｳｴｵｶｷｸｹｺｻｼｽｾｿﾀﾁﾂﾃﾄﾅﾆﾇﾈﾉﾊﾋﾌﾍﾎﾏﾐﾑﾒﾓﾔﾕﾖ  ﾗﾘﾙﾚﾛﾜｦﾝ  ｧｨｩｪｫｬｭｮｯ ｰﾞﾟ､｡",
        " mzMZ lyLY kxKX0jwJW9ivIV8huHU7gtGT6fsFS5erER4dqDQ3cpCP2boBO1anAN"
        }
        Dim r As RadioButtonEx = DirectCast(sender, RadioButtonEx)
        ' 選択解除をスキップ
        If r.Checked = False Then
            Exit Sub
        End If

        Dim mode As Integer = CInt(r.Tag)
        Dim ratio As Double = ClientSize.Width / OrgWidth
        For Each b As ButtonEx In Buttons
            ' 文字キーを取得する
            If b.TabIndex < words(mode).Length Then
                b.Cmn.Text = words(mode)(b.TabIndex)
            Else
                b.Cmn.Text = " "
            End If
            ' 文字キーなし  ボタン無効化
            b.Visible = CBool(IIf(b.Cmn.Text = " ", False, True))
            ' コントロールを対ClientSize比率でサイズ変更、フォント変更する
            b.Cmn.ResizeByOrigins(ratio, fonts(mode))
            b.Invalidate()
        Next
    End Sub

    ' 閉じるクリック
    Private Sub ButtonClose_Click(sender As System.Object, e As System.EventArgs) Handles ButtonClose.Click
        Me.Close()
    End Sub

    ' 文字キーコードクリック （ボタンテキスト）
    Private Sub ButtonWordKeyCode_Click(sender As System.Object, e As System.EventArgs) Handles Button64.Click, Button63.Click, Button62.Click, Button61.Click, Button60.Click, Button59.Click, Button58.Click, Button57.Click, Button56.Click, Button55.Click, Button54.Click, Button53.Click, Button52.Click, Button51.Click, Button50.Click, Button49.Click, Button48.Click, Button47.Click, Button46.Click, Button45.Click, Button44.Click, Button43.Click, Button42.Click, Button41.Click, Button40.Click, Button39.Click, Button38.Click, Button37.Click, Button36.Click, Button35.Click, Button34.Click, Button33.Click, Button32.Click, Button31.Click, Button30.Click, Button29.Click, Button28.Click, Button27.Click, Button26.Click, Button25.Click, Button24.Click, Button23.Click, Button22.Click, Button21.Click, Button20.Click, Button19.Click, Button18.Click, Button17.Click, Button16.Click, Button15.Click, Button14.Click, Button13.Click, Button12.Click, Button11.Click, Button10.Click, Button09.Click, Button08.Click, Button07.Click, Button06.Click, Button05.Click, Button04.Click, Button03.Click, Button02.Click, Button01.Click, Button00.Click
        ' UNICODE文字キーボード入力送信
        NativeMethods.SendInputKeybordUNICODE(CUShort(AscW(DirectCast(sender, ButtonEx).Cmn.Text(0))))
    End Sub

    ' 拡張キーコードソフトキークリック
    Private Sub ButtonExtendKeyCode_Click(sender As System.Object, e As System.EventArgs) Handles ButtonUp.Click, ButtonRight.Click, ButtonLeft.Click, ButtonEnter.Click, ButtonDown.Click, ButtonDel.Click, ButtonBS.Click, ButtonSpace.Click
        ' 拡張キーボード入力送信
        NativeMethods.SendInputKeybordExtend(UShort.Parse(DirectCast(sender, ButtonEx).Tag.ToString, System.Globalization.NumberStyles.HexNumber))
    End Sub
End Class

'--------------------------------------------------------
' NativeMethodsクラス ※DllImport属性を用いて定義したメソッドは、NativeMethodsというクラスに含める規則
'--------------------------------------------------------
Friend NotInheritable Class NativeMethods
    ' ・SendInput 関数でキー操作に関する動作等を指定する KEYBDINPUT 構造体
    <StructLayout(LayoutKind.Sequential)>
    Private Structure KEYBDINPUT
        Public wVk As UShort          ' 仮想キーコードのコード254の範囲内の値を設定
        Public wScan As UShort        ' ハードウェアキーのスキャンコードを設定
        Public dwFlags As Integer     ' キーボードの動作を指定するフラグを設定
        Public time As Integer        ' タイムスタンプ(このメンバは無視されます)
        Public dwExtraInfo As IntPtr  ' 追加情報(このメンバは無視されます)
    End Structure

    ' ・SendInput 関数の設定に使用する INPUT 構造体 （注）Size:=40が必須
    '   type    SendInput 関数の使用目的 INPUT_KEYBOARDのみ使用
    '   ki      KEYBDINPUT 構造体
    <StructLayout(LayoutKind.Sequential, Size:=40)>
    Private Structure INPUT
        Public type As Integer
        Public ki As KEYBDINPUT
    End Structure

    Private Sub New()
    End Sub

    '--------------------------------------------------------------------------------
    '- 外部モジュールの宣言
    '--------------------------------------------------------------------------------
    ' ・キーストローク、マウスの動き、ボタンのクリックなどを合成します。
    '   nInputs     入力イベントの数
    '   pInputs()   挿入する入力イベントの配列
    '   cbsize      構造体のサイズ
    '    戻り値     挿入することができたイベントの数を返す。
    '               ブロックされている場合は 0 を返す
    <DllImport("user32.dll", CharSet:=CharSet.Auto)>
    Private Shared Function SendInput(ByVal nInputs As Integer, ByVal pInputs() As INPUT, ByVal cbsize As Integer) As Integer
    End Function

    ' ・仮想キーコード・ASCII値・スキャンコード間でコードを変換する
    '   wCode       仮想キーコードまたはスキャンコード
    '   wMapType    実行したい変換の種類
    '   戻り値      スキャンコード、仮想キーコード、ASCII 値のいずれかが返ります。
    '               変換されないときは、0 が返ります。
    <DllImport("user32.dll", CharSet:=CharSet.Auto)>
    Private Shared Function MapVirtualKey(ByVal wCode As Integer, ByVal wMapType As Integer) As Integer
    End Function

    ' キーボード入力送信
    Private Shared Sub SendInputKeybord(wVk As UShort, wScan As UShort, dwFlags As Integer)
        Const INPUT_KEYBOARD As Integer = 1
        Const KEYEVENTF_KEYUP As Integer = &H2
        Try ' Win32APIの例外エラーを無視して強制終了を防ぐ
            Dim inp(2) As INPUT
            inp(0).type = INPUT_KEYBOARD    ' キーボード入力
            inp(0).ki.wVk = wVk             ' 仮想キーコード
            inp(0).ki.wScan = wScan         ' スキャンコード
            inp(0).ki.time = 0              ' タイムスタンプ
            inp(0).ki.dwFlags = dwFlags     ' 動作フラグ キーダウン = 0
            inp(0).ki.dwExtraInfo = IntPtr.Zero ' 追加情報
            inp(1) = inp(0) ' KEYDOWN→KEYUPコピー
            inp(1).ki.dwFlags = dwFlags Or KEYEVENTF_KEYUP  '  動作フラグ キーアップ
            ' 入力送信
            SendInput(2, inp, Marshal.SizeOf(GetType(INPUT)))
        Catch ex As System.Exception
        End Try
    End Sub

    ' UNICODE文字キーボード入力送信
    Friend Shared Sub SendInputKeybordUNICODE(wVk As UShort)
        Const KEYEVENTF_UNICODE As Integer = &H4
        SendInputKeybord(0, wVk, KEYEVENTF_UNICODE)
    End Sub

    ' 拡張キーボード入力送信
    Friend Shared Sub SendInputKeybordExtend(wVk As UShort)
        Const KEYEVENTF_EXTENDEDKEY = &H1
        SendInputKeybord(wVk, CUShort(NativeMethods.MapVirtualKey(wVk, 0)), KEYEVENTF_EXTENDEDKEY)
    End Sub

    ' 拡張キーボード入力送信 スキャンコード指定
    Friend Shared Sub SendInputKeybordExtend(wVk As UShort, wScan As UShort)
        Const KEYEVENTF_EXTENDEDKEY = &H1
        SendInputKeybord(wVk, wScan, KEYEVENTF_EXTENDEDKEY)
    End Sub

End Class

'--------------------------------------------------------
' 共通機能クラス
'--------------------------------------------------------
Class CommonFunc
    ' ・元情報構造体
    Private Structure ORIGINS
        Public Width As Integer       ' 横幅
        Public Height As Integer      ' 縦幅
        Public X As Integer           ' x座標
        Public Y As Integer           ' y座標
        Public FontSize As Single     ' フォントサイズ
    End Structure
    '--------------------------------------------------------
    ' メンバ変数
    '--------------------------------------------------------
    Public Text As String               ' 描画文字列
    Public StringFmt As StringFormat    ' 文字書式

    Private Shared BaseStringFmt As StringFormat ' 基本文字書式
    Private Ctrl As Control             ' 自コントロール
    Private Orgs As ORIGINS             ' 元情報
    Private FontName As String          ' フォント名

    ' 共通機能クラス生成
    Sub New(c As Control)
        MyBase.New()
        Ctrl = c
        ' 基本文字書式
        If StringFmt Is Nothing Then
            BaseStringFmt = New StringFormat()
            BaseStringFmt.Alignment = StringAlignment.Center
            BaseStringFmt.LineAlignment = StringAlignment.Center
        End If
        StringFmt = BaseStringFmt
        Orgs.FontSize = 1   ' ※初期化しないとPaintの処理でデザイナーがエラーとなる
    End Sub

    ' コントロールの元情報を取得する
    Public Sub GetOrigins()
        ' ラベル以外は自前描画の為、文字列退避
        If Ctrl.GetType().Equals(GetType(LabelEx)) = False Then
            Text = Ctrl.Text
            Ctrl.Text = ""
        End If
        ' 初期フォント保持
        FontName = Ctrl.Font.Name
        ' 元情報を取得
        Orgs.Width = Ctrl.Width
        Orgs.Height = Ctrl.Height
        Orgs.X = Ctrl.Location.X
        Orgs.Y = Ctrl.Location.Y
        Orgs.FontSize = Ctrl.Font.Size
    End Sub

    ' コントロールを元情報でサイズ変更、フォント変更する
    Public Sub ResizeByOrigins(ratio As Double, Optional fname As String = "")
        Try
            Dim pos As Point
            Ctrl.Width = CInt(Orgs.Width * ratio)
            Ctrl.Height = CInt(Orgs.Height * ratio)
            pos.X = CInt(Orgs.X * ratio)
            pos.Y = CInt(Orgs.Y * ratio)
            Ctrl.Location = pos
            If fname <> "" Then
                FontName = fname
            End If
            Dim fontsize As Single = CSng(Orgs.FontSize * ratio)
            Ctrl.Font = New Font(FontName, fontsize, Ctrl.Font.Style)
        Finally
        End Try
    End Sub

    ' 描画処理 ※ボタン内の余白が大きいし拡縮でずれる！ → しかたないので自前で追加描画する
    Public Sub Paint(g As Graphics)
        Try
            Dim brush As Brush = New SolidBrush(Ctrl.ForeColor)
            Dim rect As Rectangle = Ctrl.ClientRectangle
            rect.Y = CInt(3 * Ctrl.Font.Size / Orgs.FontSize)  ' 縦余白補正
            g.TextRenderingHint = Drawing.Text.TextRenderingHint.SystemDefault
            g.DrawString(Text, Ctrl.Font, brush, rect, StringFmt)
            brush.Dispose()
        Finally
        End Try
    End Sub
End Class

'--------------------------------------------------------
' ラベルクラス
'--------------------------------------------------------
Class LabelEx
    Inherits System.Windows.Forms.Label
    '--------------------------------------------------------
    ' メンバ変数
    '--------------------------------------------------------
    Public Cmn As CommonFunc    ' 共通機能クラス
    Private Shared BaseStringFmt As StringFormat    ' 基本文字書式
    ' ラベル作成
    Sub New()
        MyBase.New()
        Cmn = New CommonFunc(Me)
        ' 基本文字書式
        If BaseStringFmt Is Nothing Then
            BaseStringFmt = New StringFormat()
            BaseStringFmt.Alignment = StringAlignment.Near
        End If
        Cmn.StringFmt = BaseStringFmt
    End Sub
End Class

'--------------------------------------------------------
' ラジオボタンクラス
'--------------------------------------------------------
Class RadioButtonEx
    Inherits System.Windows.Forms.RadioButton
    '--------------------------------------------------------
    ' メンバ変数
    '--------------------------------------------------------
    Public Cmn As CommonFunc    ' 共通機能クラス
    ' ラジオボタン作成
    Sub New()
        MyBase.New()
        Me.SetStyle(ControlStyles.Selectable, False)    ' フォーカス回避
        Cmn = New CommonFunc(Me)
    End Sub
    ' 描画処理
    Protected Overrides Sub OnPaint(e As PaintEventArgs)
        MyBase.OnPaint(e)
        Cmn.Paint(e.Graphics)
    End Sub
End Class

'--------------------------------------------------------
' ボタンクラス
'--------------------------------------------------------
Class ButtonEx
    Inherits System.Windows.Forms.Button
    '--------------------------------------------------------
    ' メンバ変数
    '--------------------------------------------------------
    Public Cmn As CommonFunc    ' 共通機能クラス
    ' ボタン作成
    Sub New()
        MyBase.New()
        Me.SetStyle(ControlStyles.Selectable, False)    ' フォーカス回避
        Cmn = New CommonFunc(Me)
    End Sub
    ' 描画処理
    Protected Overrides Sub OnPaint(e As PaintEventArgs)
        MyBase.OnPaint(e)
        Cmn.Paint(e.Graphics)
    End Sub
End Class
