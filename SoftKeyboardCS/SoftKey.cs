using System;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Collections;
using System.Collections.ObjectModel;

namespace SoftKeyboard
{
    public partial class SoftKey : Form
    {
        public SoftKey()
        {
            InitializeComponent();
        }
        //--------------------------------------------------------
        // メンバ変数
        //--------------------------------------------------------
        // マウスのクリック位置を記憶 （画面ドラッグ移動用）
        private Point MousePoint;
        // フォームロード済みフラグ （初期位置サイズ確定前動作抑制用）
        private bool FormLoaded = false;
        // 全コントロール
        private Control[] AllControls;
        // 五十音ボタン配列 13×5＝65個
        private ButtonEx[] Buttons = new ButtonEx[65];
        // 元幅
        private int OrgWidth;

        // メイン関数
        [STAThread()]
        public static void Main()
        {
            // 二重起動をチェックする
            if (System.Diagnostics.Process.GetProcessesByName(System.Diagnostics.Process.GetCurrentProcess().ProcessName).Length > 1)
            {
                return;
            }
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new SoftKey());
        }

        // すべてのコントロールを再帰的に取得する。
        private Control[] GetAllControls(Control top)
        {
            ArrayList buf = new ArrayList();
            foreach (Control c in top.Controls)
            {
                buf.Add(c);
                buf.AddRange(GetAllControls(c));
            }
            return (Control[])buf.ToArray(typeof(Control));
        }

        // フォームがアクティブにならないようにする。
        protected override CreateParams CreateParams
        {
            get
            {
                const int WS_EX_NOACTIVATE = 0x8000000;
                CreateParams p = base.CreateParams;
                if (!base.DesignMode)
                {
                    p.ExStyle = p.ExStyle | WS_EX_NOACTIVATE;
                    // 非アクティブ化 拡張スタイル指定
                }
                return p;
            }
        }

        //  フォームロード
        private void SoftKeyForm_Load(object sender, System.EventArgs e)
        {
            // すべてのコントロールを再帰的に取得する。
            AllControls = GetAllControls(this);
            // 初期位置サイズを確定する
            foreach (Control c in AllControls)
            {
                // ボタン
                if (c.GetType().Equals(typeof(ButtonEx)))
                {
                    ((ButtonEx)c).Cmn.GetOrigins();
                    // 五十音ボタン配列を取得する
                    if (c.TabIndex < Buttons.Length)
                    {
                        Buttons[c.TabIndex] = (ButtonEx)c;
                    }
                // ラジオボタン
                }
                else if (c.GetType().Equals(typeof(RadioButtonEx)))
                {
                    ((RadioButtonEx)c).Cmn.GetOrigins();
                // ラベル
                }
                else if (c.GetType().Equals(typeof(LabelEx)))
                {
                    ((LabelEx)c).Cmn.GetOrigins();
                }
            }
            // 元幅取得
            OrgWidth = ClientSize.Width;

            // フォームロード済みフラグON （初期位置サイズ確定前動作抑制用）
            FormLoaded = true;

            // モード
            RadioMode1.Checked = true;
        }

        // Windowsメッセージ処理
        // ・矩形サイズ型構造体 WndProcハンドラ WM_SIZING パラメータ用
        private struct RECT
        {
            public int left;
            public int top;
            public int right;
            public int bottom;
        }
  
        // 位置・サイズの縦横比率補正と下限・上限補正
        private double static_WndProc_aspect_ratio = 0;
        protected override void WndProc(ref System.Windows.Forms.Message m)
        {
            const int WM_SIZING = 0x214;
            const int WM_NCHITTEST = 0x84;
            if (m.Msg == WM_NCHITTEST & m.HWnd.Equals(this.Handle))
            {
                // 隅リサイズ有効化
                const int HTTOPLEFT = 13;
                const int HTTOPRIGHT = 14;
                const int HTBOTTOMLEFT = 16;
                const int HTBOTTOMRIGHT = 17;
                Point p = this.PointToClient(new Point(m.LParam.ToInt32() % 65536, m.LParam.ToInt32() / 65536));
                if (p.X < this.ClientRectangle.Left + 15)
                {
                    if (p.Y < this.ClientRectangle.Top + 15)
                    {
                        m.Result = (IntPtr)HTTOPLEFT;           // 左上
                        return;
                    }
                    if (p.Y > this.ClientRectangle.Bottom - 15)
                    {
                        m.Result = (IntPtr)HTBOTTOMLEFT;        // 左下
                        return;
                    }
                }
                if (p.X > this.ClientRectangle.Right - 15)
                {
                    if (p.Y < this.ClientRectangle.Top + 15)
                    {
                        m.Result = (IntPtr)HTTOPRIGHT;        // 右上
                        return;
                    }
                    if (p.Y > this.ClientRectangle.Bottom - 15)
                    {
                        m.Result = (IntPtr)HTBOTTOMRIGHT;    // 右下
                        return;
                    }
                }
            }
            else if (m.Msg == WM_SIZING & m.HWnd.Equals(this.Handle))
            {
                // 縦横比率
                const int WMSZ_LEFT = 1;
                const int WMSZ_TOP = 3;
                const int WMSZ_TOPLEFT = 4;
                const int WMSZ_TOPRIGHT = 5;
                const int WMSZ_BOTTOMLEFT = 7;
                
                // lParamを矩形サイズ型に変換
                RECT r = (RECT)Marshal.PtrToStructure(m.LParam, typeof(RECT));
                // 現在の幅と高さを取得
                double width = r.right - r.left;
                double height = r.bottom - r.top;
                // サイズなし状態スキップ
                if (width == 0 | height == 0)
                {
                    base.WndProc(ref m);
                    // 標準処理を実行する
                    return;
                }

                // 初回に縦横比率を計算保持
                if (static_WndProc_aspect_ratio == 0)
                {
                    static_WndProc_aspect_ratio = height / width;
                }

                // 縦横比率による幅と高さを計算
                if (height / width > static_WndProc_aspect_ratio)
                {
                    width = height / static_WndProc_aspect_ratio;
                    // 高く細い場合、幅を補正
                }
                else
                {
                    height = width * static_WndProc_aspect_ratio;
                    // 低く広い場合、高さを補正
                }

                // 下限・上限補正
                if (width < 400)
                {
                    width = 400;
                    height = width * static_WndProc_aspect_ratio;
                }
                else if (width > 950)
                {
                    width = 950;
                    height = width * static_WndProc_aspect_ratio;
                }

                // 横リサイズによる左右位置補正
                if (m.WParam.ToInt32() == WMSZ_LEFT | m.WParam.ToInt32() == WMSZ_TOPLEFT | m.WParam.ToInt32() == WMSZ_BOTTOMLEFT)
                {
                    r.left = r.right - Convert.ToInt32(width);
                    // 左位置補正
                }
                else
                {
                    r.right = r.left + Convert.ToInt32(width);
                    // 右位置補正
                }

                // 縦リサイズによる上下位置補正
                if (m.WParam.ToInt32() == WMSZ_TOP | m.WParam.ToInt32() == WMSZ_TOPLEFT | m.WParam.ToInt32() == WMSZ_TOPRIGHT)
                {
                    r.top = r.bottom - Convert.ToInt32(height);
                    // 上位置補正
                }
                else
                {
                    r.bottom = r.top + Convert.ToInt32(height);
                    // 下位置補正
                }

                // メッセージのLParamを更新する
                Marshal.StructureToPtr(r, m.LParam, true);
            }

            base.WndProc(ref m);
            // 標準処理を実行する
        }

        // リサイズ
        private void SoftKeyForm_Resize(object sender, System.EventArgs e)
        {
            // 初期位置サイズ確定前スキップ（フォームロード前）
            if (FormLoaded == false)
            {
                return;
            }
            double ratio = (double)ClientSize.Width / OrgWidth;
            // コントロールを対ClientSize比率でサイズ変更、フォント変更する
            foreach (Control c in AllControls)
            {
                // ボタン
                if (c.GetType().Equals(typeof(ButtonEx)))
                {
                    ((ButtonEx)c).Cmn.ResizeByOrigins(ratio);
                    // ラジオボタン
                }
                else if (c.GetType().Equals(typeof(RadioButtonEx)))
                {
                    ((RadioButtonEx)c).Cmn.ResizeByOrigins(ratio);
                    // ラベル
                }
                else if (c.GetType().Equals(typeof(LabelEx)))
                {
                    ((LabelEx)c).Cmn.ResizeByOrigins(ratio);
                }
            }
        }

        // マウスダウンイベントハンドラ 画面のマウスドラッグ移動用
        private void SoftKeyForm_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
        {
            //マウスのボタンが押されたとき
            if ((e.Button & MouseButtons.Left) == MouseButtons.Left)
            {
                //位置を記憶する
                MousePoint = new Point(e.X, e.Y);
            }
        }

        // マウス移動イベントハンドラ 画面のマウスドラッグ移動用
        private void SoftKeyForm_MouseMove(object sender, System.Windows.Forms.MouseEventArgs e)
        {
            //マウスが動いたとき
            if ((e.Button & MouseButtons.Left) == MouseButtons.Left)
            {
                Location = new Point(Location.X + e.X - MousePoint.X, Location.Y + e.Y - MousePoint.Y);
            }
        }

        // モード選択変更
        public static readonly ReadOnlyCollection<string> static_RadioMode_CheckedChanged_fonts =
            Array.AsReadOnly(new string[] { "メイリオ", "ＭＳ ゴシック", "メイリオ", "メイリオ"});
        public static readonly ReadOnlyCollection<string> static_RadioMode_CheckedChanged_words =
            Array.AsReadOnly(new string[] {
                "あいうえおかきくけこさしすせそたちつてとなにぬねのはひふへほまみむめもやゆよ  らりるれろわをん  ぁぃぅぇぉゃゅょっ ー゛゜、。",
                "アイウエオカキクケコサシスセソタチツテトナニヌネノハヒフヘホマミムメモヤユヨ  ラリルレロワヲン  ァィゥェォャュョッ ー゛゜、。",
                "ｱｲｳｴｵｶｷｸｹｺｻｼｽｾｿﾀﾁﾂﾃﾄﾅﾆﾇﾈﾉﾊﾋﾌﾍﾎﾏﾐﾑﾒﾓﾔﾕﾖ  ﾗﾘﾙﾚﾛﾜｦﾝ  ｧｨｩｪｫｬｭｮｯ ｰﾞﾟ､｡",
                " mzMZ lyLY kxKX0jwJW9ivIV8huHU7gtGT6fsFS5erER4dqDQ3cpCP2boBO1anAN"
            });
        private void RadioMode_CheckedChanged(System.Object sender, System.EventArgs e)
        {
            RadioButtonEx r = (RadioButtonEx)sender;
            // 選択解除をスキップ
            if (r.Checked == false)
            {
                return;
            }

            int mode = Convert.ToInt32(r.Tag);
            double ratio = (double)ClientSize.Width / OrgWidth;
            foreach (ButtonEx b in Buttons)
            {
                // 文字キーを取得する
                if (b.TabIndex < static_RadioMode_CheckedChanged_words[mode].Length)
                {
                    b.Cmn.Text = static_RadioMode_CheckedChanged_words[mode][b.TabIndex].ToString();
                }
                else
                {
                    b.Cmn.Text = " ";
                }
                // 文字キーなし  ボタン無効化
                b.Visible = Convert.ToBoolean((b.Cmn.Text == " " ? false : true));
                // コントロールを対ClientSize比率でサイズ変更、フォント変更する
                b.Cmn.ResizeByOrigins(ratio, static_RadioMode_CheckedChanged_fonts[mode]);
                b.Invalidate();
            }
        }

        // 閉じるクリック
        private void ButtonClose_Click(System.Object sender, System.EventArgs e)
        {
            this.Close();
        }

        // 文字キーコードクリック （ボタンテキスト）
        private void ButtonWordKeyCode_Click(System.Object sender, System.EventArgs e)
        {
            // UNICODE文字キーボード入力送信
            NativeMethods.SendInputKeybordUNICODE(Convert.ToUInt16(((ButtonEx)sender).Cmn.Text[0]));
        }

        // 拡張キーコードソフトキークリック
        private void ButtonExtendKeyCode_Click(System.Object sender, System.EventArgs e)
        {
            // 拡張キーボード入力送信
            NativeMethods.SendInputKeybordExtend(ushort.Parse(((ButtonEx)sender).Tag.ToString(), System.Globalization.NumberStyles.HexNumber));
        }
    }
    //--------------------------------------------------------
    // NativeMethodsクラス ※DllImport属性を用いて定義したメソッドは、NativeMethodsというクラスに含める規則
    //--------------------------------------------------------
    internal sealed class NativeMethods
    {
        // ・SendInput 関数でキー操作に関する動作等を指定する KEYBDINPUT 構造体
        [StructLayout(LayoutKind.Sequential)]
        private struct KEYBDINPUT
        {
            // 仮想キーコードのコード254の範囲内の値を設定
            public ushort wVk;
            // ハードウェアキーのスキャンコードを設定
            public ushort wScan;
            // キーボードの動作を指定するフラグを設定
            public int dwFlags;
            // タイムスタンプ(このメンバは無視されます)
            public int time;
            // 追加情報(このメンバは無視されます)
            public IntPtr dwExtraInfo;
        }

        // ・SendInput 関数の設定に使用する INPUT 構造体 （注）Size:=40が必須
        //   type    SendInput 関数の使用目的 INPUT_KEYBOARDのみ使用
        //   ki      KEYBDINPUT 構造体
        [StructLayout(LayoutKind.Sequential, Size = 40)]
        private struct INPUT
        {
            public int type;
            public KEYBDINPUT ki;
        }

        private NativeMethods()
        {
        }

        //--------------------------------------------------------------------------------
        //- 外部モジュールの宣言
        //--------------------------------------------------------------------------------
        // ・キーストローク、マウスの動き、ボタンのクリックなどを合成します。
        //   nInputs     入力イベントの数
        //   pInputs()   挿入する入力イベントの配列
        //   cbsize      構造体のサイズ
        //    戻り値     挿入することができたイベントの数を返す。
        //               ブロックされている場合は 0 を返す
        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        private static extern int SendInput(int nInputs, INPUT[] pInputs, int cbsize);

        // ・仮想キーコード・ASCII値・スキャンコード間でコードを変換する
        //   wCode       仮想キーコードまたはスキャンコード
        //   wMapType    実行したい変換の種類
        //   戻り値      スキャンコード、仮想キーコード、ASCII 値のいずれかが返ります。
        //               変換されないときは、0 が返ります。
        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        private static extern int MapVirtualKey(int wCode, int wMapType);

        // キーボード入力送信
        private static void SendInputKeybord(ushort wVk, ushort wScan, int dwFlags)
        {
            const int INPUT_KEYBOARD = 1;
            const int KEYEVENTF_KEYUP = 0x2;
            // Win32APIの例外エラーを無視して強制終了を防ぐ
            try
            {
                INPUT[] inp = new INPUT[3];
                inp[0].type = INPUT_KEYBOARD;
                // キーボード入力
                inp[0].ki.wVk = wVk;
                // 仮想キーコード
                inp[0].ki.wScan = wScan;
                // スキャンコード
                inp[0].ki.time = 0;
                // タイムスタンプ
                inp[0].ki.dwFlags = dwFlags;
                // 動作フラグ キーダウン = 0
                inp[0].ki.dwExtraInfo = IntPtr.Zero;
                // 追加情報
                inp[1] = inp[0];
                // KEYDOWN→KEYUPコピー
                inp[1].ki.dwFlags = dwFlags | KEYEVENTF_KEYUP;
                //  動作フラグ キーアップ
                // 入力送信
                SendInput(2, inp, Marshal.SizeOf(typeof(INPUT)));
            }
            finally
            {
            }
        }

        // UNICODE文字キーボード入力送信
        static internal void SendInputKeybordUNICODE(ushort wVk)
        {
            const int KEYEVENTF_UNICODE = 0x4;
            SendInputKeybord(0, wVk, KEYEVENTF_UNICODE);
        }

        // 拡張キーボード入力送信
        static internal void SendInputKeybordExtend(ushort wVk)
        {
            const int KEYEVENTF_EXTENDEDKEY = 0x1;
            SendInputKeybord(wVk, Convert.ToUInt16(MapVirtualKey(wVk, 0)), KEYEVENTF_EXTENDEDKEY);
        }

        // 拡張キーボード入力送信 スキャンコード指定
        static internal void SendInputKeybordExtend(ushort wVk, ushort wScan)
        {
            const int KEYEVENTF_EXTENDEDKEY = 0x1;
            SendInputKeybord(wVk, wScan, KEYEVENTF_EXTENDEDKEY);
        }
    }

    //--------------------------------------------------------
    // 共通機能クラス
    //--------------------------------------------------------
    class CommonFunc
    {
        // ・元情報構造体
        private struct ORIGINS
        {
            // 横幅
            public int Width;
            // 縦幅
            public int Height;
            // x座標
            public int X;
            // y座標
            public int Y;
            // フォントサイズ
            public float FontSize;
        }
        //--------------------------------------------------------
        // メンバ変数
        //--------------------------------------------------------
        // 描画文字列
        public string Text;
        // 文字書式
        public StringFormat StringFmt;

        // 基本文字書式
        private static StringFormat BaseStringFmt;
        // 自コントロール
        private Control Ctrl;
        // 元情報
        private ORIGINS Orgs;
        // フォント名
        private string FontName;

        // 共通機能クラス生成
        public CommonFunc(Control c) : base()
        {
            Ctrl = c;
            // 基本文字書式
            if (StringFmt == null)
            {
                BaseStringFmt = new StringFormat();
                BaseStringFmt.Alignment = StringAlignment.Center;
                BaseStringFmt.LineAlignment = StringAlignment.Center;
            }
            StringFmt = BaseStringFmt;
            Orgs.FontSize = 1;
            // ※初期化しないとPaintの処理でデザイナーがエラーとなる
        }

        // コントロールの元情報を取得する
        public void GetOrigins()
        {
            // ラベル以外は自前描画の為、文字列退避
            if (Ctrl.GetType().Equals(typeof(LabelEx)) == false)
            {
                Text = Ctrl.Text;
                Ctrl.Text = "";
            }
            // 初期フォント保持
            FontName = Ctrl.Font.Name;
            // 元情報を取得
            Orgs.Width = Ctrl.Width;
            Orgs.Height = Ctrl.Height;
            Orgs.X = Ctrl.Location.X;
            Orgs.Y = Ctrl.Location.Y;
            Orgs.FontSize = Ctrl.Font.Size;
        }

        // コントロールを元情報でサイズ変更、フォント変更する
        public void ResizeByOrigins(double ratio, string fname = "")
        {
            try
            {
                Point pos = default(Point);
                Ctrl.Width = Convert.ToInt32(Orgs.Width * ratio);
                Ctrl.Height = Convert.ToInt32(Orgs.Height * ratio);
                pos.X = Convert.ToInt32(Orgs.X * ratio);
                pos.Y = Convert.ToInt32(Orgs.Y * ratio);
                Ctrl.Location = pos;
                if (!string.IsNullOrEmpty(fname))
                {
                    FontName = fname;
                }
                float fontsize = Convert.ToSingle(Orgs.FontSize * ratio);
                Ctrl.Font = new Font(FontName, fontsize, Ctrl.Font.Style);
            }
            finally
            {
            }
        }

        // 描画処理 ※ボタン内の余白が大きいし拡縮でずれる！ → しかたないので自前で追加描画する
        public void Paint(Graphics g)
        {
            try
            {
                Brush brush = new SolidBrush(Ctrl.ForeColor);
                Rectangle rect = Ctrl.ClientRectangle;
                rect.Y = Convert.ToInt32(3 * Ctrl.Font.Size / Orgs.FontSize);
                // 縦余白補正
                g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.SystemDefault;
                g.DrawString(Text, Ctrl.Font, brush, rect, StringFmt);
                brush.Dispose();
            }
            finally
            {
            }
        }
    }

    //--------------------------------------------------------
    // ラベルクラス
    //--------------------------------------------------------
    class LabelEx : System.Windows.Forms.Label
    {
        //--------------------------------------------------------
        // メンバ変数
        //--------------------------------------------------------
        // 共通機能クラス
        public CommonFunc Cmn;
        // 基本文字書式
        private static StringFormat BaseStringFmt;
        // ラベル作成
        public LabelEx() : base()
        {
            Cmn = new CommonFunc(this);
            // 基本文字書式
            if (BaseStringFmt == null)
            {
                BaseStringFmt = new StringFormat();
                BaseStringFmt.Alignment = StringAlignment.Near;
            }
            Cmn.StringFmt = BaseStringFmt;
        }
    }

    //--------------------------------------------------------
    // ラジオボタンクラス
    //--------------------------------------------------------
    class RadioButtonEx : System.Windows.Forms.RadioButton
    {
        //--------------------------------------------------------
        // メンバ変数
        //--------------------------------------------------------
        // 共通機能クラス
        public CommonFunc Cmn;
        // ラジオボタン作成
        public RadioButtonEx() : base()
        {
            this.SetStyle(ControlStyles.Selectable, false);
            // フォーカス回避
            Cmn = new CommonFunc(this);
        }
        // 描画処理
        protected override void OnPaint(PaintEventArgs e)
        {
            base.OnPaint(e);
            Cmn.Paint(e.Graphics);
        }
    }

    //--------------------------------------------------------
    // ボタンクラス
    //--------------------------------------------------------
    class ButtonEx : System.Windows.Forms.Button
    {
        //--------------------------------------------------------
        // メンバ変数
        //--------------------------------------------------------
        // 共通機能クラス
        public CommonFunc Cmn;
        // ボタン作成
        public ButtonEx() : base()
        {
            this.SetStyle(ControlStyles.Selectable, false);
            // フォーカス回避
            Cmn = new CommonFunc(this);
        }
        // 描画処理
        protected override void OnPaint(PaintEventArgs e)
        {
            base.OnPaint(e);
            Cmn.Paint(e.Graphics);
        }
    }

}
