using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Forms;//タスクトレイにアイコンを作るために必要

using wpf = System.Windows;

namespace TaskTrayMouseShaker
{
    /// <summary>
    /// Interaction logic for App.xaml
    /// </summary>
    public partial class App : wpf.Application
    {
        private readonly string _appName = "MouseShaker";

        private NotifyIcon _icon;


        private System.Timers.Timer timer = new System.Timers.Timer();
        private readonly int interval_milliseconds = 120_000;//操作チェック間隔。ms

        private readonly string appRunningIconPath = "./images/mouse_running.ico";
        private readonly string appStopIconPath = "./images/mouse_stop.ico";


        /// <summary>
        /// アプリケーション起動時の処理
        /// </summary>
        /// <param name="e"></param>
        protected override void OnStartup(StartupEventArgs e)
        {
            base.OnStartup(e);

            #region タスクトレイにアイコンを追加

            _icon = new NotifyIcon();

            //アイコン設定
            _icon.Icon = new System.Drawing.Icon(appStopIconPath);
            _icon.Text = $"{_appName} : Stop";
            //表示する
            _icon.Visible = true;

            #region Context Menuの作成

            var menu = new ContextMenuStrip();

            //終了アイテム
            ToolStripMenuItem menuItem_Exit = new ToolStripMenuItem();
            menuItem_Exit.Text = "Exit";
            menuItem_Exit.Click += MenuItem_Exit_Click;

            //開始アイテム
            ToolStripMenuItem menuItem_Start = new ToolStripMenuItem();
            menuItem_Start.Text = "Start";
            menuItem_Start.Click += MenuItem_Start_Click;

            //終了アイテム
            ToolStripMenuItem menuItem_Stop = new ToolStripMenuItem();
            menuItem_Stop.Text = "Stop";
            menuItem_Stop.Click += MenuItem_Stop_Click;

            //Context MenuにMenuItemを追加
            menu.Items.Add(menuItem_Exit);
            menu.Items.Add(menuItem_Start);
            menu.Items.Add(menuItem_Stop);
            //Menuをタスクトレイのアイコンに追加
            _icon.ContextMenuStrip = menu;

            #endregion Context Menuの作成

            #endregion タスクトレイにアイコンを追加

            //タイマー設定
            timer.Interval = interval_milliseconds;
            timer.Elapsed += Timer_Elapsed;


            using (Process currentProcess = Process.GetCurrentProcess())
            using (ProcessModule currentModule = currentProcess.MainModule)
            {
                // メソッドをマウスのイベントに紐づける。
                _mouseHookId = NativeMethods.SetWindowsHookEx(
                    (int)NativeMethods.HookType.WH_MOUSE_LL,
                    _mouseProc,
                    NativeMethods.GetModuleHandle(currentModule.ModuleName),
                    0
                );

                // メソッドをキーボードのイベントに紐づける。
                _keyboardHookId = NativeMethods.SetWindowsHookEx(
                    (int)NativeMethods.HookType.WH_KEYBOARD_LL,
                    _keyboardProc,
                    NativeMethods.GetModuleHandle(currentModule.ModuleName),
                    0
                );
            }
        }


        // デリゲート
        private static readonly NativeMethods.LowLevelMouseKeyboardProc _mouseProc = MouseInputCallback;
        private static readonly NativeMethods.LowLevelMouseKeyboardProc _keyboardProc = KeyboardInputCallback;

        // メソッドを識別するID
        private static IntPtr _mouseHookId = IntPtr.Zero;
        private static IntPtr _keyboardHookId = IntPtr.Zero;


        //private static bool isUserOperating = false;
        private static DateTime lastOperationTimeByUser = DateTime.Now;

        // マウス操作のイベントが発生したら実行されるメソッド
        private static IntPtr MouseInputCallback(int nCode, IntPtr wParam, IntPtr lParam)
        {
            //最終操作日時を更新
            lastOperationTimeByUser = DateTime.Now;

            // マウスのイベントに紐付けられた次のメソッドを実行する。メソッドがなければ処理終了。
            return NativeMethods.CallNextHookEx(_mouseHookId, nCode, wParam, lParam);
        }


        // キーボード操作のイベントが発生したら実行されるメソッド
        private static IntPtr KeyboardInputCallback(int nCode, IntPtr wParam, IntPtr lParam)
        {
            //最終操作日時を更新
            lastOperationTimeByUser = DateTime.Now;

            // キーボードのイベントに紐付けられた次のメソッドを実行する。メソッドがなければ処理終了。
            return NativeMethods.CallNextHookEx(_keyboardHookId, nCode, wParam, lParam);
        }


        private void Timer_Elapsed(object sender, System.Timers.ElapsedEventArgs e)
        {
            var elapsedFromLastOperation = DateTime.Now - lastOperationTimeByUser;

            //最終操作日時からinterval_millisecondsの間操作がなかった場合にキー入力を行う
            if (elapsedFromLastOperation > TimeSpan.FromMilliseconds(interval_milliseconds))
            {
                SendKeys.SendWait("{NUMLOCK}");
            }
        }


        #region イベントハンドラー

        /// <summary>
        /// 開始処理
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void MenuItem_Start_Click(object sender, EventArgs e)
        {
            //タイマー開始
            if (!timer.Enabled)
            {
                timer.Start();

                _icon.Icon = new System.Drawing.Icon(appRunningIconPath);
                _icon.Text = $"{_appName} : Running";
            }
        }

        /// <summary>
        /// 終了処理
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void MenuItem_Stop_Click(object sender, EventArgs e)
        {
            //タイマー終了
            if (timer.Enabled)
            {
                timer.Stop();
                _icon.Icon = new System.Drawing.Icon(appStopIconPath);
                _icon.Text = $"{_appName} : Stop";
            }
        }

        /// <summary>
        /// アプリ終了処理
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void MenuItem_Exit_Click(object sender, EventArgs e)
        {
            System.Windows.Application.Current.Shutdown();
        }

        #endregion イベントハンドラー


        /// <summary>
        /// 接続,切断などの状態の通知を表示する
        /// </summary>
        /// <param name="title"></param>
        /// <param name="message"></param>
        /// <param name="durationByMilliseconds"></param>
        private void ShowBalloonTip(string title, string message, int durationByMilliseconds)
        {
            _icon.BalloonTipTitle = title;
            _icon.BalloonTipText = message;
            _icon.ShowBalloonTip(durationByMilliseconds);
        }

        protected override void OnExit(ExitEventArgs e)
        {
            base.OnExit(e);

            NativeMethods.UnhookWindowsHookEx(_mouseHookId);
            NativeMethods.UnhookWindowsHookEx(_keyboardHookId);
        }


    }
}