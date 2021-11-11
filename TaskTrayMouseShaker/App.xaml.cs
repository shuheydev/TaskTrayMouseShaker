using System;
using System.Windows;
using System.Windows.Forms;

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

        private bool _isSlideshowOn = false;

        private Timer timer = new Timer();

        private readonly int interval_milliseconds = 300_000;

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
            _icon.Icon = new System.Drawing.Icon("./communication.ico");
            _icon.Text = _appName;
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
            timer.Interval = interval_milliseconds;//ms
            timer.Tick += Timer_Tick;
        }

        private void Timer_Tick(object sender, EventArgs e)
        {
            //マウスの座標を取得
            System.Drawing.Point dp = System.Windows.Forms.Cursor.Position;

            //1px移動させる
            MouseController.SetPosition(dp.X + 1, dp.Y + 1);

            //戻す
            MouseController.SetPosition(dp.X, dp.Y);
        }

        #region イベントハンドラー

        /// <summary>
        /// 開始処理
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private async void MenuItem_Start_Click(object sender, EventArgs e)
        {
            //タイマー開始
            if (!timer.Enabled)
            {
                timer.Start();
            }
        }

        /// <summary>
        /// 終了処理
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private async void MenuItem_Stop_Click(object sender, EventArgs e)
        {
            //タイマー終了
            if (timer.Enabled)
            {
                timer.Stop();
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
    }
}