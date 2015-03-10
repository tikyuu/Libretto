using System;
using System.Diagnostics;

namespace LibrettoCreateTool
{
    /// <summary>
    /// デバッグでの時間測定用クラス
    /// http://gushwell.ldblog.jp/archives/cat_377440.html
    /// </summary>
    public sealed class KeepTime : IDisposable
    {
        private Stopwatch _stopwatch;

        /// <summary>
        /// コンストラクタ (スタート)
        /// </summary>
        /// <example>
        /// <code lang="C#">
        /// using (new KeepTime())
        ///     //実際のコード
        /// }
        /// </code>
        /// </example>
        public KeepTime()
        {
            Debug.WriteLine(String.Format("Start {0}", DateTime.Now.ToShortTimeString()));
            _stopwatch = Stopwatch.StartNew();
        }
        /// <summary>
        /// 解放処理 (終了処理)
        /// </summary>
        public void Dispose()
        {
            _stopwatch.Stop();
            Debug.WriteLine(String.Format("Stop {0} ({1}ミリ秒)", DateTime.Now.ToShortTimeString(), _stopwatch.ElapsedMilliseconds));
        }
    }
}
