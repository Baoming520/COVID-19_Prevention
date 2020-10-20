namespace MailSenderApp.Tools
{
    #region Namespaces
    using System;
    using System.Text;
    using System.Threading;
    #endregion

    public class ProgressBar
    {
        public ProgressBar(ProgressBarStyle style)
        {
            this.pausedEvent = new ManualResetEvent(false);
            this.paused = false;
            this.progressChars = new char[4] { '/', '-', '\\', '|' };
            this.cursorPos = new CursorPosition() { CursorLeft = 0, CursorTop = Console.CursorTop };
            this.style = style;
            this.flag = true;
        }

        public void Start()
        {
            this.subThread = new Thread(() =>
            {
                int index = 0;
                while (true)
                {
                    if (paused)
                    {
                        this.pausedEvent.WaitOne();
                    }

                    if (index > this.progressChars.Length - 1)
                    {
                        index %= this.progressChars.Length;
                    }

                    Console.SetCursorPosition(this.cursorPos.CursorLeft, this.cursorPos.CursorTop);
                    Console.Write(this.progressChars[index]);
                    Thread.Sleep(500);
                    index++;
                }
            });
            this.subThread.Start();
        }

        public void Suspend()
        {
            this.ClearUselessProgressChar();
            this.pausedEvent.Reset();
            this.paused = true;
            if (this.style == ProgressBarStyle.Rear)
            {
                Console.Write("\r\n");
            }
        }

        public void Resume()
        {
            this.pausedEvent.Set();
            paused = false;
        }

        public void Stop()
        {
            this.subThread.Abort();
            this.ClearUselessProgressChar();
        }

        public void SetInfo(string content, ConsoleColor fontColor = ConsoleColor.White)
        {
            this.ClearUselessProgressChar();
            if (this.style == ProgressBarStyle.Head)
            {
                content += "\r\n";
                Console.ForegroundColor = fontColor;
                Console.Write(content);
                Console.ResetColor();

                this.cursorPos = new CursorPosition() { CursorLeft = 0, CursorTop = Console.CursorTop };
            }
            else
            {
                int left = 0;
                if (content.Contains("\n"))
                {
                    var lastLine = content.Substring(content.IndexOf('\n') + 1);
                    left = Encoding.Default.GetByteCount(lastLine);
                }
                else
                {

                    left = Encoding.Default.GetByteCount(content);
                }

                if (this.flag)
                {
                    this.flag = false;
                }
                else
                {
                    content = "\r\n" + content;
                }

                Console.ForegroundColor = fontColor;
                Console.WriteLine(content);
                Console.ResetColor();

                this.cursorPos = new CursorPosition() { CursorLeft = left, CursorTop = Console.CursorTop - 1 };
            }
        }

        private ManualResetEvent pausedEvent;
        private volatile bool paused;
        private CursorPosition cursorPos;
        private Thread subThread;

        private char[] progressChars;
        private ProgressBarStyle style;
        private bool flag;

        private void ClearUselessProgressChar()
        {
            Console.SetCursorPosition(this.cursorPos.CursorLeft, this.cursorPos.CursorTop);
            Console.Write(" ");
            Console.SetCursorPosition(this.cursorPos.CursorLeft, this.cursorPos.CursorTop);
        }
    }

    public enum ProgressBarStyle
    {
        Head,
        Rear
    }

    public class CursorPosition
    {
        public int CursorLeft { get; set; }
        public int CursorTop { get; set; }
        public override string ToString()
        {
            return String.Format(@"CursorLeft: {0}, CursorTop: {1}", this.CursorLeft, this.CursorTop);
        }
    }
}
