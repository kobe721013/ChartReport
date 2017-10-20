using System;
using System.Diagnostics;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;


namespace MiscUtils
{
    public class Log
    {
        private class Format
        {
            private const int Alignment = 8;
            public int ColumnWidth;
            public int ColumnPosition;
            public string Message;

            public void Update(Format column0, string s)
            {
                ColumnPosition = column0.ColumnPosition + column0.ColumnWidth;
                ColumnWidth = Math.Max(ColumnWidth, ((s.Length + Alignment - 1) / Alignment) * Alignment);
                var builder = new StringBuilder(s).Append(' ', ColumnWidth - (s.Length > ColumnWidth ? ColumnWidth : s.Length));
                Message = builder.ToString();
            }
        }

        public static bool UsingTaskVerForD = false;
        public static bool UsingTaskVerForE = false;

        private static Format _column0 = new Format();
        private static Format _column1 = new Format();
        private static Format _column2 = new Format();
        private static Format _column3 = new Format();

        private static Task _lastTask = null;
        public static void D(string message,
            [CallerFilePath] string file = "",
            [CallerLineNumber] int line = 0,
            [CallerMemberName] string member = "")
        {

            string info = string.Format("[{0},{1} {2}]", member, System.IO.Path.GetFileName(file), line);
            DebugPrint(LEVEL.DEBUG, message, info);
#if false
            if (UsingTaskVerForD)
            {
                if (_lastTask == null)
                {
                    _lastTask = Task.Factory.StartNew(() =>
                    {
                        LogDump(message, file, line, member);
                    });
                }
                else
                {
                    _lastTask = _lastTask.ContinueWith(task =>
                    {
                        LogDump(message, file, line, member);
                    });
                }
            }
            else
            {
                LogDump(message, file, line, member);


            }
#endif
        }

        

        private static void LogDump(string message, string file, int line, string member)
        {
            _column1.Update(_column0, $"{file}({line}):");
            _column2.Update(_column1, $"[{member}");
            _column3.Update(_column2, $"] {message}");

            //Debug.WriteLine($"{_column1.Message}{_column2.Message}{_column3.Message}");
            //Tools.GLogW(file, member, line, "[TL] %ls\n", message);
            Debug.Print("[TL] " + message);
        }

        public static void E(Exception exception,
            [CallerFilePath] string file = "",
            [CallerLineNumber] int line = 0,
            [CallerMemberName] string member = "")
        {
            if (UsingTaskVerForE)
            {
                if (_lastTask == null)
                {
                    Task.Factory.StartNew(() =>
                    {
                        ExceptionDump(exception, file, line, member);
                    });
                }
                else
                {
                    _lastTask = _lastTask.ContinueWith(task =>
                    {
                        ExceptionDump(exception, file, line, member);
                    });
                }
            }
            else
            {
                ExceptionDump(exception, file, line, member);
            }
        }

        private static void ExceptionDump(Exception exception, string file, int line, string member)
        {
            //_column1.Update(_column0, $"{file}({line}):");
            //_column2.Update(_column1, $"[{member}");
            //_column3.Update(_column2, $"] {exception}");
            //Debug.WriteLine($"{_column1.Message}{_column2.Message}{_column3.Message}");
            //Tools.GLogW(file, member, line, "%ls\n", exception.ToString());
        }

        ///////////////
        public static void Error(string message,
            [CallerFilePath] string file = "",
            [CallerLineNumber] int line = 0,
            [CallerMemberName] string member = "")
        {
            string info = string.Format("[{0},{1} {2}]", member, System.IO.Path.GetFileName(file), line);
            ReleasePrint(LEVEL.ERROR, message, info);
        }

        

        private static void DebugPrint(LEVEL level, string message, string info)
        {
            System.Diagnostics.Debug.Print("{0}[{1,5}]: {2} {3}", DateTime.Now.ToString("HHmmss"), level.ToString(), message, info);
        }

        private static void ReleasePrint(LEVEL level, string message, string info)
        {
            System.Diagnostics.Trace.WriteLine(string.Format("{0}[{1,5}]: {2} {3}", DateTime.Now.ToString("HHmmss"), level.ToString(), message, info));
        }

        private static void TracePrint(LEVEL level, string message, string info)
        {
            System.Diagnostics.Trace.WriteLine(String.Format("{0}[{1,5}]: {2} {3}", DateTime.Now.ToString("HHmmss"), level.ToString(), message, info));
        }

        private static bool _ConsoleOutput = false;
        public static bool ConsoleOutput
        {
            get { return _ConsoleOutput; }
            set
            {
                _ConsoleOutput = value;
                if (value)
                {
                    Debug.Listeners.Remove(LOG_TYPE.CONSOLE.ToString());
                    Debug.Listeners.Remove("Default");
                    DefaultTraceListener d = new DefaultTraceListener { Name = LOG_TYPE.CONSOLE.ToString() };
                    Debug.Listeners.Add(d);
                }
                else
                {
                    Debug.Listeners.RemoveAt((int)LOG_TYPE.CONSOLE);
                }
            }
        }


        public static void FileOutput(bool enable, string folderPath, string fileName)
        {
            if (enable)
            {
                Debug.Listeners.Remove(LOG_TYPE.FILE.ToString());
                System.IO.Directory.CreateDirectory(folderPath);
                string name = String.Format("{0}_{1}.log", DateTime.Now.ToString("yyyyMMdd_HHmmss"), fileName);
                string logPath = System.IO.Path.Combine(folderPath, name);

                TextWriterTraceListener f = new TextWriterTraceListener(logPath);
                f.Name = LOG_TYPE.FILE.ToString();
                //TextWriterTraceListener[] listeners = new TextWriterTraceListener[] {
                //   new TextWriterTraceListener(logPath)};
                Debug.Listeners.Add(f);
                Debug.AutoFlush = true;

            }
            else
            {
                Debug.Listeners.Remove(LOG_TYPE.FILE.ToString());
            }

        }




        

        /*
        public static void User(string message)
        {

            Controller.Etw.Log.User(message);
        }
        */

        public enum LOG_TYPE
        {
            CONSOLE,
            FILE
        }
        public enum LEVEL
        {
            DEBUG,
            INFO,
            WARN,
            ERROR,
            FATAL,
        }

    }
}
