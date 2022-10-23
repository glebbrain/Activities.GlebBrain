using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Threading;

using BR.Core;
using BR.Core.Attributes;
using BR.Core.Base;

namespace Activities.GlebBrain
{

    [ScreenName("Окно активное ?")] // Имя активности, отображаемое в списке активностей и в заголовке шага
    [Representation(
        "Заголовок окна: [Title]\r\n" +
        "Наименование процесса: [ProcessName]\r\n" +
        "Окно активное? [IsActive]"
    )] // Представление шага. Позиции в квадратных скобках заменяются значениями свойств.
    [BR.Core.Attributes.Path("GlebBrain.Windows")] // Путь к активности в панели "Активности"
    public class IsActiveWindow : BR.Core.Activity
    {
        /// <summary>
        /// Заголовок окна
        /// </summary>
        [ScreenName("Заголовок окна")]
        [Description("Часть(\"Блокнот\") или полный заголовок(\"Безымянный – Блокнот\")")]
        public string Title { get; set; }
        /// <summary>
        /// Класс окна
        /// </summary>
        [ScreenName("Наименование процесса")]
        [Description("Часть(\"note\") или полное наименование(\"notepad.exe\")")]
        public string ProcessName { get; set; }
        /// <summary>
        /// Окно активно
        /// </summary>
        [ScreenName("Окно активно")]
        [Description("True - да, False - подробности в 'Описание ошибки'")]
        [IsOut]
        public bool IsActive { get; set; }
        /// <summary>
        /// Описание ошибки
        /// </summary>
        [ScreenName("Описание ошибки")]
        [Description("Ошибка перехваченная через try..catch..")]
        [IsOut]
        public string ExceptionMessage { get; set; }
        /// <summary>
        /// DS
        /// </summary>
        [ScreenName("Этапы работы скрипта")]
        [Description("Используется для отладки")]
        [IsOut]
        public string DS { get; set; }

        [System.Runtime.InteropServices.DllImport("user32.dll", CharSet = CharSet.Auto, ExactSpelling = true)]
        public static extern IntPtr GetForegroundWindow();

        [System.Runtime.InteropServices.DllImport("user32.dll", CharSet = CharSet.Auto, ExactSpelling = true)]
        public static extern IntPtr GetWindowThreadProcessId(IntPtr hWnd, out int lpdwProcessId);

        public override void Execute(int? optionID)
        {
            if (Title == null)
            {
                Title = "";
            }
            if (ProcessName == null)
            {
                ProcessName = "";
            }

            IsActive = false;
            ExceptionMessage = "";
            DS = "0|Activities.GlebBrain.IsWindowActive.Execute\r\n" +
                        "Title='" + Title.ToString() + "'\r\n" +
                        "ProcessName='" + ProcessName.ToString() + "';\r\n";
            try
            {
                
                DS += "1:получаем хендел активного окна;\r\n";
                IntPtr hWnd = GetForegroundWindow();
                int pid;
                DS += "2:получаем pid потока активного окна;\r\n";
                GetWindowThreadProcessId(hWnd, out pid);
                DS += "3:ищем среди всех процессов необходимый pid и проверяем процесс по Title и ProcessName;\r\n";
                using (System.Diagnostics.Process p = System.Diagnostics.Process.GetProcessById(pid))
                {
                    if ((Title != "" && (p.MainWindowTitle == Title || p.MainWindowTitle.IndexOf(Title) > -1))
                        ||
                        (ProcessName != "" && 
                            (
                                p.ProcessName.ToLower().Trim() == ProcessName.Replace(".exe", "").ToLower().Trim()
                                || p.ProcessName.ToLower().IndexOf(ProcessName.ToLower().Replace(".exe", "").Trim()) > -1)
                            )
                        )
                    {
                        IsActive = true;
                        DS += "4:процесс найден;\r\n";
                    }
                    else
                    {
                        DS += "5:Активное окно: '"+ p.ProcessName + "' - '"+ p.MainWindowTitle + "';\r\n";
                    }
                }
                DS += "6:Активность отработала успешно;\r\n";
            }
            catch (Exception ex)
            {
                DS += "7:Exception;\r\n";
                IsActive = false;
                ExceptionMessage =
                    "Пройденные этапы: '" + DS + "';\r\n" +
                    "Message: '" + ex.Message + "';\r\n" +
                    "Source: '" + ex.Source + "';\r\n" +
                    "StackTrace: '" + ex.StackTrace + "';\r\n";
            }
        }
    }

    [ScreenName("Сделать окно активным")] // Имя активности, отображаемое в списке активностей и в заголовке шага
    [Representation(
        "Заголовок окна: [Title]\r\n" +
        "Наименование процесса: [ProcessName]\r\n" +
        "Окно активное? [IsActive]"
    )] // Представление шага. Позиции в квадратных скобках заменяются значениями свойств.
    [BR.Core.Attributes.Path("GlebBrain.Windows")] // Путь к активности в панели "Активности"
    public class SetActivesWindow : BR.Core.Activity
    {
        /// <summary>
        /// Заголовок окна
        /// </summary>
        [ScreenName("Заголовок окна")]
        [Description("Часть(\"Блокнот\") или полный заголовок(\"Безымянный – Блокнот\")")]
        public string Title { get; set; }
        /// <summary>
        /// Класс окна
        /// </summary>
        [ScreenName("Наименование процесса")]
        [Description("Часть(\"note\") или полное наименование(\"notepad.exe\")")]
        public string ProcessName { get; set; }
        /// <summary>
        /// Окно активно
        /// </summary>
        [ScreenName("Окно активно")]
        [Description("True - да, False - подробности в 'Описание ошибки'")]
        [IsOut]
        public bool IsActive { get; set; }
        /// <summary>
        /// Описание ошибки
        /// </summary>
        [ScreenName("Описание ошибки")]
        [Description("Ошибка перехваченная через try..catch..")]
        [IsOut]
        public string ExceptionMessage { get; set; }
        /// <summary>
        /// DS
        /// </summary>
        [ScreenName("Этапы работы скрипта")]
        [Description("Используется для отладки")]
        [IsOut]
        public string DS { get; set; }
        [System.Runtime.InteropServices.DllImport("user32.dll")]
        public static extern IntPtr GetForegroundWindow();
        

        [System.Runtime.InteropServices.DllImport("user32.dll")]
        public static extern IntPtr SetForegroundWindow(IntPtr hWnd);
        [System.Runtime.InteropServices.DllImport("user32.dll")]
        public static extern IntPtr SetFocus(IntPtr hWnd);
        [System.Runtime.InteropServices.DllImport("user32.dll")]
        public static extern IntPtr SetActiveWindow(IntPtr hWnd);
        
        [System.Runtime.InteropServices.DllImport("user32.dll")]
        public static extern IntPtr SetWindowPos(IntPtr hWnd, int a, int b, int d, int e, int f, int g);

        [System.Runtime.InteropServices.DllImport("user32.dll")]
        public static extern IntPtr GetWindowThreadProcessId(IntPtr hWnd, out int lpdwProcessId);

        private const int SW_MAXIMIZE = 3;
        private const int SW_MINIMIZE = 6;
        private const int SW_RESTORE = 9;
        // more here: http://www.pinvoke.net/default.aspx/user32.showwindow

        [DllImport("user32.dll", EntryPoint = "FindWindow")]
        public static extern IntPtr FindWindowByCaption(IntPtr ZeroOnly, string lpWindowName);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        public override void Execute(int? optionID)
        {
            if (Title == null)
            {
                Title = "";
            }
            if (ProcessName == null)
            {
                ProcessName = "";
            }

            IsActive = false;
            ExceptionMessage = "";
            DS = "0|Activities.GlebBrain.SetActivesWindow.Execute\r\n" +
                        "Title='" + Title.ToString() + "'\r\n" +
                        "ProcessName='" + ProcessName.ToString() + "';\r\n";
            try
            {
                DS += "1:ищем процесс по Title или ProcessName и активируем его;\r\n";
                foreach (System.Diagnostics.Process p in System.Diagnostics.Process.GetProcesses())
                {
                    if ((Title != "" && (p.MainWindowTitle == Title || p.MainWindowTitle.IndexOf(Title) > -1))
                        ||
                        (ProcessName != "" &&
                            (
                                p.ProcessName.ToLower().Trim() == ProcessName.Replace(".exe", "").ToLower().Trim()
                                || p.ProcessName.ToLower().IndexOf(ProcessName.ToLower().Replace(".exe", "").Trim()) > -1
                            )
                        )
                       )
                    {
                        if (p.MainWindowHandle != IntPtr.Zero)
                        {
                            ShowWindow(p.MainWindowHandle, SW_RESTORE);
                            //SetWindowPos(p.MainWindowHandle, 0, 0, 0, 0, 0, SWP_NOZORDER | SWP_NOSIZE | SWP_SHOWWINDOW);
                            SetForegroundWindow(p.MainWindowHandle);
                            SetFocus(p.MainWindowHandle);
                            SetActiveWindow(p.MainWindowHandle);
                        }
                        DS += "2:процесс найден: '" + p.ProcessName + "' - '" + p.MainWindowTitle + "';\r\n";
                    }
                }
                Thread.Sleep(100);

                IsActive = false;
                DS += "4:получаем хендел активного окна;\r\n";
                IntPtr hWnd = GetForegroundWindow();
                int pid;
                DS += "5:получаем pid потока активного окна;\r\n";
                GetWindowThreadProcessId(hWnd, out pid);
                DS += "6:ищем среди всех процессов необходимый pid и проверяем процесс по Title и ProcessName;\r\n";
                using (System.Diagnostics.Process p = System.Diagnostics.Process.GetProcessById(pid))
                {
                    if ((Title != "" && (p.MainWindowTitle == Title || p.MainWindowTitle.IndexOf(Title) > -1))
                        ||
                        (ProcessName != "" && (
                            p.ProcessName.ToLower().Trim() == ProcessName.Replace(".exe", "").ToLower().Trim()
                            || p.ProcessName.ToLower().IndexOf(ProcessName.ToLower().Replace(".exe", "").Trim()) > -1
                        ))
                        )
                    {
                        IsActive = true;
                        DS += "7:процесс найден;\r\n";
                    }
                    else
                    {
                        IsActive = false;
                        DS += "8:Активное окно: '" + p.ProcessName + "' - '" + p.MainWindowTitle + "';\r\n";
                    }
                }
                DS += "9:Активность отработала успешно;\r\n";
            }
            catch (Exception ex)
            {
                DS += "10:Exception!;\r\n";
                IsActive = false;
                ExceptionMessage =
                    "Пройденные этапы: '" + DS + "';\r\n" +
                    "Message: '" + ex.Message + "';\r\n" +
                    "Source: '" + ex.Source + "';\r\n" +
                    "StackTrace: '" + ex.StackTrace + "';\r\n";
            }
        }
    }
}
