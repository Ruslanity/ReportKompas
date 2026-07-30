using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;

namespace ReportKompas
{
    /// <summary>
    /// Фоновый перехватчик модальных диалогов PLM "Союз-PLM" с предложением
    /// "Начать редактирование локальных копий и заблокировать в PLM?".
    /// Пока запущен, периодически ищет такое окно и автоматически нажимает "Нет",
    /// чтобы оно не прерывало автоматическую обработку файлов.
    ///
    /// Работает в отдельном потоке: диалог PLM модальный и блокирует UI-поток,
    /// поэтому нажатие должно приходить извне (через оконные сообщения Win32).
    /// </summary>
    public sealed class PlmDialogAutoNo
    {
        #region WinAPI

        private delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

        [DllImport("user32.dll")]
        private static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);

        [DllImport("user32.dll")]
        private static extern bool EnumChildWindows(IntPtr hWndParent, EnumWindowsProc lpEnumFunc, IntPtr lParam);

        [DllImport("user32.dll")]
        private static extern bool IsWindowVisible(IntPtr hWnd);

        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        private static extern int GetWindowTextW(IntPtr hWnd, StringBuilder lpString, int nMaxCount);

        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        private static extern int GetClassNameW(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);

        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        private static extern IntPtr SendMessageW(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);

        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        private static extern bool PostMessageW(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);

        private const uint BM_CLICK = 0x00F5;
        private const uint WM_COMMAND = 0x0111;
        private const int IDNO = 7; // стандартный id кнопки "Нет" в MessageBox

        #endregion

        // Признаки нужных диалогов. Заголовок — окна PLM (подходит и "Союз-PLM", и "Союз-PLM v3");
        // маркеры — устойчивые подстроки текста (без меняющихся имён файлов). Совпадение любого
        // маркера считает окно целевым — на все эти диалоги нужно отвечать "Нет".
        private static readonly string TitleMarker = "Союз-PLM";
        private static readonly string[] TextMarkers =
        {
            // Диалог "Начать редактирование локальных копий и заблокировать в PLM?"
            "редактирование локальных копий",
            "заблокировать в PLM",
            // Диалог "Для следующих документов в PLM есть обновление: ... Хотите их обновить?"
            "есть обновление",
            "Хотите их обновить"
        };

        private Thread _thread;
        private volatile bool _running;
        private readonly int _pollIntervalMs;

        public PlmDialogAutoNo(int pollIntervalMs = 120)
        {
            _pollIntervalMs = pollIntervalMs;
        }

        public bool IsRunning => _running;

        public void Start()
        {
            if (_running) return;
            _running = true;
            _thread = new Thread(Loop)
            {
                IsBackground = true,
                Name = "PlmDialogAutoNo"
            };
            _thread.Start();
        }

        public void Stop()
        {
            _running = false;
            _thread = null;
        }

        private void Loop()
        {
            while (_running)
            {
                try
                {
                    ScanOnce();
                }
                catch
                {
                    // перехватчик не должен ронять приложение
                }
                Thread.Sleep(_pollIntervalMs);
            }
        }

        private void ScanOnce()
        {
            EnumWindows((hWnd, _) =>
            {
                if (!_running) return false; // прекратить перечисление
                if (!IsWindowVisible(hWnd)) return true;

                string title = GetText(hWnd);
                if (string.IsNullOrEmpty(title) ||
                    title.IndexOf(TitleMarker, StringComparison.OrdinalIgnoreCase) < 0)
                {
                    return true;
                }

                if (IsTargetDialog(hWnd))
                {
                    ClickNo(hWnd);
                }
                return true;
            }, IntPtr.Zero);
        }

        /// <summary>
        /// Подтверждает, что это именно диалог про редактирование локальных копий,
        /// проверяя текст статических подписей внутри окна.
        /// </summary>
        private bool IsTargetDialog(IntPtr hWnd)
        {
            var sb = new StringBuilder();
            EnumChildWindows(hWnd, (child, _) =>
            {
                if (string.Equals(GetClassName(child), "Static", StringComparison.OrdinalIgnoreCase))
                {
                    sb.Append(GetText(child)).Append('\n');
                }
                return true;
            }, IntPtr.Zero);

            string text = sb.ToString();
            foreach (string marker in TextMarkers)
            {
                if (text.IndexOf(marker, StringComparison.OrdinalIgnoreCase) >= 0)
                    return true;
            }
            return false;
        }

        private void ClickNo(IntPtr hWnd)
        {
            // 1) Пытаемся найти кнопку с текстом "Нет" и кликнуть по ней.
            IntPtr noButton = FindNoButton(hWnd);
            if (noButton != IntPtr.Zero)
            {
                SendMessageW(noButton, BM_CLICK, IntPtr.Zero, IntPtr.Zero);
                return;
            }

            // 2) Фоллбэк: стандартный MessageBox закрывается командой IDNO.
            PostMessageW(hWnd, WM_COMMAND, (IntPtr)IDNO, IntPtr.Zero);
        }

        private IntPtr FindNoButton(IntPtr hWnd)
        {
            IntPtr found = IntPtr.Zero;
            EnumChildWindows(hWnd, (child, _) =>
            {
                if (!string.Equals(GetClassName(child), "Button", StringComparison.OrdinalIgnoreCase))
                    return true;

                string text = GetText(child).Replace("&", "").Trim();
                if (string.Equals(text, "Нет", StringComparison.OrdinalIgnoreCase))
                {
                    found = child;
                    return false; // нашли — прекратить
                }
                return true;
            }, IntPtr.Zero);
            return found;
        }

        private static string GetText(IntPtr hWnd)
        {
            var sb = new StringBuilder(512);
            GetWindowTextW(hWnd, sb, sb.Capacity);
            return sb.ToString();
        }

        private static string GetClassName(IntPtr hWnd)
        {
            var sb = new StringBuilder(256);
            GetClassNameW(hWnd, sb, sb.Capacity);
            return sb.ToString();
        }
    }
}
