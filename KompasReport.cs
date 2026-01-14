using Kompas6API5;
using KompasAPI7;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Management;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ReportKompas
{
    [ClassInterface(ClassInterfaceType.AutoDual)]
    public class KompasReport
    {

        [return: MarshalAs(UnmanagedType.BStr)]
        public string GetLibraryName()
        {
            return "Kompas Report";
        }
        // Головная функция библиотеки
        public void ExternalRunCommand([In] short command, [In] short mode, [In, MarshalAs(UnmanagedType.IDispatch)] object kompas_)
        {            
            string motherboardSerial = string.Empty;
            motherboardSerial = Environment.UserName + Environment.OSVersion + Environment.MachineName;

            if (File.Exists(System.Windows.Forms.Application.StartupPath.ToString() + "\\bug"))
            {
                string lic = File.ReadAllText(System.Windows.Forms.Application.StartupPath.ToString() + "\\bug");
                string licHex = DecodeFromHex(lic);
                if (DecodeFromBase64(licHex) == motherboardSerial)
                {
                    ReportKompas.GetInstance().Show();
                }
                else
                {
                    MessageBox.Show("Лицензии нет");
                }
            }
            else
            {
                MessageBox.Show("Отсутствует файл лицензии");
            }
        }

        public string GenerateMS()
        {
            string motherboardSerial = string.Empty;
            try
            {
                ManagementObjectSearcher searcher = new ManagementObjectSearcher("SELECT SerialNumber FROM Win32_BaseBoard");
                foreach (ManagementObject obj in searcher.Get())
                {
                    motherboardSerial = obj["SerialNumber"].ToString();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            return motherboardSerial;
        }

        public string DecodeFromBase64(string base64EncodedData)
        {
            byte[] base64EncodedBytes = Convert.FromBase64String(base64EncodedData);
            return Encoding.UTF8.GetString(base64EncodedBytes);
        }

        static string DecodeFromHex(string hex)
        {
            // Проверяем, что длина строки четная
            if (hex.Length % 2 != 0)
                throw new ArgumentException("Hex string must have an even length.");

            byte[] bytes = new byte[hex.Length / 2];

            for (int i = 0; i < hex.Length; i += 2)
            {
                // Преобразуем каждую пару символов в байт
                bytes[i / 2] = Convert.ToByte(hex.Substring(i, 2), 16);
            }

            // Преобразуем массив байтов обратно в строку
            return Encoding.UTF8.GetString(bytes);
        }

        #region COM Registration
        // Эта функция выполняется при регистрации класса для COM
        // Она добавляет в ветку реестра компонента раздел Kompas_Library,
        // который сигнализирует о том, что класс является приложением Компас,
        // а также заменяет имя InprocServer32 на полное, с указанием пути.
        // Все это делается для того, чтобы иметь возможность подключить
        // библиотеку на вкладке ActiveX.
        // Также выполняется регистрация в KOMPAS AddIns для автоматического подключения.
        [ComRegisterFunction]
        public static void RegisterKompasLib(Type t)
        {
            try
            {
                // COM-регистрация для KOMPAS
                RegistryKey regKey = Registry.LocalMachine;
                string keyName = @"SOFTWARE\Classes\CLSID\{" + t.GUID.ToString() + "}";
                regKey = regKey.OpenSubKey(keyName, true);
                regKey.CreateSubKey("Kompas Report");
                regKey = regKey.OpenSubKey("InprocServer32", true);
                regKey.SetValue(null, System.Environment.GetFolderPath(Environment.SpecialFolder.System) + @"\mscoree.dll");
                regKey.Close();

                // Регистрация в KOMPAS AddIns для автоматического подключения
                RegisterKompasAddIn(t);
            }
            catch (Exception ex)
            {
                MessageBox.Show(string.Format("При регистрации класса для COM-Interop произошла ошибка:\n{0}", ex));
            }
        }

        // Пути реестра для разных версий KOMPAS-3D
        private static readonly string[] KompasAddInsPaths = new string[]
        {
            @"SOFTWARE\ASCON\KOMPAS-3D\AddIns\Kompas Report",       // Общий путь для KOMPAS (v22 и новее)
            @"SOFTWARE\ASCON\KOMPAS-3D\18.0\AddIns\Kompas Report",  // KOMPAS v18
        };

        /// <summary>
        /// Регистрирует библиотеку в KOMPAS AddIns для автоматического подключения
        /// </summary>
        private static void RegisterKompasAddIn(Type t)
        {
            // Получаем путь к сборке
            string assemblyPath = t.Assembly.Location;

            // Регистрируем для всех поддерживаемых версий KOMPAS
            foreach (string addInsPath in KompasAddInsPaths)
            {
                RegisterAddInForPath(addInsPath, assemblyPath);
            }
        }

        /// <summary>
        /// Регистрирует AddIn по указанному пути реестра
        /// </summary>
        private static void RegisterAddInForPath(string addInsPath, string assemblyPath)
        {
            try
            {
                // Пробуем регистрацию в HKLM (для всех пользователей, требует прав администратора)
                try
                {
                    using (RegistryKey addInsKey = Registry.LocalMachine.CreateSubKey(addInsPath))
                    {
                        if (addInsKey != null)
                        {
                            // ProgID - идентификатор COM-компонента
                            addInsKey.SetValue("ProgID", "ReportKompas.KompasReport");
                            // Path - путь к библиотеке
                            addInsKey.SetValue("Path", assemblyPath);
                            // AutoConnect - автоматическое подключение при запуске KOMPAS
                            addInsKey.SetValue("AutoConnect", 1, RegistryValueKind.DWord);
                        }
                    }
                }
                catch
                {
                    // Если нет прав на HKLM, регистрируем в HKCU (для текущего пользователя)
                    using (RegistryKey addInsKey = Registry.CurrentUser.CreateSubKey(addInsPath))
                    {
                        if (addInsKey != null)
                        {
                            addInsKey.SetValue("ProgID", "ReportKompas.KompasReport");
                            addInsKey.SetValue("Path", assemblyPath);
                            addInsKey.SetValue("AutoConnect", 1, RegistryValueKind.DWord);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                // Не показываем ошибку, т.к. основная COM-регистрация уже выполнена
                System.Diagnostics.Debug.WriteLine("Ошибка регистрации в KOMPAS AddIns (" + addInsPath + "): " + ex.Message);
            }
        }

        // Эта функция удаляет раздел Kompas_Library из реестра
        // и удаляет регистрацию из KOMPAS AddIns
        [ComUnregisterFunction]
        public static void UnregisterKompasLib(Type t)
        {
            try
            {
                // Удаление COM-регистрации
                RegistryKey regKey = Registry.LocalMachine;
                string keyName = @"SOFTWARE\Classes\CLSID\{" + t.GUID.ToString() + "}";
                RegistryKey subKey = regKey.OpenSubKey(keyName, true);
                if (subKey != null)
                {
                    subKey.DeleteSubKey("Kompas Report", false);
                    subKey.Close();
                }

                // Удаление из KOMPAS AddIns
                UnregisterKompasAddIn();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("Ошибка при удалении регистрации: " + ex.Message);
            }
        }

        /// <summary>
        /// Удаляет регистрацию библиотеки из KOMPAS AddIns
        /// </summary>
        private static void UnregisterKompasAddIn()
        {
            // Удаляем для всех поддерживаемых версий KOMPAS
            foreach (string addInsPath in KompasAddInsPaths)
            {
                UnregisterAddInForPath(addInsPath);
            }
        }

        /// <summary>
        /// Удаляет AddIn по указанному пути реестра
        /// </summary>
        private static void UnregisterAddInForPath(string fullPath)
        {
            try
            {
                // Пробуем удалить из HKLM
                Registry.LocalMachine.DeleteSubKey(fullPath, false);
            }
            catch { }

            try
            {
                // Пробуем удалить из HKCU
                Registry.CurrentUser.DeleteSubKey(fullPath, false);
            }
            catch { }
        }
        #endregion
    }
}
