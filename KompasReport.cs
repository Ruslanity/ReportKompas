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
    public class ManagerKompas
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
        [ComRegisterFunction]
        public static void RegisterKompasLib(Type t)
        {
            try
            {
                RegistryKey regKey = Registry.LocalMachine;
                string keyName = @"SOFTWARE\Classes\CLSID\{" + t.GUID.ToString() + "}";
                regKey = regKey.OpenSubKey(keyName, true);
                regKey.CreateSubKey("Kompas Report");
                regKey = regKey.OpenSubKey("InprocServer32", true);
                regKey.SetValue(null, System.Environment.GetFolderPath(Environment.SpecialFolder.System) + @"\mscoree.dll");
                regKey.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(string.Format("При регистрации класса для COM-Interop произошла ошибка:\n{0}", ex));
            }
        }

        // Эта функция удаляет раздел Kompas_Library из реестра
        [ComUnregisterFunction]
        public static void UnregisterKompasLib(Type t)
        {
            RegistryKey regKey = Registry.LocalMachine;
            string keyName = @"SOFTWARE\Classes\CLSID\{" + t.GUID.ToString() + "}";
            RegistryKey subKey = regKey.OpenSubKey(keyName, true);
            subKey.DeleteSubKey("Kompas Report");
            subKey.Close();
        }
        #endregion
    }
}
