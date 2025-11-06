using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Xml;
using System.Xml.Serialization;

namespace ReportKompas
{
    [XmlRoot("Settings")]
    public class Settings : IDisposable
    {
        static string appDirectory;

        static public string DefaultPathSettings
        {
            get
            {
                if (appDirectory == null)
                {
                    appDirectory = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location) + @"\Settings\Settings.xml";
                }
                return appDirectory;
            }
        }

        [XmlElement("PathDictionaryMaterials")]
        public string PathDictionaryMaterials { get; set; }

        [XmlElement("PathDictionaryEquipment")]
        public string PathDictionaryEquipment { get; set; }

        [XmlElement("PathDictionarySpeedCut")]
        public string PathDictionarySpeedCut { get; set; }

        [XmlElement("CalcLaserCutTime")]
        public bool CalcLaserCutTime { get; set; }

        private bool disposed = false; // чтобы избежать повторного вызова Dispose

        // Загрузка настроек из файла
        public static Settings Load(string filePath)
        {
            if (!File.Exists(filePath))
            {
                // Можно вернуть новые настройки по умолчанию
                return new Settings();
            }

            XmlSerializer serializer = new XmlSerializer(typeof(Settings));
            using (FileStream fs = new FileStream(filePath, FileMode.Open))
            {
                return (Settings)serializer.Deserialize(fs);
            }
        }

        // Сохранение настроек в файл
        public void Save(string filePath)
        {
            XmlSerializer serializer = new XmlSerializer(typeof(Settings));
            using (FileStream fs = new FileStream(filePath, FileMode.Create))
            {
                serializer.Serialize(fs, this);
            }
        }

        // Реализация IDisposable
        public void Dispose()
        {
            Dispose(true);

            // Подавляем финализацию, если финализатор будет добавлен
            GC.SuppressFinalize(this);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (!disposed)
            {
                if (disposing)
                {
                    // Освобождение управляемых ресурсов, если они будут появляться
                }

                // Освобождение неуправляемых ресурсов (если бы были)

                disposed = true;
            }
        }
    }
}
