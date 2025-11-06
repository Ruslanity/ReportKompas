using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml;
using System.Xml.Serialization;

namespace ReportKompas
{
    [XmlRoot("SpeedCut")]
    public class SpeedCut
    {
        private List<Key> keys = new List<Key>();

        [XmlElement("key")]
        public List<Key> Keys
        {
            get => keys;
            set
            {
                // Сортируем сразу при присваивании
                keys = value?.OrderBy(k => k.Name).ToList() ?? new List<Key>();
            }
        }

        public static SpeedCut Load(string xmlFilePath)
        {
            XmlSerializer serializer = new XmlSerializer(typeof(SpeedCut));

            if (!File.Exists(xmlFilePath))
            {
                return new SpeedCut();
            }
            else
            {
                using (StreamReader reader = new StreamReader(xmlFilePath))
                {
                    return (SpeedCut)serializer.Deserialize(reader);
                }
            }
        }

        // Сохранение настроек в файл
        public void Save(string filePath)
        {
            XmlSerializer serializer = new XmlSerializer(typeof(SpeedCut));
            using (FileStream fs = new FileStream(filePath, FileMode.Create))
            {
                serializer.Serialize(fs, this);
            }
        }
    }

    // Класс для каждого элемента <key>
    public class Key
    {
        [XmlAttribute("name")]
        public string Name { get; set; }

        [XmlAttribute("burntime")]
        public double BurnTime { get; set; }

        [XmlAttribute("material")]
        public string Material { get; set; }

        [XmlAttribute("value")]
        public int Value { get; set; }
    }
}
