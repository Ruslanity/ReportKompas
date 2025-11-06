using System;
using System.Collections.Generic;
using System.Xml.Serialization;
using System.IO;

namespace ReportKompas
{
    [XmlRoot("CodeEquip")]
    public class CodeEquip
    {
        [XmlElement("Key")]
        public List<KeyCodeEquip> Keys { get; set; } = new List<KeyCodeEquip>();

        // Метод для загрузки из XML-документа
        public static CodeEquip Load(string xmlFilePath)
        {
            XmlSerializer serializer = new XmlSerializer(typeof(CodeEquip));

            if (!File.Exists(xmlFilePath))
            {
                return new CodeEquip();
            }
            else
            {
                using (StreamReader reader = new StreamReader(xmlFilePath))
                {
                    return (CodeEquip)serializer.Deserialize(reader);
                }
            }
        }

        // Метод для сохранения в XML-документ
        // Сохранение настроек в файл
        public void Save(string filePath)
        {
            XmlSerializer serializer = new XmlSerializer(typeof(CodeEquip));
            using (FileStream fs = new FileStream(filePath, FileMode.Create))
            {
                serializer.Serialize(fs, this);
            }
        }
    }

    public class KeyCodeEquip
    {
        [XmlAttribute("name")]
        public string Key { get; set; }

        [XmlElement("Value")]
        public List<string> Values { get; set; } = new List<string>();
    }
}
