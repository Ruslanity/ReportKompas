using System;
using System.Collections.Generic;
using System.Xml.Serialization;
using System.IO;

namespace ReportKompas
{
    [XmlRoot("CodeMaterial")]
    public class CodeMaterial
    {
        [XmlElement("Key")]
        public List<KeyCodeMaterial> Keys { get; set; } = new List<KeyCodeMaterial>();

        // Метод для загрузки из XML-документа
        public static CodeMaterial Load(string xmlFilePath)
        {
            XmlSerializer serializer = new XmlSerializer(typeof(CodeMaterial));

            if (!File.Exists(xmlFilePath))
            {
                return new CodeMaterial();
            }
            else
            {
                using (StreamReader reader = new StreamReader(xmlFilePath))
                {
                    return (CodeMaterial)serializer.Deserialize(reader);
                }
            }
        }

        // Метод для сохранения в XML-документ
        // Сохранение настроек в файл
        public void Save(string filePath)
        {
            XmlSerializer serializer = new XmlSerializer(typeof(CodeMaterial));
            using (FileStream fs = new FileStream(filePath, FileMode.Create))
            {
                serializer.Serialize(fs, this);
            }
        }
    }

    public class KeyCodeMaterial
    {
        [XmlAttribute("name")]
        public string Key { get; set; }

        [XmlElement("Value")]
        public List<string> Values { get; set; } = new List<string>();
    }
}
