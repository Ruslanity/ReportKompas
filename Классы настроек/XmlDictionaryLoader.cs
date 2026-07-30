using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Serialization;

namespace ReportKompas
{
    internal static class XmlDictionaryLoader
    {
        public static Dictionary<TKey, TItem> Load<TWrapper, TItem, TKey>(
            string path,
            string entityDescription,
            Func<TWrapper, IEnumerable<TItem>> itemsSelector,
            Func<TItem, TKey> keySelector)
        {
            if (!File.Exists(path))
                throw new FileNotFoundException("Не найден файл " + entityDescription + ": " + path, path);

            XmlSerializer serializer = new XmlSerializer(typeof(TWrapper));
            using (StreamReader reader = new StreamReader(path))
            {
                TWrapper wrapper = (TWrapper)serializer.Deserialize(reader);
                return itemsSelector(wrapper).ToDictionary(keySelector);
            }
        }
    }
}
