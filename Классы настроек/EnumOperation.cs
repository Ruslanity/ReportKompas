using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Serialization;

namespace ReportKompas
{
    public enum OperationType { LaserCut, BendCNC, Welding, Locksmith, Painting, Assembly }

    public class EnumOperation
    {
        [XmlAttribute("Id")] public string Id { get; set; }
        // Значение, которое пишется в Id операции выходного XML (русский код). Не путать с Id — тот служит ключом enum.
        [XmlAttribute("Code")] public string Code { get; set; }
        [XmlAttribute("Name")] public string Name { get; set; }
        [XmlAttribute("Number")] public int Number { get; set; }

        public static Dictionary<OperationType, EnumOperation> Load(string path)
        {
            return XmlDictionaryLoader.Load<EnumOperationList, EnumOperation, OperationType>(
                path,
                "справочника операций",
                w => w.Items,
                d => (OperationType)Enum.Parse(typeof(OperationType), d.Id));
        }

        [XmlRoot("Operations")]
        public class EnumOperationList
        {
            [XmlElement("Operation")]
            public List<EnumOperation> Items { get; set; } = new List<EnumOperation>();
        }
    }
}
