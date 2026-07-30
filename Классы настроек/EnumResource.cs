using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Serialization;

namespace ReportKompas
{
    public enum ResourceName
    {
        BendOperator,
        LaserCutOperator,
        PaintOperator,
        Welder,
        Locksmith,

        BendMachine,        
        LaserCutMachine1500,
        LaserCutMachine3000,
        PaintBooth        
    }

    public class EnumResource
    {
        [XmlAttribute("Id")]           public string Id           { get; set; }
        [XmlAttribute("Name")]         public string Name         { get; set; }
        [XmlAttribute("IDTurbo")]      public string IDTurbo      { get; set; }
        [XmlAttribute("ResourceType")] public string ResourceType { get; set; }
        [XmlAttribute("Designation")]  public string Designation  { get; set; }

        public static Dictionary<ResourceName, EnumResource> Load(string path)
        {
            return XmlDictionaryLoader.Load<EnumResourceList, EnumResource, ResourceName>(
                path,
                "справочника ресурсов",
                w => w.Items,
                d => (ResourceName)Enum.Parse(typeof(ResourceName), d.Id));
        }

        [XmlRoot("Resources")]
        public class EnumResourceList
        {
            [XmlElement("Resource")]
            public List<EnumResource> Items { get; set; } = new List<EnumResource>();
        }
    }
}
