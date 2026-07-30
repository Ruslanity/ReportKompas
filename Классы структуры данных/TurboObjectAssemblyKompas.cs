using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Xml.Serialization;

namespace ReportKompas
{
    public class TurboObjectAssemblyKompas
    {
        // ─── Поля (соответствуют XML-элементам) ──────────────────────────────

        // Код СИ детали/изделия (заполнен у Стандартных и Прочих изделий)
        public string Id { get; set; }
        public string Designation { get; set; }
        public string Name { get; set; }
        public string SpecificationSection { get; set; }
        public int BaseQuantity { get; set; } = 1;
        public string Mass { get; set; }
        public string OverallDimensions { get; set; }

        [XmlArray("Operations")]
        [XmlArrayItem("Operation")]
        public List<Operation> Operations { get; set; } = new List<Operation>();

        public string FullName { get; set; }
        public string PathToDXF { get; set; }

        // ─── Вложенные классы ─────────────────────────────────────────────────

        public class Operation
        {
            public string Id { get; set; }
            public string Name { get; set; }
            public int Number { get; set; }
            public string Time { get; set; }

            [XmlArray("Tools")]
            [XmlArrayItem("Tool")]
            public List<Tool> Tools { get; set; } = new List<Tool>();

            [XmlArray("Parameters")]
            [XmlArrayItem("Parameter")]
            public List<Parameter> Parameters { get; set; } = new List<Parameter>();

            [XmlArray("Resources")]
            [XmlArrayItem("Resource")]
            public List<Resource> Resources { get; set; } = new List<Resource>();
        }

        public class Tool
        {
            public string Id { get; set; }
            public string Value { get; set; }
        }

        public class Parameter
        {
            public string Id { get; set; }
            public string Value { get; set; }
        }

        public class Resource
        {
            public string Id { get; set; }
            public string Designation { get; set; }
            public string Name { get; set; }
            public string ResourceType { get; set; }
            public string Quantity { get; set; }
        }

        // ─── Обёртка для XML-корня ────────────────────────────────────────────

        [XmlRoot("Parts")]
        public class PartsList
        {
            [XmlElement("Part")]
            public List<TurboObjectAssemblyKompas> Parts { get; set; } = new List<TurboObjectAssemblyKompas>();
        }

        // ─── Сериализация / Десериализация ────────────────────────────────────

        public static List<TurboObjectAssemblyKompas> Load(string filePath)
        {
            if (!File.Exists(filePath))
                return new List<TurboObjectAssemblyKompas>();

            XmlSerializer serializer = new XmlSerializer(typeof(PartsList));
            using (StreamReader reader = new StreamReader(filePath))
            {
                return ((PartsList)serializer.Deserialize(reader)).Parts;
            }
        }

        public static void Save(IEnumerable<TurboObjectAssemblyKompas> parts, string filePath)
        {
            XmlSerializer serializer = new XmlSerializer(typeof(PartsList));
            System.Xml.XmlWriterSettings settings = new System.Xml.XmlWriterSettings
            {
                Indent   = true, // делаем читаемый отступ
                Encoding = System.Text.Encoding.UTF8
            };
            using (System.Xml.XmlWriter writer = System.Xml.XmlWriter.Create(filePath, settings))
            {
                serializer.Serialize(writer, new PartsList { Parts = parts.ToList() });
            }
        }

        // ─── Словарь определений операций (загружается из EnumOperation.xml) ──────

        // Каталог рядом с самой DLL (а не AppDomain.BaseDirectory) — иначе при загрузке как COM-библиотеки в КОМПАС-3D путь указывает на каталог KOMPAS, а не на наш аддин.
        private static readonly string AssemblyDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);

        private static readonly Dictionary<OperationType, EnumOperation> OpDefs    = EnumOperation.Load(Path.Combine(AssemblyDir, "Settings", "EnumOperation.xml"));
        private static readonly Dictionary<ResourceName, EnumResource>   ResDefs    = EnumResource.Load(Path.Combine(AssemblyDir, "Settings", "EnumResource.xml"));

        // ─── Маппинг секций спецификации → тип ресурса ───────────────────────

        private static readonly Dictionary<string, string> SectionToResourceType = new Dictionary<string, string>
        {
            { "Сборочные единицы",  "Сборочные единицы" },
            { "Детали",             "Деталь" },
            { "Стандартные изделия","Стандартные изделия" },
        };

        // ─── Конвертация из дерева ObjectAssemblyKompas ───────────────────────

        public static List<TurboObjectAssemblyKompas> FromTree(ObjectAssemblyKompas root)
        {
            var result = new List<TurboObjectAssemblyKompas>();
            CollectFlat(root, result);
            return result;
        }

        private static void CollectFlat(ObjectAssemblyKompas obj, List<TurboObjectAssemblyKompas> result)
        {
            result.Add(Convert(obj));

            if (obj.Children != null)
                foreach (var child in obj.Children)
                    CollectFlat(child, result);
        }

        private static TurboObjectAssemblyKompas Convert(ObjectAssemblyKompas obj)
        {
            var ops = BuildOperations(obj);
            return new TurboObjectAssemblyKompas
            {
                Id                   = obj.CodeEquipment,
                Designation          = obj.Designation,
                Name                 = obj.Name,
                SpecificationSection = obj.SpecificationSection,
                BaseQuantity         = 1,
                Mass                 = obj.Mass.ToString("G", CultureInfo.InvariantCulture),
                OverallDimensions    = obj.OverallDimensions,
                FullName             = obj.FullName,
                PathToDXF            = obj.PathToDXF,
                Operations           = ops
            };
        }

        private static List<Operation> BuildOperations(ObjectAssemblyKompas obj)
        {
            var ops = new List<Operation>();

            if (obj.Children != null && obj.Children.Any() && obj.SpecificationSection == "Сборочные единицы")
                ops.Add(BuildAssemblyOperation(obj));
            else
                ops.AddRange(BuildLeafOperations(obj));

            var coatingOp = BuildCoatingOperation(obj);
            if (coatingOp != null) ops.Add(coatingOp);

            var locksmithOp = BuildLocksmithOperation(obj);
            if (locksmithOp != null) ops.Add(locksmithOp);

            return ops;
        }

        private static Operation BuildAssemblyOperation(ObjectAssemblyKompas obj)
        {
            var op = new Operation
            {
                Id     = OpDefs[OperationType.Assembly].Code,
                Name   = OpDefs[OperationType.Assembly].Name,
                Number = OpDefs[OperationType.Assembly].Number,
                Time   = ""
            };

            foreach (var child in obj.Children)
            {
                string resourceType;
                if (!SectionToResourceType.TryGetValue(child.SpecificationSection ?? "", out resourceType))
                    resourceType = child.SpecificationSection ?? "";

                op.Resources.Add(new Resource
                {
                    // Код СИ заполнен только у Стандартных и Прочих изделий (см. FillCodeEquip) — у деталей/узлов пусто
                    Id           = child.CodeEquipment ?? "",
                    Designation  = child.SpecificationSection == "Стандартные изделия" ? "" : (child.Designation ?? ""),
                    Name         = child.Name ?? "",
                    ResourceType = resourceType,
                    Quantity     = child.Quantity.ToString()
                });
            }

            //если есть сварка то заменить время сборки, тут же выполняется установка кто собирает слесарь или сварщик
            double weldingTime;
            if (!string.IsNullOrEmpty(obj.Welding) && double.TryParse(obj.Welding, System.Globalization.NumberStyles.Any, CultureInfo.InvariantCulture, out weldingTime) && weldingTime > 0)
            {
                op.Time = obj.Welding;
                op.Resources.Add(new Resource
                {
                    Id           = ResDefs[ResourceName.Welder].IDTurbo,
                    Designation  = ResDefs[ResourceName.Welder].Designation,
                    Name         = ResDefs[ResourceName.Welder].Name,
                    ResourceType = ResDefs[ResourceName.Welder].ResourceType,
                    Quantity     = ""
                });
            }
            else
            {
                op.Resources.Add(new Resource
                {
                    Id           = ResDefs[ResourceName.Locksmith].IDTurbo,
                    Designation  = ResDefs[ResourceName.Locksmith].Designation,
                    Name         = ResDefs[ResourceName.Locksmith].Name,
                    ResourceType = ResDefs[ResourceName.Locksmith].ResourceType,
                    Quantity     = ""
                });
            }

            return op;
        }

        private static IEnumerable<Operation> BuildLeafOperations(ObjectAssemblyKompas obj)
        {
            //если в материале есть слово "Лист" и время резки >0, то добавится операция лазерной резки
            double tc;
            if (obj.Material != null && obj.Material.Contains("Лист") && double.TryParse(obj.TimeCut, NumberStyles.Any, CultureInfo.InvariantCulture, out tc) && tc > 0)
            {
                var op = new Operation
                {
                    Id     = OpDefs[OperationType.LaserCut].Code,
                    Name   = OpDefs[OperationType.LaserCut].Name,
                    Number = OpDefs[OperationType.LaserCut].Number,
                    Time   = obj.TimeCut
                };
                if (!string.IsNullOrEmpty(obj.DxfDimensions))
                    op.Parameters.Add(new Parameter { Id = "DxfDimensions", Value = obj.DxfDimensions });
                op.Parameters.Add(new Parameter { Id = "text", Value = "Резка лазером в среде воздуха" });
                if (!string.IsNullOrEmpty(obj.Material))
                {
                    op.Resources.Add(new Resource
                    {
                        Id           = obj.CodeMaterial ?? "",
                        Designation  = "",
                        Name         = obj.Material,
                        ResourceType = "Материал",
                        Quantity     = obj.Mass.ToString("G", CultureInfo.InvariantCulture)
                    });
                    op.Resources.Add(new Resource
                    {
                        Id           = ResDefs[ResourceName.LaserCutOperator].IDTurbo,
                        Designation  = ResDefs[ResourceName.LaserCutOperator].Designation,
                        Name         = ResDefs[ResourceName.LaserCutOperator].Name,
                        ResourceType = ResDefs[ResourceName.LaserCutOperator].ResourceType,
                        Quantity     = ""
                    });
                    var machineName = obj.SheetThickness < 2 ? ResourceName.LaserCutMachine1500 : ResourceName.LaserCutMachine3000;
                    op.Resources.Add(new Resource
                    {
                        Id           = ResDefs[machineName].IDTurbo,
                        Designation  = ResDefs[machineName].Designation,
                        Name         = ResDefs[machineName].Name,
                        ResourceType = ResDefs[machineName].ResourceType,
                        Quantity     = ""
                    });
                }
                yield return op;
            }

            //если есть параметры гибки то появится операция гибки
            if (!string.IsNullOrEmpty(obj.R) || !string.IsNullOrEmpty(obj.V) || !string.IsNullOrEmpty(obj.Q))
            {
                var op = new Operation
                {
                    Id     = OpDefs[OperationType.BendCNC].Code,
                    Name   = OpDefs[OperationType.BendCNC].Name,
                    Number = OpDefs[OperationType.BendCNC].Number,
                    Time   = (double.Parse(obj.Q) * 15).ToString()
                };
                if (!string.IsNullOrEmpty(obj.R)) op.Tools.Add(new Tool { Id = "R", Value = obj.R });
                if (!string.IsNullOrEmpty(obj.V)) op.Tools.Add(new Tool { Id = "V", Value = obj.V });
                if (!string.IsNullOrEmpty(obj.Q)) op.Parameters.AddRange(new[] {
                    new Parameter { Id = "Q",    Value = obj.Q },
                    new Parameter { Id = "text", Value = "Недогиб не допускается" }
                });
                op.Resources.Add(new Resource
                {
                    Id           = ResDefs[ResourceName.BendOperator].IDTurbo,
                    Designation  = ResDefs[ResourceName.BendOperator].Designation,
                    Name         = ResDefs[ResourceName.BendOperator].Name,
                    ResourceType = ResDefs[ResourceName.BendOperator].ResourceType,
                    Quantity     = ""
                });
                op.Resources.Add(new Resource
                {
                    Id           = ResDefs[ResourceName.BendMachine].IDTurbo,
                    Designation  = ResDefs[ResourceName.BendMachine].Designation,
                    Name         = ResDefs[ResourceName.BendMachine].Name,
                    ResourceType = ResDefs[ResourceName.BendMachine].ResourceType,
                    Quantity     = ""
                });
                yield return op;
            }
        }

        private static Operation BuildCoatingOperation(ObjectAssemblyKompas obj)
        {
            if (string.IsNullOrEmpty(obj.Coating))
                return null;

            //если есть назначенное покрытие в свойстве Coating, то появится операция покраски
            var op = new Operation
            {
                Id     = OpDefs[OperationType.Painting].Code,
                Name   = OpDefs[OperationType.Painting].Name,
                Number = OpDefs[OperationType.Painting].Number,
                Time   = ""
            };
            if (!string.IsNullOrEmpty(obj.Area))
                op.Parameters.Add(new Parameter { Id = "Area", Value = obj.Area });
            if (obj.CoverageArea > 0)
                op.Parameters.Add(new Parameter { Id = "CoverageArea", Value = obj.CoverageArea.ToString() });

            string resourceType;
            if (!SectionToResourceType.TryGetValue(obj.SpecificationSection ?? "", out resourceType))
                resourceType = obj.SpecificationSection ?? "";
            //это добавляю саму деталь
            op.Resources.Add(new Resource
            {
                Id = "",
                Designation = obj.Designation ?? "",
                Name = obj.Name ?? "",
                ResourceType = resourceType,
                Quantity = "1"
            });
            //это добавляю название краски
            op.Resources.Add(new Resource
            {
                Id = "153153",
                Designation = "",
                Name = obj.Coating ?? "",
                ResourceType = "Материал",
                Quantity = ""
            });
            //добавляю оператора
            op.Resources.Add(new Resource
            {
                Id = ResDefs[ResourceName.PaintOperator].IDTurbo,
                Designation = ResDefs[ResourceName.PaintOperator].Designation,
                Name = ResDefs[ResourceName.PaintOperator].Name,
                ResourceType = ResDefs[ResourceName.PaintOperator].ResourceType,
                Quantity = ""
            });
            //добавляю покрасочную камеру
            op.Resources.Add(new Resource
            {
                Id = ResDefs[ResourceName.PaintBooth].IDTurbo,
                Designation = ResDefs[ResourceName.PaintBooth].Designation,
                Name = ResDefs[ResourceName.PaintBooth].Name,
                ResourceType = ResDefs[ResourceName.PaintBooth].ResourceType,
                Quantity = ""
            });
            return op;
        }

        private static Operation BuildLocksmithOperation(ObjectAssemblyKompas obj)
        {
            if (string.IsNullOrEmpty(obj.LocksmithWork))
                return null;

            //если есть назначенные слесарные работы то появится операция слесарная
            return new Operation
            {
                Id     = OpDefs[OperationType.Locksmith].Code,
                Name   = OpDefs[OperationType.Locksmith].Name,
                Number = OpDefs[OperationType.Locksmith].Number,
                Time   = obj.LocksmithWork
            };
        }
    }
}
