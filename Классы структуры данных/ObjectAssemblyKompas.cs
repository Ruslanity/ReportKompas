using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReportKompas
{
    public class ObjectAssemblyKompas
    {
        public string Designation { get; set; }
        public string Name { get; set; }
        public int Quantity { get; set; }
        public string SpecificationSection { get; set; }
        public string Material { get; set; }
        public double Mass { get; set; }
        public string R { get; set; }
        public string V { get; set; }
        public string Q { get; set; }
        public string Parent { get; set; }
        public string TopParent { get; set; }
        //public string Bending { get; set; }
        public string FullName { get; set; }
        public string PathToDXF { get; set; }
        public string OverallDimensions { get; set; }
        public string Coating { get; set; }
        public string Welding { get; set; }
        public string LocksmithWork { get; set; }
        public string Note { get; set; }
        public string Area { get; set; }
        public string CodeEquipment { get; set; }
        public string CodeMaterial { get; set; }
        public string TimeCut { get; set; }
        public string DxfDimensions { get; set; }
        public bool IsLocal { get; set; }
        public string IsPainted { get; set; }
        public double CoverageArea { get; set; }
        public byte[] PreviewImage { get; set; }  // PNG формат по умолчанию
        public string TechnologicalRoute { get; set; }

        // Связи в дереве
        public ObjectAssemblyKompas ParentK { get; set; }
        public List<ObjectAssemblyKompas> Children { get; private set; }

        public ObjectAssemblyKompas() { }

        // Добавить ребёнка
        public void AddChild(ObjectAssemblyKompas child)
        {
            if (child == null)
                throw new ArgumentNullException(nameof(child));

            // Проверка, инициализирован ли список Children
            if (Children == null)
            {
                Children = new List<ObjectAssemblyKompas>();
            }

            // Поиск в коллекции по совпадению поля Name и Designation
            var existingChild = Children.FirstOrDefault(c =>
                c.Name == child.Name && c.Designation == child.Designation);

            if (existingChild != null)
            {
                // Если объект найден, увеличиваем Quantity
                existingChild.Quantity += 1;
                // Обновляем, возможно, остальные поля, если нужно дополнять их
                // например, Mass, SpecificationSection и т.д.
                // Но здесь предполагается, что только Quantity увеличивается.
            }
            else
            {
                // Если не найден, добавляем новый объект
                child.ParentK = this;
                Children.Add(child);
            }
        }

        // Удалить ребёнка
        public bool RemoveChild(ObjectAssemblyKompas child)
        {
            if (child == null)
                return false;

            if (Children.Remove(child))
            {
                child.ParentK = null;
                return true;
            }
            return false;
        }
        // Найти узел по имени (рекурсивно)
        public ObjectAssemblyKompas FindChild(string designation = null, string name = null)
        {
            // Проверяем текущий узел
            bool matchesDesignation = string.IsNullOrEmpty(designation) || (Designation != null && Designation.Contains(designation));
            bool matchesName = string.IsNullOrEmpty(name) || (Name != null && Name.Contains(name));

            //if (matchesDesignation && matchesName)
            //{
            //    return this;
            //}
            // Рекурсивно ищем в дочерних узлах
            if (Children != null)
            {
                foreach (var child in Children)
                {
                    var found = child.FindChild(designation, name);
                    if (found != null)
                        return found; // Немедленный возврат при первом совпадении
                }
            }

            return null; // Если ничего не найдено
        }

        //Сортировка детей
        public void SortChildrenBySpecificationSection()
        {
            if (Children == null || !Children.Any())
                return;

            Children = Children
                .OrderBy(c =>
                {
            // Назначение приоритета, для сортировки групп
            switch (c.SpecificationSection)
                    {
                        case "Сборочные единицы":
                            return 1;
                        case "Детали":
                            return 2;
                        case "Стандартные изделия":
                            return 3;
                        case "Прочие изделия":
                            return 4;
                        default:
                            return 5; // Для элементов со SpecificationSection == "" или любыми прочими
            }
                })
                .ThenBy(c =>
                {
            // Внутригрупповая сортировка
            switch (c.SpecificationSection)
                    {
                        case "Сборочные единицы":
                        case "Детали":
                            return c.Designation ?? "";
                        case "Стандартные изделия":
                        case "Прочие изделия":
                            return c.Name ?? "";
                        default:
                            return c.SpecificationSection ?? "";
                    }
                })
                .ToList();
        }

        public void SortTreeNodes(ObjectAssemblyKompas node)
        {
            if (node == null)
                return;

            // Если у узла есть дети, сортируем их
            if (node.Children != null && node.Children.Any())
            {
                node.SortChildrenBySpecificationSection(); // вызываем сортировку для текущего уровня
                foreach (var child in node.Children)
                {
                    // рекурсивно продолжаем обход
                    SortTreeNodes(child);
                }
            }
        }

        // Метод для обработки полей Material у объектов с определенными условиями
        public void ReplaceMaterial()
        {
            ProcessDetailsMaterialRecursive(this);
        }

        // Внутренний рекурсивный метод
        private void ProcessDetailsMaterialRecursive(ObjectAssemblyKompas node)
        {
            if (node == null)
                return;

            // Проверка условий
            if (node.SpecificationSection == "Детали" && string.IsNullOrEmpty(node.Material))
            {
                if (!string.IsNullOrEmpty(node.Designation))
                {
                    if (node.Designation.Contains("1.5mm_") && node.Designation.Contains("_Zn"))
                    {
                        node.Material = "Лист ОЦd1,5 ГОСТ 19904-90;08пс ГОСТ 14918-80";
                    }
                    if (node.Designation.Contains("1.5mm_") && node.Designation.Contains("_Aisi"))
                    {
                        node.Material = @"Лист нерж. 1,5ммх1250х2500 AISI430 4N+PE";
                    }
                    if (node.Designation.Contains("1.5mm_") && node.Designation.Contains("_Aisi Bronze"))
                    {
                        node.Material = @"Лист нерж. 1,5ммх1250х2500 AISI430 4N+PE (Bronze)";
                    }
                    if (node.Designation.Contains("1mm_") && node.Designation.Contains("_Zn"))
                    {
                        node.Material = @"Лист ОЦd1,0 ГОСТ 19904-90;08пс ГОСТ 14918-80";
                    }
                    if (node.Designation.Contains("1.5mm_") && node.Designation.Contains("_AL"))
                    {
                        node.Material = @"Лист квинтет 1,5mm AL";
                    }
                    if (node.Designation.Contains("1.5mm_") && node.Designation.Contains("_Forbo"))
                    {
                        node.Material = @"Forbo flooring";
                    }
                }
            }

            // Обработка детей рекурсиво
            if (node.Children != null && node.Children.Any())
            {
                foreach (var child in node.Children)
                {
                    ProcessDetailsMaterialRecursive(child);
                }
            }
        }
    }
}
