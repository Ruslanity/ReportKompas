using KompasAPI7;
using System;
using System.Globalization;
using System.IO;

namespace ReportKompas
{
    /// <summary>
    /// Анализатор листового металла: расчёт параметров гибки (R, V, Q) и поиск DXF файлов
    /// </summary>
    public static class SheetMetalAnalyzer
    {
        /// <summary>
        /// Анализирует деталь из листового металла и заполняет параметры гибки в объекте
        /// </summary>
        /// <param name="part">Деталь KOMPAS</param>
        /// <param name="obj">Объект для заполнения параметров</param>
        public static void Analyze(IPart7 part, ObjectAssemblyKompas obj)
        {
            if (part == null || obj == null)
                return;

            try
            {
                ISheetMetalContainer container = part as ISheetMetalContainer;
                if (container == null)
                    return;

                if (obj.Children != null && obj.Children.Count > 0)
                {
                    obj.CoilBatch = obj.Designation + " СБ" + " - " + obj.Name;
                }

                // Обработка листового тела
                ProcessSheetMetalBody(container, part, obj);

                // Обработка преобразованных в листовое тело
                ProcessConvertedSheetMetal(container, part, obj);
            }
            catch (ArgumentException)
            {
                obj.PathToDXF = "";
            }
        }

        /// <summary>
        /// Обрабатывает стандартное листовое тело
        /// </summary>
        private static void ProcessSheetMetalBody(ISheetMetalContainer container, IPart7 part, ObjectAssemblyKompas obj)
        {
            ISheetMetalBodies bodies = container.SheetMetalBodies;
            ISheetMetalBody body = bodies.SheetMetalBody[0];

            if (body == null)
                return;

            double thickness = body.Thickness;

            // Подсчёт сгибов
            int bendCount = CountBodyBends(body) +
                            CountBendsInCollection(container.SheetMetalBends) +
                            CountBendsInCollection(container.SheetMetalSketchBends) +
                            CountBendsInCollection(container.SheetMetalLineBends);

            FillBendingParams(obj, thickness, bendCount);
            FillCoilBatch(obj, thickness);
            obj.PathToDXF = FindDxfPath(part.FileName, thickness, part.Marking);
        }

        /// <summary>
        /// Обрабатывает тела, преобразованные в листовой металл
        /// </summary>
        private static void ProcessConvertedSheetMetal(ISheetMetalContainer container, IPart7 part, ObjectAssemblyKompas obj)
        {
            IConvertsToSheetMetals converts = container.ConvertsToSheetMetals;
            if (converts.Count == 0)
                return;

            IConvertToSheetMetal convert = converts.ConvertToSheetMetal[0];
            double thickness = convert.Thickness;

            obj.R = GetBendRadius(thickness).ToString();
            obj.V = GetDieGrooveWidth(thickness);
            FillCoilBatch(obj, thickness);

            if (convert.BendsCount > 0)
            {
                obj.Q = convert.BendsCount.ToString();
                obj.PathToDXF = FindDxfPath(part.FileName, thickness, part.Marking);
            }
        }

        /// <summary>
        /// Заполняет атрибут CoilBatch (бухта) на основе наличия детей и обозначения
        /// </summary>
        private static void FillCoilBatch(ObjectAssemblyKompas obj, double thickness)
        {
            if (string.IsNullOrEmpty(obj.Designation))
                return;

            if (obj.Children != null && obj.Children.Count > 0)
            {
                // Если есть дети - используем полное обозначение
                obj.CoilBatch = obj.Designation + " - " + obj.Name;
            }
            else if (obj.Designation.StartsWith("АЛ.", StringComparison.OrdinalIgnoreCase) && obj.Designation.Length > 3)
            {
                // Если детей нет и обозначение начинается на "АЛ." - формируем имя с толщиной
                string thicknessStr = thickness.ToString(CultureInfo.CreateSpecificCulture("en-GB"));
                string markingSuffix = obj.Designation.Remove(0, 3);
                obj.CoilBatch = $"{thicknessStr}mm_{markingSuffix}";
            }
            else
            {
                // Если детей нет и обозначение НЕ начинается на "АЛ." - записываем Designation как есть
                obj.CoilBatch = obj.Designation;
            }
        }

        /// <summary>
        /// Подсчитывает сгибы в листовом теле
        /// </summary>
        private static int CountBodyBends(ISheetMetalBody body)
        {
            IFeature7 feature = body.Owner;
            Object[] subFeatures = feature.SubFeatures[0, false, false];
            return subFeatures?.Length ?? 0;
        }

        /// <summary>
        /// Подсчитывает количество сгибов в коллекции
        /// </summary>
        private static int CountBendsInCollection(dynamic bendCollection)
        {
            int count = 0;
            for (int i = 0; i < bendCollection.Count; i++)
            {
                IFeature7 feature = (IFeature7)bendCollection[i];
                if (!feature.Excluded)
                {
                    Object[] subFeatures = feature.SubFeatures[0, false, false];
                    if (subFeatures != null)
                        count += subFeatures.Length;
                }
            }
            return count;
        }

        /// <summary>
        /// Заполняет параметры гибки в объекте
        /// </summary>
        private static void FillBendingParams(ObjectAssemblyKompas obj, double thickness, int bendCount)
        {
            if (bendCount <= 0)
                return;

            obj.R = GetBendRadius(thickness).ToString();
            obj.V = GetDieGrooveWidth(thickness);
            obj.Q = bendCount.ToString();
        }

        /// <summary>
        /// Определяет радиус гибки по толщине металла
        /// </summary>
        private static double GetBendRadius(double thickness)
        {
            if (thickness < 3) return 1;
            if (thickness < 6) return 3;
            if (thickness < 11) return 6;
            return 0;
        }

        /// <summary>
        /// Определяет ширину канавки матрицы (V) по толщине металла
        /// </summary>
        private static string GetDieGrooveWidth(double thickness)
        {
            switch (thickness)
            {
                case 0.7:
                case 0.8:
                case 1:
                case 1.2:
                    return "8";
                case 1.5:
                case 2:
                    return "12";
                case 3:
                    return "16";
                case 4:
                    return "30";
                case 5:
                    return "35";
                case 6:
                case 8:
                    return "50";
                case 10:
                    return "80";
                default:
                    return "";
            }
        }

        /// <summary>
        /// Ищет DXF файл и возвращает путь к директории
        /// </summary>
        private static string FindDxfPath(string partFileName, double thickness, string marking)
        {
            if (string.IsNullOrEmpty(marking) || marking.Length < 4)
                return null;

            FileInfo fi = new FileInfo(partFileName);
            string thicknessWithDot = thickness.ToString(CultureInfo.CreateSpecificCulture("en-GB"));
            string thicknessWithComma = thicknessWithDot.Replace('.', ',');
            string markingSuffix = marking.Remove(0, 3);

            string[] fileNames =
            {
                $"{thicknessWithDot}mm_{markingSuffix}.dxf",
                $"{thicknessWithComma}mm_{markingSuffix}.dxf"
            };

            foreach (string fileName in fileNames)
            {
                if (File.Exists(Path.Combine(fi.DirectoryName, fileName)))
                    return fi.DirectoryName;
            }

            return null;
        }
    }
}
