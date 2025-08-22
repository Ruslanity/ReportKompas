using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml;
using netDxf;
using netDxf.Entities;

namespace ReportKompas
{
    class DXFCountBurnPoint
    {
        DxfDocument dxf;
        public double burnPoint;
        //private Settings _settings;
        private Dictionary<string, double> DictionarySpeedBurn;

        class PointKey
        {
            public double X { get; }
            public double Y { get; }

            public PointKey(double x, double y)
            {
                X = Math.Round(x, 6); // округление для избегания ошибок точности
                Y = Math.Round(y, 6);
            }

            public override bool Equals(object obj)
            {
                if (obj is PointKey other)
                {
                    return X == other.X && Y == other.Y;
                }
                return false;
            }

            public override int GetHashCode()
            {
                return X.GetHashCode() ^ Y.GetHashCode();
            }
        }

        // Связь между сегментами по концам
        class Segment
        {
            public int Id { get; }
            public PointKey Start { get; }
            public PointKey End { get; }

            public Segment(int id, PointKey start, PointKey end)
            {
                Id = id;
                Start = start;
                End = end;
            }
        }

        void LoadSpeedBurn()
        {
            DictionarySpeedBurn = new Dictionary<string, double>();
            XmlDocument xmlDoc = new XmlDocument();

            using (var settings = new Settings())
            {
                // Загружаем XML файл
                string filePath = Path.Combine(settings?.Speed_Cut_textBox.Text, "SpeedCut.xml");
                xmlDoc.Load(filePath);

                // Получаем все узлы <Key>
                XmlNodeList keyNodes = xmlDoc.SelectNodes("/Dictionary/Key");

                foreach (XmlNode keyNode in keyNodes)
                {
                    // Получаем атрибут "name"
                    var nameAttr = keyNode.Attributes["name"];
                    if (nameAttr == null)
                        continue; // пропускаем, если нет атрибута name

                    string key = nameAttr.Value;

                    // Получаем атрибут "burntime"
                    var burntimeAttr = keyNode.Attributes["burntime"];
                    if (burntimeAttr == null)
                        continue; // пропускаем, если нет burntime

                    string burntimeStr = burntimeAttr.Value;

                    // Преобразуем в double
                    if (double.TryParse(burntimeStr, out double burntime))
                    {
                        DictionarySpeedBurn[key] = burntime;
                    }
                    else
                    {
                        DictionarySpeedBurn[key] = 0.8;
                    }
                }
            }
        }

        public DXFCountBurnPoint(string dxfFilePath)
        {
            dxf = DxfDocument.Load(dxfFilePath);
            var segments = new List<Segment>();
            int segmentId = 0;

            foreach (var block in dxf.Blocks.Items)
            {
                foreach (var entity in block.Entities)
                {
                    if (entity.Type == EntityType.Line)
                    {
                        var line = entity as Line;
                        var startPoint = new PointKey(line.StartPoint.X, line.StartPoint.Y);
                        var endPoint = new PointKey(line.EndPoint.X, line.EndPoint.Y);
                        segments.Add(new Segment(segmentId++, startPoint, endPoint));
                    }
                    if (entity.Type == EntityType.Arc)
                    {
                        var arc = entity as Arc;
                        // Для дуги можно разбить на несколько линий или оставить как есть
                        // Для простоты возьмем дугу как один сегмент с началом и концом по углам
                        // Но лучше разбивать дугу на мелкие отрезки для точности
                        // Ниже пример разбиения дуги на несколько линий

                        int parts = 2; // число частей для аппроксимации дуги
                        double startAngle = arc.StartAngle * Math.PI / 180.0;
                        double endAngle = arc.EndAngle * Math.PI / 180.0;

                        if (endAngle < startAngle)
                            endAngle += 2 * Math.PI;

                        double angleStep = (endAngle - startAngle) / parts;

                        PointKey prevPoint = null;

                        for (int i = 0; i <= parts; i++)
                        {
                            double angle = startAngle + i * angleStep;
                            double x = arc.Center.X + arc.Radius * Math.Cos(angle);
                            double y = arc.Center.Y + arc.Radius * Math.Sin(angle);
                            var currentPoint = new PointKey(x, y);

                            if (i > 0)
                            {
                                segments.Add(new Segment(segmentId++, prevPoint, currentPoint));
                            }
                            prevPoint = currentPoint;
                        }
                    }
                    // Внутри цикла foreach, после обработки линий и дуг
                    if (entity.Type == EntityType.Circle)
                    {
                        var circle = entity as netDxf.Entities.Circle;
                        int parts = 36; // число сегментов для аппроксимации окружности
                        double startAngle = 0;
                        double endAngle = 2 * Math.PI; // полная окружность

                        double angleStep = (endAngle - startAngle) / parts;

                        PointKey prevPoint = null;

                        for (int i = 0; i <= parts; i++)
                        {
                            double angle = startAngle + i * angleStep;
                            double x = circle.Center.X + circle.Radius * Math.Cos(angle);
                            double y = circle.Center.Y + circle.Radius * Math.Sin(angle);
                            var currentPoint = new PointKey(x, y);

                            if (i > 0)
                            {
                                segments.Add(new Segment(segmentId++, prevPoint, currentPoint));
                            }
                            prevPoint = currentPoint;
                        }
                    }
                }
                // Можно добавить обработку других типов кривых при необходимости
            }
            // Построение графа связных сегментов по концам
            var adjacencyList = new Dictionary<PointKey, List<int>>(); // точка -> список сегментов

            foreach (var segment in segments)
            {
                if (!adjacencyList.ContainsKey(segment.Start))
                    adjacencyList[segment.Start] = new List<int>();
                adjacencyList[segment.Start].Add(segment.Id);

                if (!adjacencyList.ContainsKey(segment.End))
                    adjacencyList[segment.End] = new List<int>();
                adjacencyList[segment.End].Add(segment.Id);
            }
            // Поиск связных компонент с помощью обхода в глубину
            var visitedSegments = new HashSet<int>();
            int contourCount = 0;

            foreach (var segment in segments)
            {
                if (!visitedSegments.Contains(segment.Id))
                {
                    // Начинаем обход с этого сегмента
                    var stack = new Stack<int>();
                    stack.Push(segment.Id);

                    while (stack.Count > 0)
                    {
                        int currentId = stack.Pop();
                        if (!visitedSegments.Contains(currentId))
                        {
                            visitedSegments.Add(currentId);
                            var currentSegment = segments[currentId];

                            // Находим соседние сегменты по концам
                            foreach (var point in new[] { currentSegment.Start, currentSegment.End })
                            {
                                foreach (var neighborId in adjacencyList[point])
                                {
                                    if (!visitedSegments.Contains(neighborId))
                                    {
                                        stack.Push(neighborId);
                                    }
                                }
                            }
                        }
                    }
                    contourCount++;
                }
            }

            double factor = 1; // значение по умолчанию

            LoadSpeedBurn();
            foreach (var key in DictionarySpeedBurn.Keys)
            {
                if (dxfFilePath.Contains(key))
                {
                    factor = DictionarySpeedBurn[key];
                    break; // если нашли подходящий ключ, можно выйти из цикла
                }
            }

            burnPoint = contourCount * factor;
        }
    }
}
