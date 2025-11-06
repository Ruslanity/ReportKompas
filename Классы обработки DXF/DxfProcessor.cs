using netDxf;
using netDxf.Entities;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace ReportKompas
{
    class DxfProcessor
    {
        DxfDocument dxf;
        SpeedCut speedCut;
        Key matchedKey; //Подходящий ключ из словаря
        // Путь до файла DXF
        public string FilePath { get; set; }

        // Габаритные размеры
        public DimensionsDXF Size { get; private set; }

        // Общая длина линий в DXF (в мм)
        private double totalLengthMm { get; set; }

        //Массив сегментов DXF в виде "ID, начальная точка, конечная точка"
        List<Segment> segments;

        //Массив начальных точек контуров DXF документа
        List<PointKey> contourStartPoints;

        // Время работы лазера, прожига, холостого хода
        protected double LaserOperationTime { get; set; } // Время работы лазера
        protected double EngravingTime { get; set; } // Время прожига
        protected double IdleTime { get; set; } // Время холостого хода

        public DxfProcessor(string filePath)
        {
            FilePath = filePath;
            LaserOperationTime = 0;
            EngravingTime = 0;
            IdleTime = 0;
            Size = new DimensionsDXF();
            if (!File.Exists(FilePath))
                throw new FileNotFoundException("DXF файл не найден", FilePath);
            if (string.IsNullOrEmpty(FilePath))
                throw new FileNotFoundException("Путь не указан", FilePath);
            if (!Path.IsPathRooted(FilePath))
                throw new FileNotFoundException("Файл по указанному пути не существует", FilePath);
            else dxf = DxfDocument.Load(filePath);
            using (Settings settings = Settings.Load(Settings.DefaultPathSettings))
                speedCut = SpeedCut.Load(settings.PathDictionarySpeedCut);
            CalculateDimensions(); //считаю габаритные размеры DXF
            SearchSegmentsAndTotalLength(); //считаю сегменты и общую длину линий в dxf
            SearchMatchedKey(); //ищю в словаре подходящий ключ
            CalculateLaserOperationTime();
            CalculateEngravingTime();
            CalculateIdleTime();            
        }

        // Метод для чтения DXF и вычисления габаритных размеров
        public void CalculateDimensions()
        {
            var dimensions = new DimensionsDXF();
            foreach (var block in dxf.Blocks.Items)
            {
                foreach (var entity in block.Entities)
                {
                    var points = GetEntityPoints(entity);

                    foreach (var point in points)
                    {
                        if (point.X < dimensions.MinX) dimensions.MinX = point.X;
                        if (point.X > dimensions.MaxX) dimensions.MaxX = point.X;

                        if (point.Y < dimensions.MinY) dimensions.MinY = point.Y;
                        if (point.Y > dimensions.MaxY) dimensions.MaxY = point.Y;
                    }
                }
            }
            // Заглушка — здесь должна быть логика разбора DXF файла
            Size.Width = dimensions.MaxX - dimensions.MinX;
            Size.Height = dimensions.MaxY - dimensions.MinY;
        }

        public double TotalCuttingTime // Время лазерной резки (в секундах)
        {
            get { return LaserOperationTime + 10 + EngravingTime + IdleTime; }
        }

        //Рассчет времени работы лазера
        private void CalculateLaserOperationTime()
        {
            // Расчет времени в секундах
            double CuttingSpeed = 10000; // значение по умолчанию

            if (matchedKey != null)
            {
                CuttingSpeed = matchedKey.Value;
            }
            LaserOperationTime = Math.Round(CalculateLaserCuttingTime(totalLengthMm, CuttingSpeed, 10000), 2);
        }

        //Рассчет времени прожига
        private void CalculateEngravingTime()
        {
            // список точек начала контуров
            contourStartPoints = new List<PointKey>();

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
                                        stack.Push(neighborId);
                                }
                            }
                        }
                    }
                    contourCount++;
                    // Если был найден стартовая точка для блока, добавляем ее к списку
                    contourStartPoints.Add(segment.Start);
                }
            }

            // Расчет прожогов
            double factor = 1; // значение по умолчанию

            if (matchedKey != null)
                factor = matchedKey.BurnTime;

            EngravingTime = contourCount * factor;
        }

        //Рассчет времени холостого хода
        private void CalculateIdleTime()
        {
            for (int i = 0; i < contourStartPoints.Count - 1; i++)
            {
                IdleTime += Distance(contourStartPoints[i], contourStartPoints[i + 1]);
            }
            IdleTime = Math.Round((IdleTime / 100000) * 60, 2);
        }

        private static double Distance(Vector3 p1, Vector3 p2)
        {
            return Math.Sqrt(Math.Pow(p2.X - p1.X, 2) +
                             Math.Pow(p2.Y - p1.Y, 2));
        }
        private static double DistanceLwPolyline(Vector2 p1, Vector2 p2)
        {
            return Math.Sqrt(Math.Pow(p2.X - p1.X, 2) +
                             Math.Pow(p2.Y - p1.Y, 2));
        }
        private static double CalculatePolylineLength(LwPolyline polyline)
        {
            double length = 0;
            for (int i = 0; i < polyline.Vertexes.Count - 1; i++)
            {
                var v1 = polyline.Vertexes[i];
                var v2 = polyline.Vertexes[i + 1];
                length += DistanceLwPolyline(v1.Position, v2.Position);
            }
            return length;
        }
        private double CalculateLaserCuttingTime(double lengthMm, double maxSpeedMmPerMin, double accelerationMmPerSec2)
        {
            double maxSpeedMmPerSec = maxSpeedMmPerMin / 60; //перевожу в секунды
            double t_acc = maxSpeedMmPerSec / accelerationMmPerSec2;
            double s_acc = 0.5 * accelerationMmPerSec2 * t_acc * t_acc;

            if (lengthMm > 2 * s_acc)
            {
                double s_const = lengthMm - 2 * s_acc;
                double t_const = s_const / maxSpeedMmPerSec;
                return 2 * t_acc + t_const; // время в секундах
            }
            else
            {
                double v_peak = Math.Sqrt(lengthMm * accelerationMmPerSec2);
                return 2 * v_peak / accelerationMmPerSec2; // время в секундах
            }
        }
        private void SearchMatchedKey()
        {
            if (speedCut != null)
            {
                foreach (var key in speedCut.Keys)
                {
                    string F = $"Key: {key.Name}";
                    if (FilePath.Contains(key.Name))
                    {
                        matchedKey = key;
                        break;
                    }
                }
            }
        }
        private void SearchSegmentsAndTotalLength()
        {
            segments = new List<Segment>();
            int segmentId = 0;
            foreach (var block in dxf.Blocks.Items)
            {
                foreach (var entity in block.Entities)
                {
                    PointKey startPoint = null, endPoint = null;

                    if (entity.Type == EntityType.LwPolyline)
                    {
                        var polyline = entity as LwPolyline;
                        totalLengthMm += CalculatePolylineLength(polyline);

                        // Добавляем сегменты для каждой пары последовательных вершин
                        for (int i = 0; i < polyline.Vertexes.Count - 1; i++)
                        {
                            var vStart = polyline.Vertexes[i];
                            var vEnd = polyline.Vertexes[i + 1];

                            startPoint = new PointKey(vStart.Position.X, vStart.Position.Y);
                            endPoint = new PointKey(vEnd.Position.X, vEnd.Position.Y);

                            segments.Add(new Segment(segmentId++, startPoint, endPoint));
                        }
                        // Если полилиния замкнута, добавляем сегмент из последней вершины в первую
                        if (polyline.IsClosed && polyline.Vertexes.Count > 1)
                        {
                            var vLast = polyline.Vertexes[polyline.Vertexes.Count - 1];
                            var vFirst = polyline.Vertexes[0];

                            startPoint = new PointKey(vLast.Position.X, vLast.Position.Y);
                            endPoint = new PointKey(vFirst.Position.X, vFirst.Position.Y);

                            segments.Add(new Segment(segmentId++, startPoint, endPoint));
                        }
                    }
                    if (entity.Type == EntityType.Line)
                    {
                        var line = entity as Line;
                        totalLengthMm += Distance(line.StartPoint, line.EndPoint);

                        startPoint = new PointKey(line.StartPoint.X, line.StartPoint.Y);
                        endPoint = new PointKey(line.EndPoint.X, line.EndPoint.Y);
                        segments.Add(new Segment(segmentId++, startPoint, endPoint));
                    }
                    if (entity.Type == EntityType.Circle)
                    {
                        var circle = entity as Circle;
                        totalLengthMm += 2 * Math.PI * circle.Radius;

                        int parts = 36; // число сегментов для аппроксимации окружности
                        double startAngle = 0;
                        double endAngle = 2 * Math.PI; // полная окружность
                        double angleStep = (endAngle - startAngle) / parts;

                        //PointKey prevPoint = null;
                        PointKey prevPointCircle = null;

                        for (int i = 0; i <= parts; i++)
                        {
                            double angle = startAngle + i * angleStep;
                            double x = circle.Center.X + circle.Radius * Math.Cos(angle);
                            double y = circle.Center.Y + circle.Radius * Math.Sin(angle);
                            var currentPoint = new PointKey(x, y);
                            if (i > 0)
                                segments.Add(new Segment(segmentId++, prevPointCircle, currentPoint));
                            if (i == 0)
                                startPoint = currentPoint;
                            prevPointCircle = currentPoint;
                        }
                        endPoint = prevPointCircle;
                    }
                    if (entity.Type == EntityType.Arc)
                    {
                        var arc = entity as Arc;
                        double angle1 = Math.Abs(arc.EndAngle - arc.StartAngle);
                        double arcLength = arc.Radius * (((angle1 > 180 ? 360 - angle1 : angle1) * Math.PI) / 180);
                        totalLengthMm += arcLength;

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
                                segments.Add(new Segment(segmentId++, prevPoint, currentPoint));
                            if (i == 0)
                                startPoint = currentPoint; // первая точка дуги — старт контура
                            prevPoint = currentPoint;
                        }
                        endPoint = prevPoint; // конец дуги
                    }
                }
            }
        }
        private static double Distance(PointKey p1, PointKey p2)
        {
            double deltaX = p2.X - p1.X;
            double deltaY = p2.Y - p1.Y;
            return Math.Sqrt(deltaX * deltaX + deltaY * deltaY);
        }
        private System.Collections.Generic.IEnumerable<netDxf.Vector3> GetEntityPoints(EntityObject entity)
        {
            switch (entity.Type)
            {
                case EntityType.Line:
                    var line = entity as Line;
                    yield return line.StartPoint;
                    yield return line.EndPoint;
                    break;

                case EntityType.Circle:
                    var circle = entity as Circle;
                    var center = circle.Center;
                    var radius = circle.Radius;

                    yield return new netDxf.Vector3(center.X - radius, center.Y - radius, center.Z);
                    yield return new netDxf.Vector3(center.X + radius, center.Y + radius, center.Z);
                    break;

                case EntityType.Polyline:
                    var polyline = entity as Polyline;
                    foreach (var vertex in polyline.Vertexes)
                        yield return vertex.Position;
                    break;

                // Можно добавить обработку других типов по необходимости

                default:
                    // Для неподдерживаемых типов ничего не делаем
                    break;
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
    }
}
