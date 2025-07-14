using netDxf;
using netDxf.Entities;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace ReportKompas
{
    public class DxfHelper
    {
        public static DxfDimensions GetDxfDimensions(string dxfFilePath)
        {
            if (!File.Exists(dxfFilePath))
                throw new FileNotFoundException("Файл не найден", dxfFilePath);

            // Загружаем DXF файл
            DxfDocument dxf = DxfDocument.Load(dxfFilePath);
            if (dxf == null)
                throw new Exception("Не удалось загрузить DXF файл");

            var dimensions = new DxfDimensions();

            // Обрабатываем все сущности в файле
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

                        if (point.Z < dimensions.MinZ) dimensions.MinZ = point.Z;
                        if (point.Z > dimensions.MaxZ) dimensions.MaxZ = point.Z;
                    }
                }                
            }
            return dimensions;
        }

        private static System.Collections.Generic.IEnumerable<netDxf.Vector3> GetEntityPoints(EntityObject entity)
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

                //case EntityType.Arc:
                //    var arc = entity as Arc;
                //    yield return arc.StartPoint;
                //    yield return arc.EndPoint;
                //    break;

                case EntityType.Polyline:
                    var polyline = entity as Polyline;
                    foreach (var vertex in polyline.Vertexes)
                        yield return vertex.Position;
                    break;

                //case EntityType.LwPolyline:
                //    var lwPolyline = entity as LwPolyline;
                //    foreach (var vertex in lwPolyline.Vertexes)
                //        yield return vertex.Position;
                //    break;

                //case EntityType.Ellipse:
                //    var ellipse = entity as Ellipse;
                //    var eCenter = ellipse.Center;

                //    // Приблизительные габариты по центру и радиусам
                //    var eMajorAxisLength = ellipse.MajorAxis.Length * ellipse.Ratio;
                //    var eMinorAxisLength = ellipse.MinorAxis.Length;

                //    yield return new netDxf.Vector3(eCenter.X - eMajorAxisLength / 2, eCenter.Y - eMinorAxisLength / 2, eCenter.Z);
                //    yield return new netDxf.Vector3(eCenter.X + eMajorAxisLength / 2, eCenter.Y + eMinorAxisLength / 2, eCenter.Z);
                //    break;

                // Можно добавить обработку других типов по необходимости

                default:
                    // Для неподдерживаемых типов ничего не делаем
                    break;
            }
        }
    }
}
