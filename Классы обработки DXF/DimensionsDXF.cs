namespace ReportKompas
{
    class DimensionsDXF
    {
        public double Width { get; set; }
        public double Height { get; set; }
        public double MinX { get; set; }
        public double MaxX { get; set; }
        public double MinY { get; set; }
        public double MaxY { get; set; }

        public DimensionsDXF()
        {
            Width = 0;
            Height = 0;
        }

        public DimensionsDXF(double width, double height)
        {
            Width = width;
            Height = height;
        }
    }
}
