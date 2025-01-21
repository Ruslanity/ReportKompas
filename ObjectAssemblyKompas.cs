﻿using System;
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
        public string Coating { get; set; }
        public string Parent { get; set; }
        public string Bending { get; set; }
        public string FullName { get; set; }
        public string PathToDXF { get; set; }
        public string OverallDimensions { get; set; }

        public ObjectAssemblyKompas() { }        
        public ObjectAssemblyKompas(string designation,
                                    string name,
                                    int quantity,
                                    string specificationSection,
                                    string material,
                                    double mass,
                                    string coating, //покрытие(краска)
                                    string parent,
                                    string bending, //гибки(кол-во гибов, указание инструмента для гибки)
                                    string fullName,
                                    string pathToDXF,
                                    string overallDimensions)
        {
            Designation = designation;
            Name = name;
            Quantity = quantity;
            SpecificationSection = specificationSection;
            Material = material;
            Mass = mass;
            Coating = coating;
            Parent = parent;
            Bending = bending;
            FullName = fullName;
            PathToDXF = pathToDXF;
            OverallDimensions = overallDimensions;
        }
    }
}
