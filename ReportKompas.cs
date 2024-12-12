using Kompas6API5;
using KompasAPI7;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;


namespace ReportKompas
{
    /// <summary> Пример взял тут
    /// https://allineed.ru/development/dotnet-development/charp-development/80-csharp-working-with-datagridview
    /// </summary>
    public partial class ReportKompas : Form
    {
        IApplication application;
        IKompasDocument3D document3D;
        KompasObject kompas;
        ksDocument3D ksDocument3D;
        List<ObjectAssemblyKompas> objectsAssemblyKompas;
        private bool cancelContextMenu = false;
        public ReportKompas()
        {
            InitializeComponent();
        }

        private void DisassembleObject(IPart7 part7, string Name)
        {
            ObjectAssemblyKompas objectAssemblyKompas = new ObjectAssemblyKompas();
            IPropertyMng propertyMng = (IPropertyMng)application;
            var properties = propertyMng.GetProperties(document3D);
            IPropertyKeeper propertyKeeper = (IPropertyKeeper)part7;
            foreach (IProperty item in properties)
            {
                if (item.Name == "Наименование")
                {
                    dynamic info;
                    bool source;
                    propertyKeeper.GetPropertyValue((_Property)item, out info, false, out source);
                    objectAssemblyKompas.Name = info;
                }
                if (item.Name == "Обозначение")
                {
                    dynamic info;
                    bool source;
                    propertyKeeper.GetPropertyValue((_Property)item, out info, false, out source);
                    objectAssemblyKompas.Designation = info;
                }
                if (item.Name == "Материал")
                {
                    dynamic info;
                    bool source;
                    propertyKeeper.GetPropertyValue((_Property)item, out info, false, out source);
                    objectAssemblyKompas.Material = info;
                }
                if (item.Name == "Раздел спецификации")
                {
                    dynamic info;
                    bool source;
                    propertyKeeper.GetPropertyValue((_Property)item, out info, false, out source);
                    objectAssemblyKompas.SpecificationSection = info;
                }
                if (item.Name == "Масса")
                {
                    dynamic info;
                    bool source;
                    propertyKeeper.GetPropertyValue((_Property)item, out info, false, out source);
                    objectAssemblyKompas.Mass = Math.Round(info,2);
                }
                if (item.Name == "Покрытие")
                {
                    dynamic info;
                    bool source;
                    propertyKeeper.GetPropertyValue((_Property)item, out info, false, out source);
                    objectAssemblyKompas.Coating = info;
                }
            }
            if(Name!="0")
            {
                objectAssemblyKompas.Parent = Name;
            }
            ObjectAssemblyKompas index = objectsAssemblyKompas.Find((ObjectAssemblyKompas)=>
            ObjectAssemblyKompas.Designation == objectAssemblyKompas.Designation &&
            ObjectAssemblyKompas.Name == objectAssemblyKompas.Name &&
            ObjectAssemblyKompas.Parent == objectAssemblyKompas.Parent);
            {
                if (index!=null)
                {
                    index.Quantity++;
                }
                else if (index == null)
                {
                    objectAssemblyKompas.Quantity++;
                    objectsAssemblyKompas.Add(objectAssemblyKompas);
                }
            }
        }

        private void FillTable()
        {
            dataGridView1.Rows.Clear();
            dataGridView1.DataSource = objectsAssemblyKompas;
            dataGridView1.Columns["Designation"].HeaderText = "Обозначение";
            dataGridView1.Columns["Name"].HeaderText = "Наименование";
            dataGridView1.Columns["Quantity"].HeaderText = "Кол-во";
            dataGridView1.Columns["Material"].HeaderText = "Материал";
            dataGridView1.Columns["SpecificationSection"].HeaderText = "Раздел спецификации";
            dataGridView1.Columns["Mass"].HeaderText = "Масса";
            //dataGridView1.Columns["Coating"].HeaderText = "Покрытие";
            dataGridView1.Columns["Parent"].HeaderText = "Куда входит";
            dataGridView1.RowHeadersVisible = false;
            dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            dataGridView1.Columns["Designation"].Width = 200;
            dataGridView1.Columns["Name"].Width = 200;
            dataGridView1.Columns["Quantity"].Width = 50;
            dataGridView1.Columns["Material"].Width = 150;
            dataGridView1.Columns["SpecificationSection"].Width = 200;
            dataGridView1.Columns["Mass"].Width = 50;
            //dataGridView1.Columns["Coating"].Width = 100;
            dataGridView1.Columns["Parent"].Width = 200;
        }
        private void toolStripButton1_Click(object sender, EventArgs e)
        {
            objectsAssemblyKompas = new List<ObjectAssemblyKompas>();

            kompas = (KompasObject)Marshal.GetActiveObject("KOMPAS.Application.5");            
            kompas.Visible = true;
            kompas.ActivateControllerAPI();
            ksDocument3D = (ksDocument3D)kompas.ActiveDocument3D();

            document3D = kompas.TransferInterface(ksDocument3D, 2, 0);
            application = kompas.ksGetApplication7();

            ksPartCollection _ksPartCollection = ksDocument3D.PartCollection(true);
            
            IPart7 part7 = document3D.TopPart;
            switch (document3D.DocumentType)
            {
                case Kompas6Constants.DocumentTypeEnum.ksDocumentUnknown:
                    break;
                case Kompas6Constants.DocumentTypeEnum.ksDocumentDrawing:
                    break;
                case Kompas6Constants.DocumentTypeEnum.ksDocumentFragment:
                    break;
                case Kompas6Constants.DocumentTypeEnum.ksDocumentSpecification:
                    break;
                case Kompas6Constants.DocumentTypeEnum.ksDocumentPart:
                    {
                        //FillTable(application);
                        break;
                    }
                case Kompas6Constants.DocumentTypeEnum.ksDocumentAssembly:
                    {
                        DisassembleObject(part7, "0");
                        for (int i = 0; i < _ksPartCollection.GetCount(); i++)
                        {
                            ksPart ksPart = _ksPartCollection.GetByIndex(i);
                            if (ksPart.excluded != true)
                            {
                                IPart7 _part7 = kompas.TransferInterface(ksPart, 2, 0);
                                DisassembleObject(_part7, part7.Marking+" - "+part7.Name);
                            }
                        }
                        FillTable();
                        break;
                    }
                case Kompas6Constants.DocumentTypeEnum.ksDocumentTextual:
                    break;
                case Kompas6Constants.DocumentTypeEnum.ksDocumentTechnologyAssembly:
                    break;
                default:
                    {
                        break;
                    }
            }
        }
    }
}
