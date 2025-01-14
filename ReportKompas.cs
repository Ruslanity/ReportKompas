using ClosedXML.Excel;
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
        BindingList<ObjectAssemblyKompas> sortedListObjects;
        private bool cancelContextMenu = false;
        string fileName;
        public ReportKompas()
        {
            InitializeComponent();
        }
        private void Recursion(IPart7 Part, string ParentName)
        {
            DisassembleObject(Part, ParentName);
            foreach (IPart7 item in Part.Parts)
            {
                if (item.Detail == true) DisassembleObject(item, Part.Marking + " - " + Part.Name);
                if (item.Detail == false) Recursion(item, Part.Marking + " - " + Part.Name);
            }
        }
        private void GetSortedObjectsKompas()
        {
            if (sortedListObjects != null)
            {
                sortedListObjects.Clear();
            }
            else
            {
                sortedListObjects = new BindingList<ObjectAssemblyKompas>();
            }
            ObjectAssemblyKompas kompasObject = objectsAssemblyKompas.SingleOrDefault((ObjectAssemblyKompas) => ObjectAssemblyKompas.Parent == null);
            sortedListObjects.Add(kompasObject);
            void RecursionMethod(string Parent)
            {
                var kOSpecificationSection = objectsAssemblyKompas.FindAll((ObjectAssemblyKompas) => ObjectAssemblyKompas.Parent == Parent &&
                                                                                                      ObjectAssemblyKompas.SpecificationSection == "Сборочные единицы");
                List<ObjectAssemblyKompas> kOSpecificationSectionSorted = (List<ObjectAssemblyKompas>)kOSpecificationSection;
                kOSpecificationSectionSorted.Sort(delegate (ObjectAssemblyKompas x, ObjectAssemblyKompas y)
                { return x.Designation.CompareTo(y.Designation); });
                foreach (ObjectAssemblyKompas item in kOSpecificationSectionSorted)
                {
                    sortedListObjects.Add(item);
                    RecursionMethod(item.Designation + " - " + item.Name);
                }
                var kOSpecificationSection2 = objectsAssemblyKompas.FindAll((ObjectAssemblyKompas) => ObjectAssemblyKompas.Parent == Parent &&
                                                                                                      ObjectAssemblyKompas.SpecificationSection == "Детали");
                List<ObjectAssemblyKompas> kOSpecificationSectionSorted2 = (List<ObjectAssemblyKompas>)kOSpecificationSection2;
                kOSpecificationSectionSorted2.Sort(delegate (ObjectAssemblyKompas x, ObjectAssemblyKompas y)
                { return x.Designation.CompareTo(y.Designation); });
                foreach (ObjectAssemblyKompas item in kOSpecificationSectionSorted2)
                {
                    sortedListObjects.Add(item);
                }
                var kOSpecificationSection3 = objectsAssemblyKompas.FindAll((ObjectAssemblyKompas) => ObjectAssemblyKompas.Parent == Parent &&
                                                                                                      ObjectAssemblyKompas.SpecificationSection == "Стандартные изделия");
                List<ObjectAssemblyKompas> kOSpecificationSectionSorted3 = (List<ObjectAssemblyKompas>)kOSpecificationSection3;
                kOSpecificationSectionSorted3.Sort(delegate (ObjectAssemblyKompas x, ObjectAssemblyKompas y)
                { return x.Name.CompareTo(y.Name); });
                foreach (ObjectAssemblyKompas item in kOSpecificationSectionSorted3)
                {
                    sortedListObjects.Add(item);
                }
                var kOSpecificationSection4 = objectsAssemblyKompas.FindAll((ObjectAssemblyKompas) => ObjectAssemblyKompas.Parent == Parent &&
                                                                                                      ObjectAssemblyKompas.SpecificationSection == "Прочие изделия");
                List<ObjectAssemblyKompas> kOSpecificationSectionSorted4 = (List<ObjectAssemblyKompas>)kOSpecificationSection4;
                kOSpecificationSectionSorted4.Sort(delegate (ObjectAssemblyKompas x, ObjectAssemblyKompas y)
                { return x.Name.CompareTo(y.Name); });
                foreach (ObjectAssemblyKompas item in kOSpecificationSectionSorted4)
                {
                    sortedListObjects.Add(item);
                }
            }
            RecursionMethod(kompasObject.Designation + " - " + kompasObject.Name);
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
                    objectAssemblyKompas.Mass = Math.Round(info, 2);
                }
                if (item.Name == "Покрытие")
                {
                    dynamic info;
                    bool source;
                    propertyKeeper.GetPropertyValue((_Property)item, out info, false, out source);
                    objectAssemblyKompas.Coating = info;
                }
            }
            //присваиваю полное имя
            objectAssemblyKompas.FullName = part7.FileName;

            #region Тут присваиваю свойство куда входит
            if (Name != "0")
            {
                objectAssemblyKompas.Parent = Name;
            }
            else { fileName = objectAssemblyKompas.Designation + " - " + objectAssemblyKompas.Name; }
            #endregion

            #region Тут расчет кол-ва и добавление в коллекцию
            ObjectAssemblyKompas objectK = objectsAssemblyKompas.SingleOrDefault((ObjectAssemblyKompas) => ObjectAssemblyKompas.Designation == objectAssemblyKompas.Designation &&
                                                                                                           ObjectAssemblyKompas.Name == objectAssemblyKompas.Name &&
                                                                                                           ObjectAssemblyKompas.Parent == objectAssemblyKompas.Parent);
            if (objectK != null)
            {
                objectK.Quantity++;
            }
            else if (objectK == null)
            {
                objectAssemblyKompas.Quantity++;
                objectsAssemblyKompas.Add(objectAssemblyKompas);
            }
            #endregion
        }

        private void FillTable()
        {
            GetSortedObjectsKompas();
            //dataGridView1.Rows.Clear();
            //dataGridView1.Columns.Clear();

            dataGridView1.DataSource = sortedListObjects; /*objectsAssemblyKompas;*/
            dataGridView1.Columns["Designation"].HeaderText = "Обозначение";
            dataGridView1.Columns["Name"].HeaderText = "Наименование";
            dataGridView1.Columns["Quantity"].HeaderText = "Кол-во";
            dataGridView1.Columns["Material"].HeaderText = "Материал";
            dataGridView1.Columns["SpecificationSection"].HeaderText = "Раздел спецификации";
            dataGridView1.Columns["Mass"].HeaderText = "Масса";
            dataGridView1.Columns["Coating"].HeaderText = "Покрытие";
            dataGridView1.Columns["Parent"].HeaderText = "Куда входит";
            dataGridView1.Columns["Bending"].HeaderText = "Гибка";
            dataGridView1.Columns["FullName"].HeaderText = "Путь до файла";

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
            dataGridView1.AllowUserToAddRows = false;
        }
        private void toolStripButton1_Click(object sender, EventArgs e)
        {
            if (objectsAssemblyKompas != null)
            {
                objectsAssemblyKompas.Clear();
            }
            else
            {
                objectsAssemblyKompas = new List<ObjectAssemblyKompas>();
            }
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
                        DisassembleObject(part7, "0");
                        FillTable();
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
                                Recursion(_part7, part7.Marking + " - " + part7.Name);
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
        private void toolStripButton2_Click(object sender, EventArgs e)
        {
            XLWorkbook excelWorkbook = new XLWorkbook();
            IXLWorksheet worksheet = excelWorkbook.Worksheets.Add("Тест");
            string path = System.Reflection.Assembly.GetExecutingAssembly().Location.Remove(System.Reflection.Assembly.GetExecutingAssembly().Location.Length - 16);

            #region создаю стиль заголовков
            var myCustomStyle1 = XLWorkbook.DefaultStyle;
            myCustomStyle1.Font.FontName = "Arial Cyr";
            myCustomStyle1.Font.Bold = false;
            myCustomStyle1.Font.Italic = false;
            myCustomStyle1.Font.FontSize = 10;
            myCustomStyle1.Border.LeftBorder = XLBorderStyleValues.Thin;
            myCustomStyle1.Border.RightBorder = XLBorderStyleValues.Thin;
            myCustomStyle1.Border.TopBorder = XLBorderStyleValues.Thin;
            myCustomStyle1.Border.BottomBorder = XLBorderStyleValues.Thin;
            myCustomStyle1.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
            myCustomStyle1.Alignment.SetVertical(XLAlignmentVerticalValues.Center);
            #endregion

            #region создаю стиль ячеек
            var myCustomStyle2 = XLWorkbook.DefaultStyle;
            myCustomStyle2.Font.FontName = "Arial Cyr";
            myCustomStyle2.Font.Bold = false;
            myCustomStyle2.Font.Italic = false;
            myCustomStyle2.Font.FontSize = 10;
            myCustomStyle2.Border.LeftBorder = XLBorderStyleValues.Thin;
            myCustomStyle2.Border.RightBorder = XLBorderStyleValues.Thin;
            myCustomStyle2.Border.TopBorder = XLBorderStyleValues.Thin;
            myCustomStyle2.Border.BottomBorder = XLBorderStyleValues.Thin;
            myCustomStyle2.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Left);
            myCustomStyle2.Alignment.SetVertical(XLAlignmentVerticalValues.Center);
            #endregion
            worksheet.Cell(1, 2).Value = "Обозначение";
            worksheet.Cell(1, 3).Value = "Наименование";
            worksheet.Cell(1, 4).Value = "Количество";
            worksheet.Cell(1, 5).Value = "Раздел спецификации";
            worksheet.Cell(1, 6).Value = "Материал";
            worksheet.Cell(1, 7).Value = "Масса";
            worksheet.Cell(1, 8).Value = "Покрытие";
            worksheet.Cell(1, 9).Value = "Куда входит";
            for (int i = 2; i < 10; i++)
            {
                worksheet.Cell(1, i).Style = myCustomStyle1;
            }


            for (int i = 0; i < dataGridView1.RowCount; i++)
            {
                for (int j = 0; j < dataGridView1.ColumnCount; j++)
                {
                    worksheet.Cell(i + 2, j + 2).Value = dataGridView1.Rows[i].Cells[j].Value;
                    worksheet.Cell(i + 2, j + 2).Style = myCustomStyle2;
                }
            }
            worksheet.Columns().AdjustToContents();
            excelWorkbook.SaveAs(path + fileName + ".xlsx");
        }

        private void contextMenuStrip1_Opening(object sender, CancelEventArgs e)
        {
            if (cancelContextMenu)
            {
                e.Cancel = true;
            }
        }

        private void MenuItemSelected_Click(object sender, EventArgs e)
        {
            DataGridViewSelectedRowCollection selectedRows = dataGridView1.SelectedRows;
            foreach (DataGridViewRow selectedRow in selectedRows)
            {
                int rowIndex = selectedRow.Index;

                if (rowIndex < 0)
                {
                    continue;
                }
                ObjectAssemblyKompas objectAssemblyKompas = objectsAssemblyKompas[rowIndex];
                IDocuments document = application.Documents;
                document.Open(objectAssemblyKompas.FullName, true, false);


            }
        }
        private void dataGridView1_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                var hitTestInfo = dataGridView1.HitTest(e.X, e.Y);
                if (hitTestInfo.RowIndex >= 0 && hitTestInfo.ColumnIndex >= 0)
                {
                    dataGridView1.ClearSelection();
                    dataGridView1.Rows[hitTestInfo.RowIndex].Selected = true;
                    cancelContextMenu = false;
                }
                else
                {
                    cancelContextMenu = true;
                }
            }
        }
    }
}
