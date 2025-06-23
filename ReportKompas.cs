using ClosedXML.Excel;
using Kompas6API5;
using KompasAPI7;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;

namespace ReportKompas
{
    /// <summary> Пример взял тут
    /// https://allineed.ru/development/dotnet-development/charp-development/80-csharp-working-with-datagridview
    /// </summary>
    public partial class ReportKompas : Form
    {
        public static IApplication application;
        IKompasDocument3D document3D;
        KompasObject kompas;
        ksDocument3D ksDocument3D;
        private static ReportKompas instance;

        public static List<ObjectAssemblyKompas> objectsAssemblyKompas;
        BindingList<ObjectAssemblyKompas> sortedListObjects;
        private bool cancelContextMenu = false;
        string fileName;
        string pathForExcel;
        string topParent;

        public static ReportKompas GetInstance()
        {
            if (instance == null || instance.IsDisposed)
            {
                instance = new ReportKompas();
            }
            return instance;
        }

        public ReportKompas()
        {
            InitializeComponent();
        }

        private void Recursion(IPart7 Part, string ParentName)
        {
            DisassembleObject(Part, ParentName);
            foreach (IPart7 item in Part.Parts)
            {
                ksPart ksPart2 = kompas.TransferInterface(item, 1, 0);
                if (ksPart2.excluded != true)
                {
                    if (item.Detail == true) DisassembleObject(item, Part.Marking + " - " + Part.Name);
                    if (item.Detail == false) Recursion(item, Part.Marking + " - " + Part.Name);
                }
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
            objectsAssemblyKompas.Remove(kompasObject);
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
                    objectsAssemblyKompas.Remove(item);
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
                    objectsAssemblyKompas.Remove(item);
                }
                var kOSpecificationSection3 = objectsAssemblyKompas.FindAll((ObjectAssemblyKompas) => ObjectAssemblyKompas.Parent == Parent &&
                                                                                                      ObjectAssemblyKompas.SpecificationSection == "Стандартные изделия");
                List<ObjectAssemblyKompas> kOSpecificationSectionSorted3 = (List<ObjectAssemblyKompas>)kOSpecificationSection3;
                kOSpecificationSectionSorted3.Sort(delegate (ObjectAssemblyKompas x, ObjectAssemblyKompas y)
                { return x.Name.CompareTo(y.Name); });
                foreach (ObjectAssemblyKompas item in kOSpecificationSectionSorted3)
                {
                    sortedListObjects.Add(item);
                    objectsAssemblyKompas.Remove(item);
                }
                var kOSpecificationSection4 = objectsAssemblyKompas.FindAll((ObjectAssemblyKompas) => ObjectAssemblyKompas.Parent == Parent &&
                                                                                                      ObjectAssemblyKompas.SpecificationSection == "Прочие изделия");
                List<ObjectAssemblyKompas> kOSpecificationSectionSorted4 = (List<ObjectAssemblyKompas>)kOSpecificationSection4;
                kOSpecificationSectionSorted4.Sort(delegate (ObjectAssemblyKompas x, ObjectAssemblyKompas y)
                { return x.Name.CompareTo(y.Name); });
                foreach (ObjectAssemblyKompas item in kOSpecificationSectionSorted4)
                {
                    sortedListObjects.Add(item);
                    objectsAssemblyKompas.Remove(item);
                }
            }
            RecursionMethod(kompasObject.Designation + " - " + kompasObject.Name);
        }

        /// <summary>
        /// Вытащит все свойства модели и передаст в коллекцию если в ней такого объекта нет
        /// </summary>
        /// <param name="part7"></param>
        /// <param name="Name"></param>
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
                if (item.Name == "Сварочные работы")
                {
                    dynamic info;
                    bool source;
                    propertyKeeper.GetPropertyValue((_Property)item, out info, false, out source);
                    objectAssemblyKompas.Welding = info;
                }
                if (item.Name == "Слесарные работы")
                {
                    dynamic info;
                    bool source;
                    propertyKeeper.GetPropertyValue((_Property)item, out info, false, out source);
                    objectAssemblyKompas.LocksmithWork = info;
                }
                if (item.Name == "Примечание")
                {
                    dynamic info;
                    bool source;
                    propertyKeeper.GetPropertyValue((_Property)item, out info, false, out source);
                    objectAssemblyKompas.Note = info;
                }
            }
            //присваиваю полное имя
            objectAssemblyKompas.FullName = part7.FileName;

            #region Тут присваиваю свойство куда входит
            if (Name != "0")
            {
                objectAssemblyKompas.Parent = Name;
                objectAssemblyKompas.TopParent = topParent;
            }
            else { fileName = objectAssemblyKompas.Designation + " - " + objectAssemblyKompas.Name; }
            #endregion

            #region Заполняю габаритные размеры и площадь поверхности
            //IFeature7 featureDim = (IFeature7)part7;
            //IBody7 bodyies = featureDim.ResultBodies;

            //IModelObject modelObject = (IModelObject)part7;
            //IFeature7 feature72 = modelObject.Owner;
            //IBody7 body7 = feature72.ResultBodies;
            //IBody7 body7 = (IBody7)part7.Owner.ResultBodies;
            ksPart ksPart = kompas.TransferInterface(part7, 1, 0);

            if (ksPart != null)
            {
                double x1, x2, y1, y2, z1, z2;
                ksPart.GetGabarit(true, true, out x1, out y1, out z1, out x2, out y2, out z2);
                string TemporaryVariable = String.Format("{0}x{1}x{2}", Math.Round(x2 - x1),
                                                                        Math.Round(y2 - y1),
                                                                        Math.Round(z2 - z1));
                if (TemporaryVariable.Contains("E") != true)
                {
                    objectAssemblyKompas.OverallDimensions = TemporaryVariable;
                }

                uint bitVector = 0x3;
                ksMassInertiaParam ksMassInertiaParam = ksPart.CalcMassInertiaProperties(bitVector);
                objectAssemblyKompas.Area = Math.Round(ksMassInertiaParam.F,2).ToString();

            }
            #endregion            

            #region Присваиваю путь до DXF и заполняю графу гибка
            try
            {
                ISheetMetalContainer sheetMetalContainer = part7 as ISheetMetalContainer;
                ISheetMetalBodies sheetMetalBodies = sheetMetalContainer.SheetMetalBodies;
                ISheetMetalBody sheetMetalBody = sheetMetalBodies.SheetMetalBody[0];

                if (sheetMetalBody != null) //если у детали нет свойства Толщина металла или не будет dxf в папке там же где модель то путь до DXF не будет указан
                {
                    string save_to_name = sheetMetalBody.Thickness.ToString(CultureInfo.CreateSpecificCulture("en-GB")) + "mm_" + part7.Marking.Remove(0, 3) + ".dxf";

                    string save_to_name2 = sheetMetalBody.Thickness.ToString(CultureInfo.CreateSpecificCulture("en-GB")).Replace('.', ',') + "mm_" + part7.Marking.Remove(0, 3) + ".dxf";

                    #region Заполняю свойство Гибка
                    //тут подсчет сколько гибов в теле "листовое тело"
                    IFeature7 pFeat = sheetMetalBody.Owner;
                    Object[] featCol = pFeat.SubFeatures[0, false, false];
                    int featColCount = 0;
                    if (featCol != null)
                    {
                        featColCount = featCol.Count();
                    }

                    double R = 0;
                    string V = "";
                    int Q;
                    if (sheetMetalBody.Thickness < 3)
                    { R = 1; }
                    else if (sheetMetalBody.Thickness > 2 && sheetMetalBody.Thickness < 6)
                    { R = 3; }
                    else if (sheetMetalBody.Thickness > 5 && sheetMetalBody.Thickness < 11)
                    { R = 6; }
                    switch (sheetMetalBody.Thickness)
                    {
                        case 0.7:
                            V = "8";
                            break;
                        case 0.8:
                            V = "8";
                            break;
                        case 1:
                            V = "8";
                            break;
                        case 1.2:
                            V = "8";
                            break;
                        case 1.5:
                            V = "12";
                            break;
                        case 2:
                            V = "12";
                            break;
                        case 3:
                            V = "16";
                            break;
                        case 4:
                            V = "22/30";
                            break;
                        case 5:
                            V = "35";
                            break;
                        case 6:
                            V = "35/50";
                            break;
                        case 8:
                            V = "50";
                            break;
                        case 10:
                            V = "80";
                            break;
                    }
                    int Q2 = 0; //считаю сгибы
                    for (int i = 0; i < sheetMetalContainer.SheetMetalBends.Count; i++)
                    {
                        IFeature7 feature7 = (IFeature7)sheetMetalContainer.SheetMetalBends[i];
                        Object[] featCol2 = feature7.SubFeatures[0, false, false];
                        if (feature7.Excluded != true)
                        {
                            Q2 = Q2 + featCol2.Count();
                        }
                    }
                    int Q3 = 0; //считаю сгибы нарисованные по линии
                    for (int i = 0; i < sheetMetalContainer.SheetMetalSketchBends.Count; i++)
                    {
                        IFeature7 feature7 = (IFeature7)sheetMetalContainer.SheetMetalSketchBends[i];
                        Object[] featCol3 = feature7.SubFeatures[0, false, false];
                        if (feature7.Excluded != true)
                        {
                            Q3 = Q3 + featCol3.Count();
                        }
                    }
                    int Q4 = 0; //считаю сгибы нарисованные по линии
                    for (int i = 0; i < sheetMetalContainer.SheetMetalLineBends.Count; i++)
                    {
                        IFeature7 feature7 = (IFeature7)sheetMetalContainer.SheetMetalLineBends[i];
                        Object[] featCol4 = feature7.SubFeatures[0, false, false];
                        if (feature7.Excluded != true)
                        {
                            Q4 = Q4 + featCol4.Count();
                        }
                    }
                    Q = featColCount + Q2 + Q3 + Q4;
                    if (Q != 0)
                    {
                        objectAssemblyKompas.R = R.ToString();
                        objectAssemblyKompas.V = V.ToString();
                        objectAssemblyKompas.Q = Q.ToString();
                        //objectAssemblyKompas.Bending = "R=" + R.ToString() + "  V=" + V + "  Q=" + Q;
                    }

                    #endregion

                    //обработка строки FileName
                    //string pattern = "\\\\";
                    //string replacement = "\\";
                    //Regex rgx = new Regex(pattern);
                    //string pathFull = rgx.Replace(part7.FileName, replacement);

                    FileInfo fi = new FileInfo(part7.FileName);
                    if (File.Exists(fi.DirectoryName + "\\" + save_to_name) || File.Exists(fi.DirectoryName + "\\" + save_to_name2))
                    {
                        objectAssemblyKompas.PathToDXF = fi.DirectoryName;
                    }
                }
            }
            catch (ArgumentException)
            {
                objectAssemblyKompas.PathToDXF = "";
            }
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
            dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            dataGridView1.Columns["Designation"].HeaderText = "Обозначение";
            dataGridView1.Columns["Designation"].DisplayIndex = 0;
            dataGridView1.Columns["Designation"].Width = 200;
            dataGridView1.Columns["Name"].HeaderText = "Наименование";
            dataGridView1.Columns["Name"].DisplayIndex = 1;
            dataGridView1.Columns["Name"].Width = 200;
            dataGridView1.Columns["Quantity"].HeaderText = "Кол-во";
            dataGridView1.Columns["Quantity"].DisplayIndex = 2;
            dataGridView1.Columns["Quantity"].Width = 50;
            dataGridView1.Columns["SpecificationSection"].HeaderText = "Раздел спецификации";
            dataGridView1.Columns["SpecificationSection"].DisplayIndex = 3;
            dataGridView1.Columns["SpecificationSection"].Width = 120;
            dataGridView1.Columns["Material"].HeaderText = "Материал";
            dataGridView1.Columns["Material"].DisplayIndex = 4;
            dataGridView1.Columns["Material"].Width = 150;
            dataGridView1.Columns["Mass"].HeaderText = "Масса";
            dataGridView1.Columns["Mass"].Width = 50;
            dataGridView1.Columns["Mass"].DisplayIndex = 5;
            dataGridView1.Columns["R"].HeaderText = "Пуансон";
            dataGridView1.Columns["R"].DisplayIndex = 6;
            dataGridView1.Columns["R"].Width = 50;
            dataGridView1.Columns["V"].HeaderText = "Матрица";
            dataGridView1.Columns["V"].DisplayIndex = 7;
            dataGridView1.Columns["V"].Width = 50;
            dataGridView1.Columns["Q"].HeaderText = "Кол-во гибов";
            dataGridView1.Columns["Q"].DisplayIndex = 8;
            dataGridView1.Columns["Q"].Width = 50;
            dataGridView1.Columns["Parent"].HeaderText = "Узел-1";
            dataGridView1.Columns["Parent"].DisplayIndex = 9;
            dataGridView1.Columns["TopParent"].HeaderText = "Узел верхний";
            dataGridView1.Columns["TopParent"].DisplayIndex = 10;
            dataGridView1.Columns["FullName"].HeaderText = "Путь до файла";
            dataGridView1.Columns["FullName"].DisplayIndex = 11;
            dataGridView1.Columns["PathToDXF"].HeaderText = "Путь до DXF";
            dataGridView1.Columns["PathToDXF"].Width = 200;
            dataGridView1.Columns["PathToDXF"].DisplayIndex = 12;
            dataGridView1.Columns["PathToDXF"].Visible = false;
            dataGridView1.Columns["OverallDimensions"].HeaderText = "Габаритные размеры";
            dataGridView1.Columns["OverallDimensions"].Width = 200;
            dataGridView1.Columns["OverallDimensions"].DisplayIndex = 13;
            dataGridView1.Columns["Coating"].HeaderText = "Покрытие";
            dataGridView1.Columns["Coating"].DisplayIndex = 14;
            dataGridView1.Columns["Coating"].Width = 120;
            dataGridView1.Columns["Welding"].HeaderText = "Сварочные работы";
            dataGridView1.Columns["Welding"].DisplayIndex = 15;
            dataGridView1.Columns["Welding"].Width = 120;
            dataGridView1.Columns["LocksmithWork"].HeaderText = "Слесарные работы";
            dataGridView1.Columns["LocksmithWork"].DisplayIndex = 15;
            dataGridView1.Columns["LocksmithWork"].Width = 120;
            dataGridView1.Columns["Note"].HeaderText = "Примечание";
            dataGridView1.Columns["Note"].DisplayIndex = 16;
            dataGridView1.Columns["Note"].Width = 75;
            dataGridView1.Columns["Area"].HeaderText = "Площадь поверхности";
            dataGridView1.Columns["Area"].DisplayIndex = 17;
            dataGridView1.Columns["Area"].Width = 75;
            dataGridView1.RowHeadersVisible = false;
            dataGridView1.AllowUserToAddRows = false;
        }

        private void toolStripButtonShowData_Click(object sender, EventArgs e)
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
            topParent = part7.Marking + " - " + part7.Name;
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

        private void contextMenuStrip1_Opening(object sender, CancelEventArgs e)
        {
            if (cancelContextMenu)
            {
                e.Cancel = true;
            }
        }

        private void MenuItemOpenInKompas_Click(object sender, EventArgs e)
        {
            DataGridViewSelectedRowCollection selectedRows = dataGridView1.SelectedRows;
            foreach (DataGridViewRow selectedRow in selectedRows)
            {
                int rowIndex = selectedRow.Index;

                if (rowIndex < 0)
                {
                    continue;
                }
                ObjectAssemblyKompas objectAssemblyKompas = sortedListObjects[rowIndex];
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

        private void OpenExplorer_MenuItem_Click(object sender, EventArgs e)
        {
            DataGridViewSelectedRowCollection selectedRows = dataGridView1.SelectedRows;

            string path = "";

            foreach (DataGridViewRow selectedRow in selectedRows)
            {
                int rowIndex = selectedRow.Index;

                if (rowIndex < 0)
                {
                    continue;
                }
                FileInfo fi = new FileInfo(sortedListObjects[rowIndex].FullName);
                path = fi.DirectoryName;
            }

            Process.Start("explorer.exe", path);
        }

        private void ShowLostParts_toolStripButton_Click(object sender, EventArgs e)
        {
            try
            {
                if (objectsAssemblyKompas.Count != 0 && objectsAssemblyKompas != null)
                {
                    LostParts lostParts = new LostParts();
                    lostParts.Show();
                    lostParts.dataGridView2.DataSource = objectsAssemblyKompas;
                    lostParts.dataGridView2.Columns["Designation"].HeaderText = "Обозначение";
                    lostParts.dataGridView2.Columns["Designation"].Width = 200;
                    lostParts.dataGridView2.Columns["Name"].HeaderText = "Наименование";
                    lostParts.dataGridView2.Columns["Name"].Width = 200;
                    lostParts.dataGridView2.RowHeadersVisible = false;
                    //lostParts.dataGridView2.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
                    lostParts.dataGridView2.Columns["Quantity"].Visible = false;
                    lostParts.dataGridView2.Columns["Material"].Visible = false;
                    lostParts.dataGridView2.Columns["SpecificationSection"].Visible = false;
                    lostParts.dataGridView2.Columns["Mass"].Visible = false;
                    lostParts.dataGridView2.Columns["Coating"].Visible = false;
                    lostParts.dataGridView2.Columns["TopParent"].Visible = false;
                    lostParts.dataGridView2.Columns["R"].Visible = false;
                    lostParts.dataGridView2.Columns["V"].Visible = false;
                    lostParts.dataGridView2.Columns["Q"].Visible = false;
                    lostParts.dataGridView2.Columns["FullName"].HeaderText = "Путь до файла";
                    lostParts.dataGridView2.Columns["Parent"].HeaderText = "Узел-1";
                    lostParts.dataGridView2.Columns["PathToDXF"].Visible = false;
                    lostParts.dataGridView2.Columns["OverallDimensions"].HeaderText = "Габаритные размеры";

                    lostParts.dataGridView2.AllowUserToAddRows = false;
                    lostParts.dataGridView2.Columns["FullName"].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
                }
                else
                {
                    MessageBox.Show("Пропущенных компонентов нет");
                }
            }
            catch (NullReferenceException)
            {
                MessageBox.Show("Пропущенных компонентов нет");
            }


        }

        private void SaveExcel_ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            XLWorkbook excelWorkbook = new XLWorkbook();
            pathForExcel = System.Reflection.Assembly.GetExecutingAssembly().Location.Remove(System.Reflection.Assembly.GetExecutingAssembly().Location.Length - 16);
            DataTable dt = new DataTable();
            foreach (DataGridViewColumn column in dataGridView1.Columns)
            {
                dt.Columns.Add(column.HeaderText, column.ValueType);
            }

            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                dt.Rows.Add();
                foreach (DataGridViewCell cell in row.Cells)
                {
                    if (cell.Value != null)
                    {
                        dt.Rows[dt.Rows.Count - 1][cell.ColumnIndex] = cell.Value.ToString();
                    }
                }
            }

            IXLWorksheet worksheet = excelWorkbook.Worksheets.Add(dt, "Отчет");
            worksheet.Outline.SummaryVLocation = XLOutlineSummaryVLocation.Top;

            int rowsCount = worksheet.LastRowUsed().RowNumber(); //кол-во заполненных строк
            worksheet.Rows(3, rowsCount).Group();
            //IXLCells cells = worksheet.CellsUsed(x => x.Value.ToString() == "Сборочные единицы");

            IXLCells cells2 = worksheet.CellsUsed(x => x.Value.ToString() == "Узел-1");
            int columnNumber = 0;
            foreach (IXLCell item in cells2)
            {
                columnNumber = item.WorksheetColumn().ColumnNumber();
            }

            IXLCells xLCells = worksheet.CellsUsed(c => c.WorksheetColumn().ColumnNumber() == columnNumber);

            List<int> rowNumbers = new List<int>();

            for (int i = 4; i < worksheet.LastRowUsed().RowNumber(); i++)
            {
                string temporaryName = worksheet.Row(i).Cell(columnNumber).Value.ToString();
                if (temporaryName.Contains(fileName))
                {
                    break;
                }
                rowNumbers.Add(worksheet.Row(i).RowNumber());
                for (int j = i; j < worksheet.LastRowUsed().RowNumber(); j++)
                {
                    if (worksheet.Row(j).Cell(columnNumber).Value.ToString() == temporaryName)
                    {
                        rowNumbers.Add(worksheet.Row(j).RowNumber());
                        if (j != worksheet.LastRowUsed().RowNumber())
                        {
                            i = j + 1;
                        }
                    }
                }
                if (rowNumbers.Count > 1)
                {
                    worksheet.Rows(rowNumbers[0], rowNumbers[rowNumbers.Count - 1]).Group();
                    rowNumbers.Clear();
                    temporaryName = "";
                }
            }

            //foreach (IXLCell item in xLCells)
            //{
            //    string temporaryName = "";
            //    int iterator = 0;
            //    if (item.Value.ToString() != "" && item.Value.ToString() != fileName && item.Value.ToString() != temporaryName && iterator == 0 && item.WorksheetRow().RowNumber() != 1)
            //    {
            //        temporaryName = item.Value.ToString();
            //        rowNumbers.Add(item.WorksheetRow().RowNumber());
            //    }
            //    if (rowNumbers.Count > 1 && item.Value.ToString() == fileName && item.Value.ToString() != "" && item.WorksheetRow().RowNumber() != 1)
            //    {
            //        worksheet.Rows(rowNumbers[0], rowNumbers[rowNumbers.Count-1]).Group();
            //        rowNumbers.Clear();
            //        temporaryName = "";
            //        iterator = 0;
            //    }
            //}

            IXLTable xLTable = worksheet.Table(0);

            xLTable.Theme = XLTableTheme.TableStyleLight8;
            worksheet.Columns().AdjustToContents();
            excelWorkbook.SaveAs(pathForExcel + fileName + ".xlsx");
            #region Сохраняю CSV файл
            //string csvFilePath = pathForExcel + fileName + ".csv";
            //using (var writer = new StreamWriter(csvFilePath))
            //{
            //    // Перебираем строки в листе
            //    foreach (var row in worksheet.RowsUsed())
            //    {
            //        // Создаем массив для хранения значений ячеек
            //        var values = new string[row.LastCellUsed().Address.ColumnNumber];

            //        // Перебираем ячейки в строке
            //        for (int i = 1; i <= row.LastCellUsed().Address.ColumnNumber; i++)
            //        {
            //            values[i - 1] = row.Cell(i).GetString();
            //        }
            //        // Записываем значения в CSV формате
            //        writer.WriteLine(string.Join(",", values));
            //    }
            //}
            #endregion
        }

        private void OpenExplorer_ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //Process.Start(System.Reflection.Assembly.GetExecutingAssembly().Location);
            string tempPath = System.Reflection.Assembly.GetExecutingAssembly().Location.ToString();
            Process.Start(tempPath.Remove(tempPath.LastIndexOf(@"\")));

            //Environment.CurrentDirectory
            //System.Reflection.Assembly.GetExecutingAssembly().Location


            //if (pathForExcel != "" & pathForExcel != null)
            //{
            //    Process.Start(pathForExcel);
            //}
        }

        private void SaveCSV_ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string pathForCSV = System.Reflection.Assembly.GetExecutingAssembly().Location.Remove(System.Reflection.Assembly.GetExecutingAssembly().Location.Length - 16);
            try
            {
                StringBuilder csvContent = new StringBuilder();

                // Записываем заголовки столбцов
                for (int i = 0; i < dataGridView1.Columns.Count; i++)
                {
                    csvContent.Append(dataGridView1.Columns[i].HeaderText);
                    if (i < dataGridView1.Columns.Count - 1)
                        csvContent.Append(";");
                }
                csvContent.AppendLine();

                // Записываем строки данных
                foreach (DataGridViewRow row in dataGridView1.Rows)
                {
                    if (!row.IsNewRow) // пропускаем новую пустую строку
                    {
                        for (int i = 0; i < dataGridView1.Columns.Count; i++)
                        {
                            var cellValue = row.Cells[i].Value?.ToString() ?? "";
                            // Оборачиваем значения, содержащие запятые, кавычками
                            if (cellValue.Contains(",") || cellValue.Contains("\"") || cellValue.Contains("\n"))
                            {
                                cellValue = $"\"{cellValue.Replace("\"", "\"\"")}\"";
                            }
                            csvContent.Append(cellValue);
                            if (i < dataGridView1.Columns.Count - 1)
                                csvContent.Append(";");
                        }
                        csvContent.AppendLine();
                    }
                }

                // Записываем в файл
                File.WriteAllText(pathForCSV + fileName + ".csv", csvContent.ToString(), Encoding.UTF8);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка при сохранении файла: " + ex.Message);
            }
        }

        private void SaveXML_ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string pathForXML = System.Reflection.Assembly.GetExecutingAssembly().Location.Remove(System.Reflection.Assembly.GetExecutingAssembly().Location.Length - 16) + fileName + ".xml";
            XmlWriterSettings settings = new XmlWriterSettings();
            settings.Indent = true;
            using (XmlWriter writer = XmlWriter.Create(pathForXML, settings))
            {
                writer.WriteStartDocument();
                writer.WriteStartElement("Rows"); // корневой элемент

                // Перебираем все строки DataGridView
                foreach (DataGridViewRow row in dataGridView1.Rows)
                {
                    if (row.IsNewRow) continue; // пропускаем новую пустую строку

                    writer.WriteStartElement("Row");

                    // Перебираем все ячейки в строке
                    foreach (DataGridViewCell cell in row.Cells)
                    {
                        string columnName = dataGridView1.Columns[cell.ColumnIndex].Name;
                        string cellValue = cell.Value?.ToString() ?? "";

                        writer.WriteElementString(columnName, cellValue);
                    }

                    writer.WriteEndElement(); // </Row>
                }

                writer.WriteEndElement(); // </Rows>
                writer.WriteEndDocument();
            }
        }
    }
}
