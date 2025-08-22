using ClosedXML.Excel;
using Kompas6API5;
using Kompas6Constants;
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
using netDxf;
using netDxf.Entities;

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
        public Dictionary<string, List<string>> DictionaryCodeEquip;
        public Dictionary<string, List<string>> DictionaryCodeMaterial;
        public Dictionary<string, double> DictionarySpeedCut;
        public Columns columns;
        public Settings _settings;

        [DllImport("user32.dll", SetLastError = true)]
        static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

        [DllImport("user32.dll")]
        static extern bool SetForegroundWindow(IntPtr hWnd);

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
            LoadSettings();
        }

        public void LoadSettings()
        {
            if (_settings == null || _settings.IsDisposed)
            {
                _settings = new Settings();
                _settings.StartPosition = FormStartPosition.CenterParent;
                _settings.TopMost = true;
            }

            DictionaryCodeEquip = new Dictionary<string, List<string>>();
            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.Load(_settings?.Path_Dictionary_Equipment_textBox.Text + @"\" + "CodeEquip.xml");
            XmlNodeList keyNodes = xmlDoc.SelectNodes("/Dictionary/Key");
            foreach (XmlNode keyNode in keyNodes)
            {
                // Получаем имя ключа из атрибута 'name'
                string keyName = keyNode.Attributes["name"]?.InnerText;

                if (keyName != null)
                {
                    var values = new List<string>();

                    // Получаем все <Value> внутри текущего <Key>
                    foreach (XmlNode valueNode in keyNode.SelectNodes("Value"))
                    {
                        string valueText = valueNode.InnerText;
                        values.Add(valueText);
                    }

                    // Добавляем в словарь
                    DictionaryCodeEquip[keyName] = values;
                }
            }
            DictionaryCodeMaterial = new Dictionary<string, List<string>>();
            XmlDocument xmlDoc2 = new XmlDocument();
            xmlDoc2.Load(_settings?.Path_Dictionary_Materials_textBox.Text + @"\" + "CodeMaterial.xml");
            XmlNodeList keyNodes2 = xmlDoc2.SelectNodes("/Dictionary/Key");
            foreach (XmlNode keyNode in keyNodes2)
            {
                // Получаем имя ключа из атрибута 'name'
                string keyName = keyNode.Attributes["name"]?.InnerText;

                if (keyName != null)
                {
                    var values = new List<string>();

                    // Получаем все <Value> внутри текущего <Key>
                    foreach (XmlNode valueNode in keyNode.SelectNodes("Value"))
                    {
                        string valueText = valueNode.InnerText;
                        values.Add(valueText);
                    }

                    // Добавляем в словарь
                    DictionaryCodeMaterial[keyName] = values;
                }
            }
            DictionarySpeedCut = new Dictionary<string, double>();
            XmlDocument xmlDoc3 = new XmlDocument();
            xmlDoc3.Load(_settings?.Speed_Cut_textBox.Text + @"\" + "SpeedCut.xml");
            XmlNodeList keyNodes3 = xmlDoc3.SelectNodes("/Dictionary/Key");
            foreach (XmlNode keyNode in keyNodes3)
            {
                // Получаем атрибут 'name'
                string keyName = keyNode.Attributes["name"].InnerText;

                // Получаем значение внутри <Value>
                XmlNode valueNode = keyNode.SelectSingleNode("Value");
                if (valueNode != null)
                {
                    // Парсим значение в double
                    if (double.TryParse(valueNode.InnerText, out double value))
                    {
                        DictionarySpeedCut[keyName] = value;
                    }
                }
            }
        }

        private void Recursion(IPart7 Part, string ParentName)
        {
            #region Провожу проверку есть ли такой объект в коллекции
            ObjectAssemblyKompas objectF = objectsAssemblyKompas.SingleOrDefault((ObjectAssemblyKompas) => ObjectAssemblyKompas.Designation == Part.Marking &&
                                                                                                           ObjectAssemblyKompas.Name == Part.Name &&
                                                                                                           ObjectAssemblyKompas.Parent == ParentName);
            if (objectF != null)
            {
                objectF.Quantity++;
            }
            else if (objectF == null)
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
            #endregion            
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
                    info = info.Replace("$", "");
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
                    objectAssemblyKompas.Mass = Math.Round(info * 1.2, 2);
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
                if (objectAssemblyKompas.Coating != null && objectAssemblyKompas.Coating.Contains("Рекуперат"))
                {
                    objectAssemblyKompas.Area = Math.Round(ksMassInertiaParam.F, 2).ToString();
                }
                else { objectAssemblyKompas.Area = Math.Round(ksMassInertiaParam.F / 2, 2).ToString(); }


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
                            V = "30";
                            break;
                        case 5:
                            V = "35";
                            break;
                        case 6:
                            V = "50";
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

            #region Расчет времени резки
            bool checkBox = _settings.Other_Param_Laser_Cut_checkBox.Checked;
            if (objectAssemblyKompas.SpecificationSection == "Детали" && checkBox == true)
            {
                CuttingTimeCalculation(objectAssemblyKompas);
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
            ReplaceMaterial();
            AddCodeMaterial();
            AddCodeEquip();

            GetSortedObjectsKompas();
            //dataGridView1.Rows.Clear();
            //dataGridView1.Columns.Clear();
            dataGridView1.DataSource = sortedListObjects; /*objectsAssemblyKompas;*/
            dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None;
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
            dataGridView1.Columns["R"].Width = 60;
            dataGridView1.Columns["V"].HeaderText = "Матрица";
            dataGridView1.Columns["V"].DisplayIndex = 7;
            dataGridView1.Columns["V"].Width = 60;
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
            dataGridView1.Columns["OverallDimensions"].Width = 100;
            dataGridView1.Columns["OverallDimensions"].DisplayIndex = 13;
            dataGridView1.Columns["Coating"].HeaderText = "Покрытие";
            dataGridView1.Columns["Coating"].DisplayIndex = 14;
            dataGridView1.Columns["Coating"].Width = 80;
            dataGridView1.Columns["Welding"].HeaderText = "Сварочные работы";
            dataGridView1.Columns["Welding"].DisplayIndex = 15;
            dataGridView1.Columns["Welding"].Width = 100;
            dataGridView1.Columns["LocksmithWork"].HeaderText = "Слесарные работы";
            dataGridView1.Columns["LocksmithWork"].DisplayIndex = 16;
            dataGridView1.Columns["LocksmithWork"].Width = 80;
            dataGridView1.Columns["Note"].HeaderText = "Примечание";
            dataGridView1.Columns["Note"].DisplayIndex = 17;
            dataGridView1.Columns["Note"].Width = 75;
            dataGridView1.Columns["Area"].HeaderText = "Площадь поверхности";
            dataGridView1.Columns["Area"].DisplayIndex = 18;
            dataGridView1.Columns["Area"].Width = 75;
            dataGridView1.Columns["CodeEquipment"].HeaderText = "Код СИ";
            dataGridView1.Columns["CodeEquipment"].DisplayIndex = 19;
            dataGridView1.Columns["CodeEquipment"].Width = 50;
            dataGridView1.Columns["CodeMaterial"].HeaderText = "Код Мат";
            dataGridView1.Columns["CodeMaterial"].DisplayIndex = 20;
            dataGridView1.Columns["CodeMaterial"].Width = 50;
            dataGridView1.Columns["TimeCut"].HeaderText = "Время резки";
            dataGridView1.Columns["TimeCut"].DisplayIndex = 21;
            dataGridView1.Columns["TimeCut"].Width = 50;
            dataGridView1.Columns["DxfDimensions"].HeaderText = "Габариты DXF";
            dataGridView1.Columns["DxfDimensions"].DisplayIndex = 22;
            dataGridView1.Columns["DxfDimensions"].Width = 80;
            dataGridView1.RowHeadersVisible = false;
            dataGridView1.AllowUserToAddRows = false;
            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                var cell = row.Cells["SpecificationSection"];
                if (cell.Value != null && cell.Value.ToString().Contains("Сборочные единицы"))
                {
                    row.DefaultCellStyle.BackColor = Color.LightGreen;
                }
            }
            string tempPath = System.Reflection.Assembly.GetExecutingAssembly().Location.ToString();
            UpdateExistingColumnsFromXml(tempPath.Remove(tempPath.LastIndexOf(@"\")) + @"\" + @"Settings\Сolumns.xml", dataGridView1);
        }

        public void UpdateExistingColumnsFromXml(string xmlFilePath, DataGridView dgv)
        {
            // Очищаем текущие колонки
            //dgv.Columns.Clear();

            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.Load(xmlFilePath);

            // Находим раздел Columns
            XmlNode columnsNode = xmlDoc.SelectSingleNode("/Columns");
            if (columnsNode == null)
            {
                MessageBox.Show("Раздел <Columns> не найден в XML.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            foreach (XmlNode columnNode in columnsNode.ChildNodes)
            {
                if (columnNode.Name != "Column")
                    continue;

                string headerText = "";
                bool isVisible = true; // по умолчанию

                if (columnNode.Attributes["HeaderText"] != null)
                    headerText = columnNode.Attributes["HeaderText"].Value;

                if (columnNode.Attributes["Visible"] != null)
                    bool.TryParse(columnNode.Attributes["Visible"].Value, out isVisible);

                // Ищем колонку с таким HeaderText
                var existingCol = dgv.Columns.Cast<DataGridViewColumn>()
                                    .FirstOrDefault(c => c.HeaderText == headerText);

                if (existingCol != null)
                {
                    // Обновляем свойство Visible
                    existingCol.Visible = isVisible;
                }
                // Можно добавить обработку для случаев отсутствия колонки, если нужно
            }
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
                        this.Activate();
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
                        this.Activate();
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
            string windowTitle = "КОМПАС-3D v18.1";
            IntPtr hWnd = FindWindow(null, windowTitle);
            if (hWnd != IntPtr.Zero)
            {
                SetForegroundWindow(hWnd);
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
                    lostParts.dataGridView2.Columns["FullName"].Visible = false;
                    lostParts.dataGridView2.Columns["Parent"].HeaderText = "Узел-1";
                    lostParts.dataGridView2.Columns["PathToDXF"].Visible = false;
                    lostParts.dataGridView2.Columns["OverallDimensions"].HeaderText = "Габаритные размеры";
                    lostParts.dataGridView2.Columns["OverallDimensions"].Visible = false;
                    lostParts.dataGridView2.Columns["Welding"].Visible = false;
                    lostParts.dataGridView2.Columns["LocksmithWork"].Visible = false;
                    lostParts.dataGridView2.Columns["Note"].Visible = false;
                    lostParts.dataGridView2.Columns["Area"].Visible = false;
                    lostParts.dataGridView2.Columns["CodeEquipment"].Visible = false;
                    lostParts.dataGridView2.Columns["CodeMaterial"].Visible = false;
                    lostParts.dataGridView2.Columns["TimeCut"].Visible = false;
                    lostParts.dataGridView2.Columns["DxfDimentions"].Visible = false;
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
                MessageBox.Show("NullReferenceException");
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
            worksheet.RangeUsed().Style.NumberFormat.Format = "@";
            worksheet.Outline.SummaryVLocation = XLOutlineSummaryVLocation.Top;

            Dictionary<string, List<int>> groups = new Dictionary<string, List<int>>();
            // Заполняем данные и собираем группы по колонке "Parent"
            for (int rowIndex = 0; rowIndex < dataGridView1.Rows.Count; rowIndex++)
            {
                var row = dataGridView1.Rows[rowIndex];
                // Предположим, что колонка "Parent" есть и её индекс известен или по имени
                string parentValue = "";
                if (dataGridView1.Columns.Contains("Parent"))
                {
                    object val = row.Cells["Parent"].Value;
                    parentValue = val != null ? val.ToString() : "";
                }

                // Заполняем данные в Excel
                for (int col = 0; col < dataGridView1.Columns.Count; col++)
                {
                    worksheet.Cell(rowIndex + 2, col + 1).Value = row.Cells[col].Value;
                }

                // Добавляем индекс строки к группе
                if (!groups.ContainsKey(parentValue))
                    groups[parentValue] = new List<int>();
                // В Excel строки начинаются с 2 (заголовки на 1-й строке)
                groups[parentValue].Add(rowIndex + 2);
            }

            // Для каждой группы применяем группировку строк
            foreach (var group in groups.Values)
            {
                if (group.Count > 1)
                {
                    int startRow = group[0];
                    int endRow = group[group.Count - 1];
                    worksheet.Rows(startRow, endRow).Group();
                }
            }


            IXLTable xLTable = worksheet.Table(0);

            xLTable.Theme = XLTableTheme.TableStyleLight8;
            worksheet.Columns().AdjustToContents();
            excelWorkbook.SaveAs(pathForExcel + fileName + ".xlsx");
        }

        private void OpenExplorer_ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string tempPath = System.Reflection.Assembly.GetExecutingAssembly().Location.ToString();
            Process.Start(tempPath.Remove(tempPath.LastIndexOf(@"\")));
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

        private void AddCodeEquip()
        {
            if (objectsAssemblyKompas != null && DictionaryCodeEquip != null)
            {
                foreach (ObjectAssemblyKompas item in objectsAssemblyKompas)
                {
                    if (item.SpecificationSection == "Стандартные изделия" || item.SpecificationSection == "Прочие изделия")
                    {
                        if (DictionaryCodeEquip.Any(kvp => kvp.Value.Contains(item.Name)))
                        {
                            item.CodeEquipment = DictionaryCodeEquip.First(kvp => kvp.Value.Contains(item.Name)).Key;
                        }
                    }
                }
            }
        }

        private void AddCodeMaterial()
        {
            if (objectsAssemblyKompas != null && DictionaryCodeMaterial != null)
            {
                foreach (ObjectAssemblyKompas item in objectsAssemblyKompas)
                {
                    if (item.SpecificationSection == "Детали" || item.SpecificationSection == "Прочие изделия" && item.Material != null)
                    {
                        if (DictionaryCodeMaterial.Any(kvp => kvp.Value.Contains(item.Material)))
                        {
                            item.CodeMaterial = DictionaryCodeMaterial.First(kvp => kvp.Value.Contains(item.Material)).Key;
                        }
                    }
                }
            }
        }

        private void ReplaceMaterial()
        {
            if (objectsAssemblyKompas != null)
            {
                foreach (ObjectAssemblyKompas item in objectsAssemblyKompas)
                {
                    if (item.SpecificationSection == "Детали" && item.Material == "")
                    {
                        if (item.Designation.Contains("1.5mm_") && item.Designation.Contains("_Zn"))
                        {
                            item.Material = @"Лист ОЦd1,5 ГОСТ 19904-90;08пс ГОСТ 14918-80";
                        }
                        if (item.Designation.Contains("1.5mm_") && item.Designation.Contains("_Aisi"))
                        {
                            item.Material = @"Лист нерж. 1,5ммх1250х2500 AISI430 4N+PE";
                        }
                        if (item.Designation.Contains("1.5mm_") && item.Designation.Contains("_Aisi Bronze"))
                        {
                            item.Material = @"Лист нерж. 1,5ммх1250х2500 AISI430 4N+PE (Bronze)";
                        }
                        if (item.Designation.Contains("1mm_") && item.Designation.Contains("_Zn"))
                        {
                            item.Material = @"Лист ОЦd1,0 ГОСТ 19904-90;08пс ГОСТ 14918-80";
                        }
                        if (item.Designation.Contains("1.5mm_") && item.Designation.Contains("_AL"))
                        {
                            item.Material = @"Лист квинтет 1,5mm AL";
                        }
                        if (item.Designation.Contains("1.5mm_") && item.Designation.Contains("_Forbo"))
                        {
                            item.Material = @"Forbo flooring";
                        }
                    }
                }
            }
        }

        private void SettingsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            LoadSettings();
            _settings.ShowDialog();
        }

        private void toolStripButton2_Click(object sender, EventArgs e)
        {
            // Путь к файлу настроек
            string tempPath = System.Reflection.Assembly.GetExecutingAssembly().Location.ToString();
            string settingsFilePath = tempPath.Remove(tempPath.LastIndexOf(@"\")) + @"\" + @"Settings\Сolumns.xml";

            // Создаем форму с CheckedListBox

            if (columns == null || columns.IsDisposed)
            {
                columns = new Columns();
            }
            else
            {
                columns.checkedListBox1.Items.Clear();
            }

            columns.checkedListBox1.CheckOnClick = true;

            // Загружаем состояние из XML перед отображением
            LoadColumnsStateFromXml(settingsFilePath);

            // Заполняем CheckedListBox названиями колонок и их состоянием
            foreach (DataGridViewColumn column in dataGridView1.Columns)
            {
                bool isChecked = column.Visible; // или можно использовать сохраненное состояние
                columns.checkedListBox1.Items.Add(column.HeaderText, isChecked);
            }

            Button btnOk = new Button();
            btnOk.Dock = DockStyle.Fill;
            btnOk.Text = "Применить";

            btnOk.Click += (s, args) =>
            {
                // Обработка выбранных элементов: скрытие/показ колонок
                for (int i = 0; i < dataGridView1.Columns.Count; i++)
                {
                    var colHeader = dataGridView1.Columns[i].HeaderText;
                    bool isChecked = false;

                    if (columns.checkedListBox1.Items.Contains(colHeader))
                    {
                        int index = columns.checkedListBox1.Items.IndexOf(colHeader);
                        isChecked = columns.checkedListBox1.GetItemChecked(index);
                    }

                    dataGridView1.Columns[i].Visible = isChecked;
                }

                // Сохраняем состояние в XML после применения
                SaveColumnsStateToXml(settingsFilePath);

                columns.Close();
            };

            columns.tableLayoutPanel1.Controls.Add(btnOk, 0, 1);

            // Показываем форму
            columns.StartPosition = FormStartPosition.CenterParent;
            columns.ShowDialog();
        }
        private void SaveColumnsStateToXml(string filePath)
        {
            XmlDocument doc = new XmlDocument();
            XmlElement root = doc.CreateElement("Columns");
            doc.AppendChild(root);

            foreach (DataGridViewColumn column in dataGridView1.Columns)
            {
                XmlElement colElem = doc.CreateElement("Column");
                colElem.SetAttribute("HeaderText", column.HeaderText);
                colElem.SetAttribute("Visible", column.Visible.ToString());
                root.AppendChild(colElem);
            }

            doc.Save(filePath);
        }
        private void LoadColumnsStateFromXml(string filePath)
        {
            if (!File.Exists(filePath))
                return;

            XmlDocument doc = new XmlDocument();
            doc.Load(filePath);

            var columnsNodes = doc.SelectNodes("/Columns/Column");
            foreach (XmlNode node in columnsNodes)
            {
                string headerText = node.Attributes["HeaderText"].Value;
                bool isVisible = bool.Parse(node.Attributes["Visible"].Value);

                var column = dataGridView1.Columns
                    .Cast<DataGridViewColumn>()
                    .FirstOrDefault(c => c.HeaderText == headerText);
                if (column != null)
                {
                    column.Visible = isVisible;
                }
            }
        }

        private void CuttingTimeCalculation(ObjectAssemblyKompas objectAssemblyKompas)
        {
            string currentDirectory = System.Reflection.Assembly.GetExecutingAssembly().Location.ToString();
            string workspacePath = Path.Combine(currentDirectory.Remove(currentDirectory.LastIndexOf(@"\")), "Workspace");
            if (!Directory.Exists(workspacePath))
            {
                Directory.CreateDirectory(workspacePath);
            }

            #region Создаю DXF
            List<string> pathnameOpenparts = new List<string>();
            IDocuments document = application.Documents;
            for (int i = 0; i < document.Count; i++)
            {
                pathnameOpenparts.Add(document[i].PathName);
            }
            OpenDocumentParam openDocumentParam = document.GetOpenDocumentParam();
            openDocumentParam.ReadOnly = false;
            IKompasDocument kompasDocument = document.OpenDocument(objectAssemblyKompas.FullName, openDocumentParam);
            IKompasDocument3D kompasDocument3D = (IKompasDocument3D)kompasDocument;
            if (kompasDocument3D == null)
            {
                return;
            }
            IPart7 topPart = kompasDocument3D.TopPart;
            ISheetMetalContainer sheetMetalContainer = topPart as ISheetMetalContainer;
            ISheetMetalBodies sheetMetalBodies = sheetMetalContainer.SheetMetalBodies;
            ISheetMetalBody sheetMetalBody = sheetMetalBodies.SheetMetalBody[0];
            if (sheetMetalBody == null)
            {
                return;
            }

            //MessageBox.Show(workspacePath.ToString());

            string save_to_name = workspacePath + @"\" +
                sheetMetalBody.Thickness.ToString(CultureInfo.CreateSpecificCulture("en-GB")) + "mm_" + topPart.Marking/*.Remove(0, 3)*/ + ".dxf";

            KompasObject kompas2 = (KompasObject)Marshal.GetActiveObject("KOMPAS.Application.5");

            ksDocumentParam documentParam = (ksDocumentParam)kompas2.GetParamStruct(35);
            documentParam.type = 1;
            documentParam.Init();
            ksDocument2D document2D = (ksDocument2D)kompas2.Document2D();
            document2D.ksCreateDocument(documentParam);

            IKompasDocument2D kompasDocument2D = (IKompasDocument2D)application.ActiveDocument;

            //Скрываем все сообщения системы - Да
            application.HideMessage = ksHideMessageEnum.ksHideMessageYes;

            IViewsAndLayersManager viewsAndLayersManager = kompasDocument2D.ViewsAndLayersManager;
            IViews views = viewsAndLayersManager.Views;
            IView pView = views.Add(Kompas6Constants.LtViewType.vt_Arbitrary);

            IAssociationView pAssociationView = pView as IAssociationView;
            pAssociationView.SourceFileName = topPart.FileName;

            //IEmbodimentsManager embodimentsManager = (IEmbodimentsManager)document3D;
            //int indexPart = embodimentsManager.CurrentEmbodimentIndex;

            IEmbodimentsManager emb = (IEmbodimentsManager)pAssociationView;
            emb.SetCurrentEmbodiment(topPart.Marking);
            pAssociationView.Angle = 0;
            pAssociationView.X = 0;
            pAssociationView.Y = 0;
            pAssociationView.BendLinesVisible = false;
            pAssociationView.BreakLinesVisible = false;
            pAssociationView.HiddenLinesVisible = false;
            pAssociationView.VisibleLinesStyle = (int)ksCurveStyleEnum.ksCSNormal;
            pAssociationView.Scale = 1;
            pAssociationView.Name = "User view";
            pAssociationView.ProjectionName = "#Развертка";
            pAssociationView.Unfold = true; //развернутый вид
            pAssociationView.BendLinesVisible = false;
            pAssociationView.CenterLinesVisible = false;
            pAssociationView.SourceFileName = topPart.FileName;
            pAssociationView.Update();
            pView.Update();
            IViewDesignation pViewDesignation = pView as IViewDesignation;
            pViewDesignation.ShowUnfold = false;
            pViewDesignation.ShowScale = false;

            pView.Update();
            document2D.ksRebuildDocument();
            document2D.ksSaveDocument(save_to_name);
            IKompasDocument kompasDocument2 = (IKompasDocument)application.ActiveDocument;
            kompasDocument2.Close(DocumentCloseOptions.kdDoNotSaveChanges);
            kompasDocument.Close(DocumentCloseOptions.kdDoNotSaveChanges);
            IDocuments document2 = application.Documents;
            for (int j = 0; j < document2.Count; j++)
            {
                if (pathnameOpenparts.Contains(document2[j].PathName) != true)
                {
                    document2[j].Close(DocumentCloseOptions.kdDoNotSaveChanges);
                }
            }

            //Скрываем все сообщения системы - Нет
            application.HideMessage = ksHideMessageEnum.ksShowMessage;

            #endregion

            #region Считаю время резки
            DxfDocument dxf = DxfDocument.Load(save_to_name);
            double totalLengthMm = 0;
            foreach (var block in dxf.Blocks.Items)
            {
                foreach (var entity in block.Entities)
                {
                    if (entity.Type == EntityType.LwPolyline)
                    {
                        var polyline = entity as LwPolyline;
                        totalLengthMm += CalculatePolylineLength(polyline);
                    }
                    if (entity.Type == EntityType.Line)
                    {
                        var line = entity as Line;
                        totalLengthMm += Distance(line.StartPoint, line.EndPoint);
                    }
                    if (entity.Type == EntityType.Circle)
                    {
                        var circle = entity as Circle;
                        totalLengthMm += 2 * Math.PI * circle.Radius;
                    }
                    if (entity.Type == EntityType.Arc)
                    {
                        var arc = entity as Arc;
                        double angle = Math.Abs(arc.EndAngle - arc.StartAngle);
                        double arcLength = arc.Radius * (((angle > 180 ? 360 - angle : angle) * Math.PI) / 180);
                        totalLengthMm += arcLength;
                    }
                }
            }

            // Расчет времени в секундах
            double divisor = 10000; // значение по умолчанию

            var sortedKeys = DictionarySpeedCut.Keys.OrderByDescending(k => k.Length);

            //List<string> keysList = new List<string>(DictionarySpeedCut.Keys);
            //keysList.Sort((a, b) => b.Length.CompareTo(a.Length));

            string matchedKey = null;
            foreach (var key in sortedKeys)
            {
                if (save_to_name.Contains(key))
                {
                    matchedKey = key;
                    break;
                }
            }
            if (matchedKey != null)
            {
                divisor = DictionarySpeedCut[matchedKey];
            }

            DXFCountBurnPoint instance = new DXFCountBurnPoint(save_to_name);
            
            // далее расчет
            double timeWorkLaser = (totalLengthMm / divisor) * 60; //здесь 100мм/мин *60 - перевод в сек
            string[] stringsToCheck =
            {
                objectAssemblyKompas.Designation,
                objectAssemblyKompas.Material
            };

            foreach (var str in stringsToCheck)
            {
                bool containsIgnoreCase = str.IndexOf("Aisi", StringComparison.OrdinalIgnoreCase) >= 0;
                if (containsIgnoreCase)
                {
                    objectAssemblyKompas.TimeCut = ((Math.Round(timeWorkLaser, 1) * 2) + instance.burnPoint).ToString();
                }
                else
                {
                    objectAssemblyKompas.TimeCut = (Math.Round(timeWorkLaser, 1) + instance.burnPoint).ToString();
                }
            }
            
            #endregion

            #region Считаю габаритные размеры DXF
            DxfDimensions dims = DxfHelper.GetDxfDimensions(save_to_name);
            double X = dims.MaxX - dims.MinX;
            double Y = dims.MaxY - dims.MinY;
            objectAssemblyKompas.DxfDimensions = Math.Round(X, 0).ToString() + "x" + Math.Round(Y, 0).ToString();
            #endregion

            //System.Threading.Thread.Sleep(5000);
            //Directory.Delete(workspacePath, true); // true — удаляет рекурсивно, если есть содержимое
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
    }
}
