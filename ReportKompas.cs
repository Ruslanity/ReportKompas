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
using BrightIdeasSoftware;

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
        private ContextMenuStrip contextMenu;
        public TreeListView treeListView;

        public ObjectAssemblyKompas root;

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
        }

        private void SettingsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            // Загружаем настройки
            using (Settings settings = Settings.Load(Settings.DefaultPathSettings))
            {
                // Создаем и показываем форму для редактирования
                SettingsForm form = new SettingsForm(settings);
                form.ShowDialog();
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
                try
                {
                    pathnameOpenparts.Add(document[i].PathName);
                }
                catch (Exception)
                {

                }
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
                try
                {
                    if (pathnameOpenparts.Contains(document2[j].PathName) != true)
                    {
                        document2[j].Close(DocumentCloseOptions.kdDoNotSaveChanges);
                    }
                }
                catch (Exception)
                {

                }

            }

            //Скрываем все сообщения системы - Нет
            application.HideMessage = ksHideMessageEnum.ksShowMessage;

            #endregion

            #region Считаю время резки
            DxfProcessor dxfProcessor = new DxfProcessor(save_to_name);

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
                    objectAssemblyKompas.TimeCut = (dxfProcessor.TotalCuttingTime * 2).ToString();
                    break;
                }
                else
                {
                    objectAssemblyKompas.TimeCut = dxfProcessor.TotalCuttingTime.ToString();
                    break;
                }
            }
            #endregion

            #region Считаю габаритные размеры DXF            
            objectAssemblyKompas.DxfDimensions = Math.Round(dxfProcessor.Size.Width, 1).ToString() + "x" + Math.Round(dxfProcessor.Size.Height, 1).ToString();
            #endregion

            //System.Threading.Thread.Sleep(5000);
            //Directory.Delete(workspacePath, true); // true — удаляет рекурсивно, если есть содержимое
        }

        private void toolStripButton4_Click(object sender, EventArgs e)
        {
            this.TopMost = true;
            kompas = (KompasObject)Marshal.GetActiveObject("KOMPAS.Application.5");
            kompas.Visible = true;
            kompas.ActivateControllerAPI();
            ksDocument3D = (ksDocument3D)kompas.ActiveDocument3D();

            document3D = kompas.TransferInterface(ksDocument3D, 2, 0);
            application = kompas.ksGetApplication7();

            ksPartCollection _ksPartCollection = ksDocument3D.PartCollection(true);

            IPart7 part7 = document3D.TopPart;
            //topParent = part7.Marking + " - " + part7.Name;

            if (root != null)
            {
                root = null;
            }
            root = PrimaryParse(part7);

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
                        //DisassembleObject(part7, "0");
                        //FillTable();
                        //this.Activate();
                        break;
                    }
                case Kompas6Constants.DocumentTypeEnum.ksDocumentAssembly:
                    {
                        foreach (IPart7 item in part7.Parts)
                        {
                            ksPart ksPart = kompas.TransferInterface(item, 1, 0);
                            if (ksPart.excluded != true)
                            {
                                RecursionK(item, root);
                            }
                        }
                        ProcessTree(root);
                        root.SortTreeNodes(root);
                        root.ReplaceMaterial();
                        FillCodeMaterial(root);
                        FillCodeEquip(root);
                        AddControl(root);

                        this.Activate();
                        toolStripLabel1.Text = String.Empty;
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

        private ObjectAssemblyKompas PrimaryParse(IPart7 part7)
        {
            var ObjectKompas = new ObjectAssemblyKompas();
            ObjectKompas.FullName = part7.FileName;
            ObjectKompas.Designation = part7.Marking;
            ObjectKompas.Name = part7.Name;
            ObjectKompas.Quantity++;
            ObjectKompas.IsLocal = part7.IsLocal;
            if (ObjectKompas.IsLocal == true)
            {
                IPropertyMng propertyMng = (IPropertyMng)application;
                var properties = propertyMng.GetProperties(document3D);
                IPropertyKeeper propertyKeeper = (IPropertyKeeper)part7;
                foreach (IProperty item in properties)
                {
                    if (item.Name == "Раздел спецификации")
                    {
                        dynamic info;
                        bool source;
                        propertyKeeper.GetPropertyValue((_Property)item, out info, false, out source);
                        ObjectKompas.SpecificationSection = info;
                    }
                }
            }
            return ObjectKompas;
        }

        private void RecursionK(IPart7 Part, ObjectAssemblyKompas parent)
        {
            ObjectAssemblyKompas objectF = parent.FindChild(Part.Marking, Part.Name);

            if (objectF != null)
            {
                parent.Quantity++;
            }
            else
            {
                var objectAssemblyKompas = PrimaryParse(Part);
                objectAssemblyKompas.ParentK = parent;
                parent.AddChild(objectAssemblyKompas);

                //DisassembleObject(Part, ParentName);
                if (objectAssemblyKompas.Designation != "" || objectAssemblyKompas.Designation != String.Empty)//заглушка не добавлять детей для детелей у которых нет обозначения
                {
                    foreach (IPart7 item in Part.Parts)
                    {
                        ksPart ksPart2 = kompas.TransferInterface(item, 1, 0);
                        if (ksPart2.excluded != true)
                        {
                            if (item.Detail) objectAssemblyKompas.AddChild(PrimaryParse(item));
                            else RecursionK(item, objectAssemblyKompas);
                        }
                    }
                }
            }
        }

        public void ProcessTree(ObjectAssemblyKompas root)
        {
            if (root == null)
                return;

            // Шаг 1: подсчёт общего количества элементов в дереве
            int totalItems = CountNodes(root);
            int currentIndex = 0; // текущий номер обрабатываемого элемента

            Stack<ObjectAssemblyKompas> stack = new Stack<ObjectAssemblyKompas>();
            stack.Push(root);

            while (stack.Count > 0)
            {
                ObjectAssemblyKompas current = stack.Pop();

                // Обработка текущего узла
                if (current.IsLocal != true)
                {
                    current = ParseObjectKompas(current);
                    currentIndex++; // увеличиваем при обработке каждого узла
                                    // Шаг 2: обновление текста с прогрессом
                    toolStripLabel1.Text = $"Обрабатывается: ({currentIndex} из {totalItems}) {current.Designation} - {current.Name}";
                }

                // Добавляем детей в стек в обратном порядке, чтобы обработка шла в правильном порядке
                if (current.Children != null && current.Children.Count > 0)
                {
                    for (int i = current.Children.Count - 1; i >= 0; i--)
                    {
                        stack.Push(current.Children[i]);
                    }
                }
            }
        }

        // Вспомогательный метод для подсчёта всех узлов в дереве
        private int CountNodes(ObjectAssemblyKompas node)
        {
            if (node == null)
                return 0;

            int count = 1; // считать текущий узел
            if (node.Children != null && node.Children.Count > 0)
            {
                foreach (var child in node.Children)
                {
                    count += CountNodes(child);
                }
            }
            return count;
        }

        private ObjectAssemblyKompas ParseObjectKompas(ObjectAssemblyKompas ObjectKompas)
        {
            IDocuments document = application.Documents;
            OpenDocumentParam openDocumentParam = document.GetOpenDocumentParam();
            openDocumentParam.ReadOnly = false;
            openDocumentParam.Visible = true;
            IKompasDocument kompasDocument = document.OpenDocument(ObjectKompas.FullName, openDocumentParam);
            IKompasDocument3D kompasDocument3D = (IKompasDocument3D)kompasDocument;
            IPart7 part7 = kompasDocument3D.TopPart;
            #region Вытаскиваю текстовые поля
            IPropertyMng propertyMng = (IPropertyMng)application;
            var properties = propertyMng.GetProperties(document3D);
            IPropertyKeeper propertyKeeper = (IPropertyKeeper)part7;
            foreach (IProperty item in properties)
            {
                //if (item.Name == "Наименование")
                //{
                //    dynamic info;
                //    bool source;
                //    propertyKeeper.GetPropertyValue((_Property)item, out info, false, out source);
                //    ObjectKompas.Name = info;
                //}
                //if (item.Name == "Обозначение")
                //{
                //    dynamic info;
                //    bool source;
                //    propertyKeeper.GetPropertyValue((_Property)item, out info, false, out source);
                //    ObjectKompas.Designation = info;
                //}
                if (item.Name == "Материал")
                {
                    dynamic info;
                    bool source;
                    propertyKeeper.GetPropertyValue((_Property)item, out info, false, out source);
                    info = info.Replace("$", "");
                    ObjectKompas.Material = info;
                }
                if (item.Name == "Раздел спецификации")
                {
                    dynamic info;
                    bool source;
                    propertyKeeper.GetPropertyValue((_Property)item, out info, false, out source);
                    ObjectKompas.SpecificationSection = info;
                }
                if (item.Name == "Масса")
                {
                    dynamic info;
                    bool source;
                    propertyKeeper.GetPropertyValue((_Property)item, out info, false, out source);
                    ObjectKompas.Mass = Math.Round(info * 1.2, 2);
                }
                if (item.Name == "Покрытие")
                {
                    dynamic info;
                    bool source;
                    propertyKeeper.GetPropertyValue((_Property)item, out info, false, out source);
                    ObjectKompas.Coating = info;
                }
                if (item.Name == "Сварочные работы")
                {
                    dynamic info;
                    bool source;
                    propertyKeeper.GetPropertyValue((_Property)item, out info, false, out source);
                    ObjectKompas.Welding = info;
                }
                if (item.Name == "Слесарные работы")
                {
                    dynamic info;
                    bool source;
                    propertyKeeper.GetPropertyValue((_Property)item, out info, false, out source);
                    ObjectKompas.LocksmithWork = info;
                }
                if (item.Name == "Примечание")
                {
                    dynamic info;
                    bool source;
                    propertyKeeper.GetPropertyValue((_Property)item, out info, false, out source);
                    ObjectKompas.Note = info;
                }
            }
            #endregion

            #region Тут присваиваю свойство куда входит

            if (ObjectKompas.ParentK != null)
            {
                ObjectKompas.Parent = ObjectKompas.ParentK.Designation + " - " + ObjectKompas.ParentK.Name;
                ObjectKompas.TopParent = root.Designation + " - " + root.Name;
            }
            else
            {
                ObjectKompas.Parent = null;
                ObjectKompas.TopParent = null;
            }
            #endregion

            #region Заполняю габаритные размеры и площадь поверхности
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
                    ObjectKompas.OverallDimensions = TemporaryVariable;
                }

                uint bitVector = 0x3;
                ksMassInertiaParam ksMassInertiaParam = ksPart.CalcMassInertiaProperties(bitVector);
                if (ObjectKompas.Coating != null && ObjectKompas.Coating.Contains("Рекуперат"))
                {
                    ObjectKompas.Area = Math.Round(ksMassInertiaParam.F, 2).ToString();
                }
                else { ObjectKompas.Area = Math.Round(ksMassInertiaParam.F / 2, 2).ToString(); }
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
                        ObjectKompas.R = R.ToString();
                        ObjectKompas.V = V.ToString();
                        ObjectKompas.Q = Q.ToString();
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
                        ObjectKompas.PathToDXF = fi.DirectoryName;
                    }
                }
            }
            catch (ArgumentException)
            {
                ObjectKompas.PathToDXF = "";
            }
            #endregion

            #region Расчет времени резки
            bool checkBox;
            using (Settings settings = Settings.Load(Settings.DefaultPathSettings))
            {
                checkBox = settings.CalcLaserCutTime;
            }
            if (ObjectKompas.SpecificationSection == "Детали" && checkBox == true)
            {
                CuttingTimeCalculation(ObjectKompas);
            }
            #endregion

            if (ObjectKompas.ParentK != null)
            {
                kompasDocument.Close(DocumentCloseOptions.kdDoNotSaveChanges);
            }
            return ObjectKompas;
        }

        public void FillCodeMaterial(ObjectAssemblyKompas node)
        {
            using (Settings settings = Settings.Load(Settings.DefaultPathSettings))
            {
                CodeMaterial codeMaterial = CodeMaterial.Load(settings.PathDictionaryMaterials);
                if (node == null)
                    return;
                // Проверка условия: SpecificationSection равен "Детали" или "Прочие изделия" и Material не null
                if (node.SpecificationSection == "Детали" || node.SpecificationSection == "Прочие изделия" && node.Material != null)
                {
                    if (codeMaterial.Keys.Any(kvp => kvp.Values.Contains(node.Material)))
                    {
                        node.CodeMaterial = codeMaterial.Keys.First(kvp => kvp.Values.Contains(node.Material)).Key;
                    }
                }
                // Рекурсивный обход детей
                if (node.Children != null && node.Children.Any())
                {
                    foreach (var child in node.Children)
                    {
                        FillCodeMaterial(child);
                    }
                }
            }
        }

        public void FillCodeEquip(ObjectAssemblyKompas node)
        {
            using (Settings settings = Settings.Load(Settings.DefaultPathSettings))
            {
                CodeEquip codeEquip = CodeEquip.Load(settings.PathDictionaryEquipment);
                if (node == null)
                    return;

                if (node.SpecificationSection == "Стандартные изделия" || node.SpecificationSection == "Прочие изделия")
                {
                    if (codeEquip.Keys.Any(kvp => kvp.Values.Contains(node.Name)))
                    {
                        node.CodeEquipment = codeEquip.Keys.First(kvp => kvp.Values.Contains(node.Name)).Key;
                    }
                }
                // Рекурсивный обход детей
                if (node.Children != null && node.Children.Any())
                {
                    foreach (var child in node.Children)
                    {
                        FillCodeEquip(child);
                    }
                }
            }
        }

        private void AddControl(ObjectAssemblyKompas objectKompas)
        {
            if (treeListView != null)
            {
                treeListView.Dispose();
            }
            treeListView = new TreeListView
            {
                Dock = DockStyle.Fill,
                FullRowSelect = true,
                UseAlternatingBackColors = true,
                OwnerDraw = true,
            };

            //var treeListView = new BrightIdeasSoftware.TreeListView
            //{
            //    Dock = DockStyle.Fill,
            //    FullRowSelect = true,
            //    UseAlternatingBackColors = true,
            //    OwnerDraw = true,
            //    //CellEditActivation = ObjectListView.CellEditActivateMode.DoubleClick
            //};
            this.Controls.Add(treeListView);
            treeListView.GridLines = true;
            treeListView.AllowColumnReorder = true;

            // Колонки            
            var colDesignation = new OLVColumn("Обозначение", "Designation") { Width = 200 };
            var colName = new OLVColumn("Наименование", "Name") { Width = 200 };
            var colQuantity = new OLVColumn("Кол-во", "Quantity") { Width = 50 };
            //colQuantity.TextAlign = HorizontalAlignment.Center;
            var colSpecSec = new OLVColumn("Раздел спецификации", "SpecificationSection") { Width = 120 };
            var colMaterial = new OLVColumn("Материал", "Material") { Width = 150 };
            var colMass = new OLVColumn("Масса", "Mass") { Width = 50 };
            var colR = new OLVColumn("R", "R") { Width = 50 };
            var colV = new OLVColumn("V", "V") { Width = 50 };
            var colQ = new OLVColumn("Q", "Q") { Width = 50 };
            var colParent = new OLVColumn("Узел-1", "Parent") { Width = 50 };
            var colTopParent = new OLVColumn("Узел верхний", "TopParent") { Width = 50 };
            var colFullName = new OLVColumn("Путь до файла", "FullName") { Width = 50 };
            var colPathToDXF = new OLVColumn("Путь до DXF", "PathToDXF") { Width = 200 };
            var colOverallDimensions = new OLVColumn("Габаритные размеры", "OverallDimensions") { Width = 100 };
            var colCoating = new OLVColumn("Покрытие", "Coating") { Width = 80 };
            var colWelding = new OLVColumn("Сварочные работы", "Welding") { Width = 100 };
            var colLocksmithWork = new OLVColumn("Слесарные работы", "LocksmithWork") { Width = 80 };
            var colNote = new OLVColumn("Примечание", "Note") { Width = 80 };
            var colArea = new OLVColumn("Площадь поверхности", "Area") { Width = 80 };
            var colCodeEquipment = new OLVColumn("Код СИ", "CodeEquipment") { Width = 50 };
            var colCodeMaterial = new OLVColumn("Код Мат", "CodeMaterial") { Width = 50 };
            var colTimeCut = new OLVColumn("Время резки", "TimeCut") { Width = 50 };
            var colDxfDimensions = new OLVColumn("Габариты DXF", "DxfDimensions") { Width = 80 };

            treeListView.AllColumns.AddRange(new[] { colDesignation,
                                                     colName,
                                                     colQuantity,
                                                     colSpecSec,
                                                     colMaterial,
                                                     colMass,
                                                     colR,
                                                     colV,
                                                     colQ,
                                                     colParent,
                                                     colTopParent,
                                                     colFullName,
                                                     colPathToDXF,
                                                     colOverallDimensions,
                                                     colCoating,
                                                     colWelding,
                                                     colLocksmithWork,
                                                     colNote,
                                                     colArea,
                                                     colCodeEquipment,
                                                     colCodeMaterial,
                                                     colTimeCut,
                                                     colDxfDimensions
                                                     });
            treeListView.RebuildColumns();

            // Обработка видимости наличия детей
            treeListView.CanExpandGetter = model =>
            {
                var itemModel = model as ObjectAssemblyKompas;
                return itemModel != null && itemModel.Children != null && itemModel.Children.Count > 0;
            };
            treeListView.ChildrenGetter = model => (model as ObjectAssemblyKompas).Children;

            // Данные
            treeListView.Roots = new List<ObjectAssemblyKompas> { objectKompas };

            string tempPath = System.Reflection.Assembly.GetExecutingAssembly().Location.ToString();
            //UpdateExistingColumns(tempPath.Remove(tempPath.LastIndexOf(@"\")) + @"\" + @"Settings\Сolumns.xml", treeListView);

            treeListView.ExpandAll();

            // Контекстное меню
            contextMenu = new ContextMenuStrip();
            var itemOpenInCompass = new ToolStripMenuItem("Открыть в Компас");
            var itemOpenInExplorer = new ToolStripMenuItem("Открыть в проводнике");
            contextMenu.Items.AddRange(new[] { itemOpenInCompass, itemOpenInExplorer });

            itemOpenInExplorer.Click += (s, e) =>
            {
                if (treeListView.SelectedObject is ObjectAssemblyKompas item)
                {
                    // Предположим, что в свойстве FullName хранится путь к папке
                    FileInfo fi = new FileInfo(item.FullName);
                    if (Directory.Exists(fi.DirectoryName))
                    {
                        Process.Start("explorer.exe", fi.DirectoryName);
                    }
                    else
                    {
                        MessageBox.Show("Путь не существует: " + fi.DirectoryName);
                    }
                }
            };
            itemOpenInCompass.Click += (s, e) =>
            {
                if (treeListView.SelectedObject is ObjectAssemblyKompas item)
                {
                    IDocuments document = application.Documents;
                    document.Open(item.FullName, true, false);
                    string windowTitle = "КОМПАС-3D v22";
                    IntPtr hWnd = FindWindow(null, windowTitle);
                    if (hWnd != IntPtr.Zero)
                    {
                        SetForegroundWindow(hWnd);
                    }
                }
            };
            // Обработка правого клика
            treeListView.MouseDown += (s, e) =>
            {
                if (e.Button == MouseButtons.Right)
                {
                    var hitTest = treeListView.HitTest(e.Location);
                    if (hitTest.Item != null)
                    {
                        treeListView.SelectedObject = hitTest.Item.Focused;
                        contextMenu.Show(treeListView, e.Location);
                    }
                }
            };
        }

        // Рекурсивный метод для обхода узлов и добавления их в таблицу
        private void AddNodeToDataTable(ObjectAssemblyKompas node, DataTable dt)
        {
            if (node == null)
                return;

            // Добавляем текущий узел
            DataRow row = dt.NewRow();
            row["Наименование"] = node.Name;
            row["Обозначение"] = node.Designation;
            row["Кол-во"] = node.Quantity;
            row["Раздел спецификации"] = node.SpecificationSection;
            row["Материал"] = node.Material;
            row["Масса"] = node.Mass;
            row["R"] = node.R;
            row["V"] = node.V;
            row["Q"] = node.Q;
            row["Узел-1"] = node.Parent;
            row["Узел верхний"] = node.TopParent;
            row["Путь до файла"] = node.FullName;
            row["Путь до DXF"] = node.PathToDXF;
            row["Габаритные размеры"] = node.OverallDimensions;
            row["Покрытие"] = node.Coating;
            row["Сварочные работы"] = node.Welding;
            row["Слесарные работы"] = node.LocksmithWork;
            row["Примечание"] = node.Note;
            row["Площадь поверхности"] = node.Area;
            row["Код СИ"] = node.CodeEquipment;
            row["Код Мат"] = node.CodeMaterial;
            row["Время резки"] = node.TimeCut;
            row["Габариты DXF"] = node.DxfDimensions;

            dt.Rows.Add(row);

            // Рекурсивно добавляем детей
            if (node.Children != null)
            {
                foreach (var child in node.Children)
                {
                    AddNodeToDataTable(child, dt);
                }
            }
        }

        private void StripButtonXML_Click(object sender, EventArgs e)
        {
            string pathForXML = System.Reflection.Assembly.GetExecutingAssembly().Location.Remove(System.Reflection.Assembly.GetExecutingAssembly().Location.Length - 16) + root.Designation + " - " + root.Name + ".xml";
            // Настройки
            XmlWriterSettings settings = new XmlWriterSettings
            {
                Indent = true, // делаем читаемый отступ
                Encoding = System.Text.Encoding.UTF8
            };
            using (XmlWriter writer = XmlWriter.Create(pathForXML, settings))
            {
                writer.WriteStartDocument();
                writer.WriteStartElement("Rows"); // корневой элемент

                foreach (var obj in treeListView.Roots)
                {
                    WriteObjectToXml(writer, root);
                }

                writer.WriteEndElement(); // </Rows>
                writer.WriteEndDocument();
            }
        }
        // рекурсивная функция для записи объекта и его детей
        private void WriteObjectToXml(XmlWriter writer, ObjectAssemblyKompas obj)
        {
            writer.WriteStartElement("Row");

            // Заполняем свойства
            WriteElementOrEmpty(writer, "Designation", obj.Designation);
            WriteElementOrEmpty(writer, "Name", obj.Name);
            WriteElementOrEmpty(writer, "Quantity", obj.Quantity.ToString());
            WriteElementOrEmpty(writer, "SpecificationSection", obj.SpecificationSection);
            WriteElementOrEmpty(writer, "Material", obj.Material);
            WriteElementOrEmpty(writer, "Mass", obj.Mass.ToString("F2", System.Globalization.CultureInfo.InvariantCulture));
            WriteElementOrEmpty(writer, "R", obj.R);
            WriteElementOrEmpty(writer, "V", obj.V);
            WriteElementOrEmpty(writer, "Q", obj.Q);
            WriteElementOrEmpty(writer, "Parent", obj.Parent);
            WriteElementOrEmpty(writer, "TopParent", obj.TopParent);
            WriteElementOrEmpty(writer, "FullName", obj.FullName);
            WriteElementOrEmpty(writer, "PathToDXF", obj.PathToDXF);
            WriteElementOrEmpty(writer, "OverallDimensions", obj.OverallDimensions);
            WriteElementOrEmpty(writer, "Coating", obj.Coating);
            WriteElementOrEmpty(writer, "Welding", obj.Welding);
            WriteElementOrEmpty(writer, "LocksmithWork", obj.LocksmithWork);
            WriteElementOrEmpty(writer, "Note", obj.Note);
            WriteElementOrEmpty(writer, "Area", obj.Area);
            WriteElementOrEmpty(writer, "CodeEquipment", obj.CodeEquipment);
            WriteElementOrEmpty(writer, "CodeMaterial", obj.CodeMaterial);
            WriteElementOrEmpty(writer, "TimeCut", obj.TimeCut);
            WriteElementOrEmpty(writer, "DxfDimensions", obj.DxfDimensions);

            writer.WriteEndElement(); // закрываем "Row"

            // Обработка детей
            if (obj.Children != null && obj.Children.Count > 0)
            {
                foreach (var child in obj.Children)
                {
                    WriteObjectToXml(writer, child);
                }
            }
        }

        private void WriteElementOrEmpty(XmlWriter writer, string elementName, string value)
        {
            // Создает тег, даже если значение пустое или null
            // Для null, можно заменить на ""
            string val = value ?? "";
            writer.WriteElementString(elementName, val);
        }

        private void StripButtonExcel_Click(object sender, EventArgs e)
        {
            if (treeListView == null)
                return;

            // Создаем DataTable с нужными колонками
            DataTable dt = new DataTable();

            // Описание колонок, совпадает с колонками treeListView
            dt.Columns.Add("Обозначение", typeof(string));
            dt.Columns.Add("Наименование", typeof(string));
            dt.Columns.Add("Кол-во", typeof(int));
            dt.Columns.Add("Раздел спецификации", typeof(string));
            dt.Columns.Add("Материал", typeof(string));
            dt.Columns.Add("Масса", typeof(double));
            dt.Columns.Add("R", typeof(string));
            dt.Columns.Add("V", typeof(string));
            dt.Columns.Add("Q", typeof(string));
            dt.Columns.Add("Узел-1", typeof(string));
            dt.Columns.Add("Узел верхний", typeof(string));
            dt.Columns.Add("Путь до файла", typeof(string));
            dt.Columns.Add("Путь до DXF", typeof(string));
            dt.Columns.Add("Габаритные размеры", typeof(string));
            dt.Columns.Add("Покрытие", typeof(string));
            dt.Columns.Add("Сварочные работы", typeof(string));
            dt.Columns.Add("Слесарные работы", typeof(string));
            dt.Columns.Add("Примечание", typeof(string));
            dt.Columns.Add("Площадь поверхности", typeof(string));
            dt.Columns.Add("Код СИ", typeof(string));
            dt.Columns.Add("Код Мат", typeof(string));
            dt.Columns.Add("Время резки", typeof(string));
            dt.Columns.Add("Габариты DXF", typeof(string));

            // Рекурсивная функция обхода
            void AddRowFromNode(ObjectAssemblyKompas node)
            {
                if (node == null) return;

                // Создаем новую строку
                var row = dt.NewRow();

                row["Обозначение"] = node.Designation ?? "";
                row["Наименование"] = node.Name ?? "";
                row["Кол-во"] = node.Quantity;
                row["Раздел спецификации"] = node.SpecificationSection ?? "";
                row["Материал"] = node.Material ?? "";
                row["Масса"] = node.Mass;
                row["R"] = node.R ?? "";
                row["V"] = node.V ?? "";
                row["Q"] = node.Q ?? "";
                row["Узел-1"] = node.Parent ?? "";
                row["Узел верхний"] = node.TopParent ?? "";
                row["Путь до файла"] = node.FullName ?? "";
                row["Путь до DXF"] = node.PathToDXF ?? "";
                row["Габаритные размеры"] = node.OverallDimensions ?? "";
                row["Покрытие"] = node.Coating ?? "";
                row["Сварочные работы"] = node.Welding ?? "";
                row["Слесарные работы"] = node.LocksmithWork ?? "";
                row["Примечание"] = node.Note ?? "";
                row["Площадь поверхности"] = node.Area ?? "";
                row["Код СИ"] = node.CodeEquipment ?? "";
                row["Код Мат"] = node.CodeMaterial ?? "";
                row["Время резки"] = node.TimeCut ?? "";
                row["Габариты DXF"] = node.DxfDimensions ?? "";

                dt.Rows.Add(row);

                // Обработка детей
                if (node.Children != null)
                {
                    foreach (var child in node.Children)
                    {
                        AddRowFromNode(child);
                    }
                }
            }

            // Основная часть: обход корней
            foreach (ObjectAssemblyKompas root in treeListView.Roots)
            {
                AddRowFromNode(root);
            }

            XLWorkbook excelWorkbook = new XLWorkbook();
            string pathForExcel = System.Reflection.Assembly.GetExecutingAssembly().Location.Remove(System.Reflection.Assembly.GetExecutingAssembly().Location.Length - 16);
            IXLWorksheet worksheet = excelWorkbook.Worksheets.Add(dt, "Отчет");

            // Записываем заголовки из DataTable в первую строку
            for (int colIndex = 0; colIndex < dt.Columns.Count; colIndex++)
            {
                worksheet.Cell(1, colIndex + 1).Value = dt.Columns[colIndex].ColumnName;
            }

            // Настройка стилей
            worksheet.RangeUsed().Style.NumberFormat.Format = "@";
            worksheet.Outline.SummaryVLocation = XLOutlineSummaryVLocation.Top;

            // Создаем таблицу
            IXLTable xLTable = worksheet.Table(0);
            xLTable.Theme = XLTableTheme.TableStyleLight8;

            Dictionary<string, List<int>> groups = new Dictionary<string, List<int>>();
            // Получаем заголовки из первой строки листа
            var headerRow = worksheet.Row(1);
            // Количество колонок
            int totalColumns = worksheet.ColumnsUsed().Count();

            // Создаем список колонок
            List<string> columnNames = new List<string>();
            for (int col = 1; col <= totalColumns; col++)
            {
                var cellValue = headerRow.Cell(col).GetString();
                columnNames.Add(cellValue);
            }

            // Обрабатываем строки данных начиная со 3-й (индекс 3)
            int lastDataRow = worksheet.LastRowUsed().RowNumber();

            for (int rowIndex = 3; rowIndex <= lastDataRow; rowIndex++)
            {
                var row = worksheet.Row(rowIndex);

                // Получение значения "Parent" из соответствующей колонки
                string parentValue = "";
                int parentColIndex = columnNames.IndexOf("Узел-1") + 1; // +1 потому что нумерация LaTeX с 1

                if (parentColIndex > 0)
                {
                    var val = row.Cell(parentColIndex).Value;
                    parentValue = val != null ? val.ToString() : "";
                }

                // Заполняем данные в Excel (если они еще не заполнены)
                // Предположим, что данные уже есть
                // Если еще нужно заполнять, то делаете это здесь аналогично вашему коду

                // Добавляем индекс строки к группе
                if (!groups.ContainsKey(parentValue))
                    groups[parentValue] = new List<int>();

                // В Excel строки начинаются с 2, и так как мы в 1-based, то rowIndex - текущая строка
                groups[parentValue].Add(rowIndex);
            }

            // Группируем строки по группам
            foreach (var group in groups.Values)
            {
                if (group.Count > 1)
                {
                    int startRow = group[0];
                    int endRow = group[group.Count - 1];
                    worksheet.Rows(startRow, endRow).Group();
                }
            }

            // Автоматическая подгонка ширины колонок
            worksheet.Columns().AdjustToContents();

            excelWorkbook.SaveAs(pathForExcel + root.Designation + " - " + root.Name + ".xlsx");

            // Теперь у вас есть DataTable dt с данными из treeListView
            // Можно, например, вывести его, открыть диалог с отчетом, привязать к GridView и т.д.            
        }

        private void toolStripOpenExplorer_Click(object sender, EventArgs e)
        {
            string tempPath = System.Reflection.Assembly.GetExecutingAssembly().Location.ToString();
            Process.Start(tempPath.Remove(tempPath.LastIndexOf(@"\")));
        }
    }
}
