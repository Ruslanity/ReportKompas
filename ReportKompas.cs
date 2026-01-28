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
using Microsoft.WindowsAPICodePack.Shell;
using reference = System.Int32;

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

        private void CollapseAllToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (treeListView != null)
            {
                treeListView.CollapseAll();
            }
        }

        private void ExpandAllToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (treeListView != null)
            {
                treeListView.ExpandAll();
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
                                // Пропускаем элементы зелёного цвета (0x00FF00)
                                //IColorParam7 itemColorParam = (IColorParam7)item;
                                //if (itemColorParam.Color == 0x00FF00)
                                //    continue;

                                if (item.RevealComposition)
                                {
                                    // Раскрываем состав: добавляем дочерние элементы напрямую к root
                                    foreach (IPart7 childItem in item.Parts)
                                    {
                                        ksPart childKsPart = kompas.TransferInterface(childItem, 1, 0);
                                        if (childKsPart.excluded != true)
                                        {
                                            // Пропускаем элементы зелёного цвета (0x00FF00)
                                            //IColorParam7 childColorParam = (IColorParam7)childItem;
                                            //if (childColorParam.Color == 0x00FF00)
                                            //    continue;

                                            RecursionK(childItem, root);
                                        }
                                    }
                                }
                                else
                                {
                                    RecursionK(item, root);
                                }
                            }
                        }
                        ProcessTree(root);
                        root.SortTreeNodes(root);
                        root.ReplaceMaterial();
                        FillCodeMaterial(root);
                        FillCodeEquip(root);
                        ReorganizeElements(root);
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
            // Пропускаем детали зелёного цвета (0x00FF00)
            IColorParam7 colorCheck = (IColorParam7)part7;
            if (colorCheck.Color == 0x00FF00)
                return null;

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
            IColorParam7 colorParam = (IColorParam7)part7;
            if (colorParam.Color == 0xFF33FF)
            {
                ObjectKompas.IsFastener = "true";
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
                if (objectAssemblyKompas == null)
                    return;

                objectAssemblyKompas.ParentK = parent;
                parent.AddChild(objectAssemblyKompas);

                if (objectAssemblyKompas.Designation != "" || objectAssemblyKompas.Designation != String.Empty)//заглушка не добавлять детей для детелей у которых нет обозначения
                {
                    foreach (IPart7 item in Part.Parts)
                    {
                        ksPart ksPart2 = kompas.TransferInterface(item, 1, 0);
                        if (ksPart2.excluded != true)
                        {
                            if (item.RevealComposition)
                            {
                                // Раскрываем состав: добавляем дочерние элементы напрямую к текущему родителю
                                foreach (IPart7 childItem in item.Parts)
                                {
                                    ksPart childKsPart = kompas.TransferInterface(childItem, 1, 0);
                                    if (childKsPart.excluded != true)
                                    {
                                        if (childItem.Detail) objectAssemblyKompas.AddChild(PrimaryParse(childItem));
                                        else RecursionK(childItem, objectAssemblyKompas);
                                    }
                                }
                            }
                            else
                            {
                                if (item.Detail) objectAssemblyKompas.AddChild(PrimaryParse(item));
                                else RecursionK(item, objectAssemblyKompas);
                            }
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
            var properties = propertyMng.GetProperties(kompasDocument3D);
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
                if (item.Name == "IsPainted")
                {
                    dynamic info;
                    bool source = true;
                    propertyKeeper.GetPropertyValue((_Property)item, out info, false, out source);
                    // Преобразуем значение в строку для единообразия
                    if (info != null)
                    {
                        ObjectKompas.IsPainted = info.ToString();
                    }
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
                if (item.Name == "Комплект крепежа")
                {
                    dynamic info;
                    bool source;
                    propertyKeeper.GetPropertyValue((_Property)item, out info, false, out source);
                    ObjectKompas.IsFastener = info;
                }
            }
            #endregion

            #region Чтение атрибута "Технологический маршрут"
            try
            {
                ksDocument3D ksDoc3D = kompas.TransferInterface(kompasDocument3D, 1, 0) as ksDocument3D;
                reference docRef = ksDoc3D.reference;
                ksAttributeObject attrObj = (ksAttributeObject)kompas.GetAttributeObject();

                if (attrObj != null && docRef != 0)
                {
                    string attrLibraryFile = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "document_attr.lat");

                    // Создаем итератор атрибутов
                    ksIterator iter = (ksIterator)kompas.GetIterator();
                    if (iter != null && iter.ksCreateAttrIterator(docRef, 0, 0, 0, 0, 0))
                    {
                        // Получаем первый атрибут
                        reference ownerRef = docRef;
                        reference pAttr = iter.ksMoveAttrIterator("F", ref ownerRef);

                        while (pAttr != 0)
                        {
                            try
                            {
                                // Получаем информацию о ключах атрибута и его тип
                                int k1 = 0, k2 = 0, k3 = 0, k4 = 0;
                                double attrTypeNum = 0;
                                attrObj.ksGetAttrKeysInfo(pAttr, out k1, out k2, out k3, out k4, out attrTypeNum);

                                // Получаем тип атрибута для извлечения имени
                                ksAttributeTypeParam type = (ksAttributeTypeParam)kompas.GetParamStruct((short)Kompas6Constants.StructType2DEnum.ko_AttributeType);
                                if (type != null)
                                {
                                    type.Init();
                                    int getTypeResult = attrObj.ksGetAttrType(attrTypeNum, attrLibraryFile, type);

                                    if (getTypeResult == 1)
                                    {
                                        string attrName = type.header;

                                        if (attrName == "Технологический маршрут")
                                        {
                                            // Читаем значение атрибута
                                            ksUserParam usPar = (ksUserParam)kompas.GetParamStruct((short)Kompas6Constants.StructType2DEnum.ko_UserParam);
                                            ksLtVariant item = (ksLtVariant)kompas.GetParamStruct((short)Kompas6Constants.StructType2DEnum.ko_LtVariant);
                                            ksDynamicArray arr = (ksDynamicArray)kompas.GetDynamicArray(23);

                                            if (usPar != null && item != null && arr != null)
                                            {
                                                usPar.Init();
                                                usPar.SetUserArray(arr);

                                                // Добавляем пустое значение в массив
                                                item.Init();
                                                item.strVal = string.Empty;
                                                arr.ksAddArrayItem(-1, item);

                                                // Читаем строку атрибута
                                                attrObj.ksGetAttrRow(pAttr, 0, 0, 0, usPar);

                                                // Получаем массив после чтения
                                                ksDynamicArray readArr = (ksDynamicArray)usPar.GetUserArray();

                                                if (readArr != null && readArr.ksGetArrayCount() > 0)
                                                {
                                                    item.Init();
                                                    int getItemResult = readArr.ksGetArrayItem(0, item);

                                                    if (getItemResult == 1)
                                                    {
                                                        ObjectKompas.TechnologicalRoute = item.strVal ?? "";
                                                    }
                                                }

                                                // Очищаем массив
                                                arr.ksDeleteArray();
                                            }

                                            break; // Нашли нужный атрибут, выходим из цикла
                                        }
                                    }
                                }
                            }
                            catch (Exception)
                            {
                                // Игнорируем ошибки при обработке отдельного атрибута
                            }

                            // Переходим к следующему атрибуту
                            ownerRef = docRef;
                            pAttr = iter.ksMoveAttrIterator("N", ref ownerRef);
                        }
                    }
                }
            }
            catch (Exception)
            {
                // Если не удалось прочитать атрибут, оставляем поле пустым
                ObjectKompas.TechnologicalRoute = null;
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
                ObjectKompas.Area = Math.Round(ksMassInertiaParam.F, 2).ToString();
                //if (ObjectKompas.Coating != null && ObjectKompas.Coating.Contains("Рекуперат"))
                //{
                //    ObjectKompas.Area = Math.Round(ksMassInertiaParam.F, 2).ToString();
                //}
                //else { ObjectKompas.Area = Math.Round(ksMassInertiaParam.F / 2, 2).ToString(); }
            }
            #endregion

            #region Получение превью-изображения из 3D модели
            try
            {
                Bitmap originalImage = null;
                bool useKompasApiFallback = true; // Принудительно использовать KOMPAS API вместо Shell
                bool imageFromShell = false;

                // Метод 1: Попытка получить миниатюру через Windows Shell (быстрый способ)
                // Windows кэширует миниатюры, поэтому повторные запросы будут очень быстрыми
                try
                {
                    if (File.Exists(ObjectKompas.FullName))
                    {
                        using (ShellFile shellFile = ShellFile.FromFilePath(ObjectKompas.FullName))
                        {
                            // ExtraLargeBitmap = 256x256, LargeBitmap = 96x96
                            Bitmap shellThumbnail = shellFile.Thumbnail.ExtraLargeBitmap;
                            if (shellThumbnail != null && shellThumbnail.Width > 1 && shellThumbnail.Height > 1)
                            {
                                // Создаём копию, т.к. оригинал будет удалён при dispose ShellFile
                                originalImage = new Bitmap(shellThumbnail);
                                imageFromShell = true;
                            }
                            else
                            {
                                useKompasApiFallback = true;
                            }
                        }
                    }
                    else
                    {
                        useKompasApiFallback = true;
                    }
                }
                catch
                {
                    // Если Shell не сработал, используем KOMPAS API
                    useKompasApiFallback = true;
                }

                // Метод 2: Fallback на KOMPAS API если Shell не сработал
                if (useKompasApiFallback || originalImage == null)
                {
                    // Сбрасываем Shell-изображение, чтобы использовать только KOMPAS API
                    if (originalImage != null)
                    {
                        originalImage.Dispose();
                        originalImage = null;
                    }

                    string tempImagePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString() + ".png");

                    try
                    {
                        // Получаем ksDocument3D для использования метода SaveAsToRasterFormat (API v5)
                        ksDocument3D ksDoc3D = kompas.TransferInterface(kompasDocument3D, 1, 0) as ksDocument3D;
                        if (ksDoc3D != null)
                        {
                            ksDoc3D.hideAllAuxiliaryGeom = true;
                            ksDoc3D.drawMode = 1; // 1=каркас, 2=полутоновое, 3=полутоновое с каркасом нужно ставить 1

                            ksViewProjectionCollection ksViewProjectionCollection = ksDoc3D.GetViewProjectionCollection();
                            if (ksViewProjectionCollection != null)
                            {
                                ksViewProjectionCollection.GetByIndex(7).SetCurrent();
                            }

                            // Получаем параметры для сохранения в растровый формат
                            ksRasterFormatParam rasterParam = (ksRasterFormatParam)ksDoc3D.RasterFormatParam();
                            if (rasterParam != null)
                            {
                                rasterParam.Init();
                                rasterParam.format = 2;                           // PNG формат
                                rasterParam.colorType = 1;                        // Цветное изображение
                                rasterParam.extResolution = 72;                   // Разрешение 72 DPI

                                // Динамический масштаб на основе габаритных размеров детали (OverallDimensions)
                                // Формат строки: "100х1008х50" или "100x1008x50" (через русскую "х" или латинскую "x")
                                // extScale в KOMPAS: количество пикселей на 1 мм при 72 DPI
                                // При extScale=1: 1 мм = 1 пиксель, деталь 100мм = 100 пикселей
                                double scale = 15.0; // масштаб по умолчанию для мелких деталей
                                if (!string.IsNullOrEmpty(ObjectKompas.OverallDimensions))
                                {
                                    // Разбиваем строку по разделителям "х" (рус) и "x" (лат)
                                    string[] parts = ObjectKompas.OverallDimensions.Split(new char[] { 'х', 'x', 'Х', 'X' }, StringSplitOptions.RemoveEmptyEntries);
                                    double maxDimension = 0;
                                    foreach (string part in parts)
                                    {
                                        double dim;
                                        if (double.TryParse(part.Trim(), NumberStyles.Any, CultureInfo.InvariantCulture, out dim))
                                        {
                                            if (dim > maxDimension)
                                                maxDimension = dim;
                                        }
                                    }
                                    // Масштаб: целевой размер изображения / максимальный габарит детали
                                    // Цель: итоговое изображение ~500 пикселей по большей стороне
                                    if (maxDimension > 0)
                                    {
                                        double targetPixels = 500.0;
                                        scale = targetPixels / maxDimension;
                                        // Ограничиваем масштаб: минимум 0.1 (для очень больших деталей), максимум 15 (для мелких)
                                        scale = Math.Max(0.1, Math.Min(scale, 15.0));
                                    }
                                }
                                rasterParam.extScale = scale;
                                rasterParam.greyScale = true;                     // Градации серого
                                rasterParam.colorBPP = 0x18;                      // 24 бита на пиксель (RGB)
                                rasterParam.onlyThinLine = true;                  // Нормальная толщина линий

                                bool result = ksDoc3D.SaveAsToRasterFormat(tempImagePath, rasterParam);

                                // Поиск файла (KOMPAS может сохранить с другим расширением)
                                string fileNameWithoutExt = Path.GetFileNameWithoutExtension(tempImagePath);
                                string tempDir = Path.GetDirectoryName(tempImagePath);
                                string[] extensions = { ".png", ".bmp", ".jpg", ".jpeg", ".gif", ".tif", ".tiff" };
                                string actualFilePath = tempImagePath;
                                foreach (string ext in extensions)
                                {
                                    string altPath = Path.Combine(tempDir, fileNameWithoutExt + ext);
                                    if (File.Exists(altPath))
                                    {
                                        actualFilePath = altPath;
                                        break;
                                    }
                                }

                                if (result && File.Exists(actualFilePath))
                                {
                                    // Загружаем изображение с уменьшением размера для избежания переполнения памяти
                                    using (FileStream fs = new FileStream(actualFilePath, FileMode.Open, FileAccess.Read))
                                    {
                                        using (Bitmap fullImage = new Bitmap(fs))
                                        {
                                            // Ограничиваем максимальный размер изображения
                                            int maxPreviewSize = 500;
                                            int newWidth, newHeight;

                                            if (fullImage.Width > maxPreviewSize || fullImage.Height > maxPreviewSize)
                                            {
                                                float ratio = Math.Min((float)maxPreviewSize / fullImage.Width, (float)maxPreviewSize / fullImage.Height);
                                                newWidth = (int)(fullImage.Width * ratio);
                                                newHeight = (int)(fullImage.Height * ratio);
                                            }
                                            else
                                            {
                                                newWidth = fullImage.Width;
                                                newHeight = fullImage.Height;
                                            }

                                            // Создаём уменьшенную копию
                                            originalImage = new Bitmap(newWidth, newHeight);
                                            using (Graphics g = Graphics.FromImage(originalImage))
                                            {
                                                g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                                                g.DrawImage(fullImage, 0, 0, newWidth, newHeight);
                                            }
                                        }
                                    }
                                }
                            }

                            ksDoc3D.shadedWireframe = true;
                        }
                    }
                    finally
                    {
                        // Удаляем временный файл
                        try
                        {
                            if (File.Exists(tempImagePath))
                                File.Delete(tempImagePath);
                        }
                        catch { }
                    }
                }

                // Обработка полученного изображения (масштабирование и конвертация)
                if (originalImage != null)
                {
                    try
                    {
                        // Рассчитываем размеры с сохранением пропорций
                        int maxSize = 150;
                        int padding = 5; // Отступ от краёв, чтобы линии не обрезались
                        int availableSize = maxSize - (padding * 2); // Доступная область для изображения

                        float aspectRatio = (float)originalImage.Width / originalImage.Height;
                        int newWidth, newHeight;

                        if (originalImage.Width > originalImage.Height)
                        {
                            newWidth = availableSize;
                            newHeight = (int)(availableSize / aspectRatio);
                        }
                        else
                        {
                            newHeight = availableSize;
                            newWidth = (int)(availableSize * aspectRatio);
                        }

                        // Создаем квадратный холст для центрирования изображения
                        using (Bitmap resizedImage = new Bitmap(maxSize, maxSize))
                        {
                            using (Graphics graphics = Graphics.FromImage(resizedImage))
                            {
                                // Заливаем фон белым цветом
                                graphics.Clear(Color.White);

                                graphics.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                                graphics.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.HighQuality;
                                graphics.PixelOffsetMode = System.Drawing.Drawing2D.PixelOffsetMode.HighQuality;

                                // Центрируем изображение на холсте с учётом отступов
                                int x = (maxSize - newWidth) / 2;
                                int y = (maxSize - newHeight) / 2;

                                graphics.DrawImage(originalImage, x, y, newWidth, newHeight);
                            }

                            // Пост-обработка: все не-белые пиксели делаем черными для контрастности
                            for (int py = 0; py < resizedImage.Height; py++)
                            {
                                for (int px = 0; px < resizedImage.Width; px++)
                                {
                                    Color pixelColor = resizedImage.GetPixel(px, py);
                                    // Если пиксель не чисто белый (с учетом небольшого порога)
                                    if (pixelColor.R < 250 || pixelColor.G < 250 || pixelColor.B < 250)
                                    {
                                        resizedImage.SetPixel(px, py, Color.Black);
                                    }
                                }
                            }

                            // Конвертируем в массив байтов
                            using (MemoryStream ms = new MemoryStream())
                            {
                                resizedImage.Save(ms, System.Drawing.Imaging.ImageFormat.Png);
                                ObjectKompas.PreviewImage = ms.ToArray();
                            }
                        }
                    }
                    finally
                    {
                        originalImage.Dispose();
                    }
                }
            }
            catch (Exception ex)
            {
                // Если не удалось получить изображение, оставляем поле пустым
                ObjectKompas.PreviewImage = null;
                System.Diagnostics.Debug.WriteLine("Ошибка получения превью для " + ObjectKompas.Designation + ": " + ex.Message);
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
                UseAlternatingBackColors = false,
                OwnerDraw = true,
                UseCellFormatEvents = true,
            };

            this.Controls.Add(treeListView);
            // Отправляем TreeListView на задний план Z-order, чтобы ToolStrip
            // с Dock.Bottom корректно занял своё место и TreeListView заполнил остаток
            treeListView.SendToBack();
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
            var colIsPainted = new OLVColumn("IsPainted", "IsPainted") { Width = 50 };
            var colCoating = new OLVColumn("Покрытие", "Coating") { Width = 80 };
            var colCoverageArea = new OLVColumn("Площадь покрытия", "CoverageArea") { Width = 100 };
            var colWelding = new OLVColumn("Сварочные работы", "Welding") { Width = 100 };
            var colLocksmithWork = new OLVColumn("Слесарные работы", "LocksmithWork") { Width = 80 };
            var colTechnologicalRoute = new OLVColumn("Технологический маршрут", "TechnologicalRoute") { Width = 80 };
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
                                                     colIsPainted,
                                                     colCoating,
                                                     colCoverageArea,
                                                     colWelding,
                                                     colLocksmithWork,
                                                     colTechnologicalRoute,
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

            // Применяем подсветку строк ДО установки данных
            ApplyRowHighlighting();

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
                    if (File.Exists(item.FullName))
                    {
                        // Открываем проводник с выделенным файлом
                        Process.Start("explorer.exe", "/select,\"" + item.FullName + "\"");
                    }
                    else
                    {
                        MessageBox.Show("Файл не существует: " + item.FullName);
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

        /// <summary>
        /// Подсвечивает строки в TreeListView:
        /// - красным: пустой Раздел спецификации
        /// - оранжевым: раздел "Стандартные изделия" и пустой Код СИ
        /// - оранжевым: раздел "Детали" и пустой Код Мат
        /// </summary>
        public void ApplyRowHighlighting()
        {
            if (treeListView == null)
                return;

            treeListView.FormatRow += (sender, e) =>
            {
                var item = e.Model as ObjectAssemblyKompas;
                if (item == null)
                    return;

                // Проверяем условие: пустой Раздел спецификации (приоритет - красный)
                if (string.IsNullOrEmpty(item.SpecificationSection))
                {
                    e.Item.BackColor = Color.Tomato;
                    return;
                }

                // Проверяем условие: раздел "Стандартные изделия" и пустой Код СИ
                bool isStandardWithoutCode = item.SpecificationSection == "Стандартные изделия" &&
                    string.IsNullOrEmpty(item.CodeEquipment);

                // Проверяем условие: раздел "Детали" и пустой Код Мат
                bool isDetailWithoutMaterial = item.SpecificationSection == "Детали" &&
                    string.IsNullOrEmpty(item.CodeMaterial);

                if (isStandardWithoutCode || isDetailWithoutMaterial)
                {
                    e.Item.BackColor = Color.Orange;
                }
            };
        }

        // Метод для формирования node "Комплект крепежа"
        // Собирает элементы с IsFastener == "true", удаляет их из родителей
        // и добавляет новый узел "Комплект крепежа" в Children переданного node
        public void ReorganizeElements(ObjectAssemblyKompas node)
        {
            if (node == null)
                return;

            string parentValue = node.Designation + " - " + node.Name;

            // Создаём узел "Комплект крепежа"
            var fastenerKit = new ObjectAssemblyKompas
            {
                Name = "Комплект крепежа " + node.Designation,
                Quantity = 1,
                Parent = parentValue,
                TopParent = parentValue
            };

            // Рекурсивно собираем крепёжные элементы, удаляем их у родителей и добавляем в fastenerKit

            CollectAndRemoveFasteners(node, fastenerKit, parentValue);

            // Если крепёжные элементы найдены, добавляем "Комплект крепежа" в Children переданного node
            if (fastenerKit.Children != null && fastenerKit.Children.Count > 0)
            {
                node.AddChild(fastenerKit);
            }
        }

        // Вспомогательный рекурсивный метод для сбора и удаления крепёжных элементов
        private void CollectAndRemoveFasteners(ObjectAssemblyKompas node, ObjectAssemblyKompas fastenerKit, string parentValue)
        {
            if (node == null || node.Children == null)
                return;

            // Создаём копию списка для безопасной итерации при удалении
            var childrenCopy = node.Children.ToList();

            foreach (var child in childrenCopy)
            {
                // Проверяем, является ли элемент крепёжным (IsFastener == "true")
                if (!string.IsNullOrEmpty(child.IsFastener) &&
                    child.IsFastener.Equals("true", StringComparison.OrdinalIgnoreCase))
                {
                    // Устанавливаем Parent и TopParent
                    child.Parent = parentValue;
                    child.TopParent = parentValue;

                    // Проверяем, есть ли уже элемент с таким же Name в коллекции fastenerKit
                    var existingFastener = fastenerKit.Children?.FirstOrDefault(c =>
                        c.Name == child.Name);

                    if (existingFastener != null)
                    {
                        // Если найден — складываем Quantity
                        existingFastener.Quantity += child.Quantity;
                    }
                    else
                    {
                        // Если не найден — добавляем как новый элемент
                        fastenerKit.AddChild(child);
                    }

                    // Удаляем у родителя
                    node.RemoveChild(child);
                }
                else
                {
                    // Рекурсивно обрабатываем дочерние элементы
                    CollectAndRemoveFasteners(child, fastenerKit, parentValue);
                }
            }
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

            // Добавляем изображение в формате Base64
            string base64Image = (obj.PreviewImage != null && obj.PreviewImage.Length > 0)
                ? Convert.ToBase64String(obj.PreviewImage)
                : "";
            WriteElementOrEmpty(writer, "PreviewImage", base64Image);

            WriteElementOrEmpty(writer, "R", obj.R);
            WriteElementOrEmpty(writer, "V", obj.V);
            WriteElementOrEmpty(writer, "Q", obj.Q);
            WriteElementOrEmpty(writer, "Parent", obj.Parent);
            WriteElementOrEmpty(writer, "TopParent", obj.TopParent);
            WriteElementOrEmpty(writer, "FullName", obj.FullName);
            WriteElementOrEmpty(writer, "PathToDXF", obj.PathToDXF);
            WriteElementOrEmpty(writer, "OverallDimensions", obj.OverallDimensions);
            WriteElementOrEmpty(writer, "Coating", obj.Coating);
            WriteElementOrEmpty(writer, "CoverageArea", obj.CoverageArea.ToString());
            WriteElementOrEmpty(writer, "Welding", obj.Welding);
            WriteElementOrEmpty(writer, "LocksmithWork", obj.LocksmithWork);

            // Обработка технологического маршрута
            string technologicalRoute = obj.TechnologicalRoute ?? "";

            // Если Coating пустое, удаляем указанные коды из TechnologicalRoute
            if (string.IsNullOrEmpty(obj.Coating))
            {
                string[] codesToRemove = { "32281,", "32281", "34492,", "34492", "17139,", "17139", "16963,", "16963" };
                foreach (string code in codesToRemove)
                {
                    technologicalRoute = technologicalRoute.Replace(code, "");
                }
                // Очищаем множественные пробелы и запятые
                technologicalRoute = System.Text.RegularExpressions.Regex.Replace(technologicalRoute, @"\s+", " ");
                technologicalRoute = System.Text.RegularExpressions.Regex.Replace(technologicalRoute, @",+", ",");
                technologicalRoute = technologicalRoute.Trim().Trim(',').Trim();
            }

            WriteElementOrEmpty(writer, "TechnologicalRoute", technologicalRoute);
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
            dt.Columns.Add("Превью", typeof(byte[])); // Колонка для превью-изображений
            dt.Columns.Add("Обозначение", typeof(string));
            dt.Columns.Add("Наименование", typeof(string));
            dt.Columns.Add("Кол-во", typeof(int));
            dt.Columns.Add("Раздел спецификации", typeof(string));
            dt.Columns.Add("Материал", typeof(string));
            dt.Columns.Add("Масса", typeof(string));
            dt.Columns.Add("R", typeof(string));
            dt.Columns.Add("V", typeof(string));
            dt.Columns.Add("Q", typeof(string));
            dt.Columns.Add("Узел-1", typeof(string));
            dt.Columns.Add("Узел верхний", typeof(string));
            dt.Columns.Add("Путь до файла", typeof(string));
            dt.Columns.Add("Путь до DXF", typeof(string));
            dt.Columns.Add("Габаритные размеры", typeof(string));
            dt.Columns.Add("Покрытие", typeof(string));
            dt.Columns.Add("Площадь покрытия", typeof(string));
            dt.Columns.Add("Сварочные работы", typeof(string));
            dt.Columns.Add("Слесарные работы", typeof(string));
            dt.Columns.Add("Технологический маршрут", typeof(string));
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

                row["Превью"] = node?.PreviewImage;
                row["Обозначение"] = node.Designation ?? "";
                row["Наименование"] = node.Name ?? "";
                row["Кол-во"] = node.Quantity;
                row["Раздел спецификации"] = node.SpecificationSection ?? "";
                row["Материал"] = node.Material ?? "";
                row["Масса"] = node.Mass.ToString("F2", CultureInfo.InvariantCulture);
                row["R"] = node.R ?? "";
                row["V"] = node.V ?? "";
                row["Q"] = node.Q ?? "";
                row["Узел-1"] = node.Parent ?? "";
                row["Узел верхний"] = node.TopParent ?? "";
                row["Путь до файла"] = node.FullName ?? "";
                row["Путь до DXF"] = node.PathToDXF ?? "";
                row["Габаритные размеры"] = node.OverallDimensions ?? "";
                row["Покрытие"] = node.Coating ?? "";
                row["Площадь покрытия"] = node.CoverageArea;
                row["Сварочные работы"] = node.Welding ?? "";
                row["Слесарные работы"] = node.LocksmithWork ?? "";

                // Обработка технологического маршрута
                string technologicalRoute = node.TechnologicalRoute ?? "";

                // Если Coating пустое, удаляем указанные коды из TechnologicalRoute
                if (string.IsNullOrEmpty(node.Coating))
                {
                    string[] codesToRemove = { "32281,", "32281", "34492,", "34492", "17139,", "17139", "16963,", "16963" };
                    foreach (string code in codesToRemove)
                    {
                        technologicalRoute = technologicalRoute.Replace(code, "");
                    }
                    // Очищаем множественные пробелы и запятые
                    technologicalRoute = System.Text.RegularExpressions.Regex.Replace(technologicalRoute, @"\s+", " ");
                    technologicalRoute = System.Text.RegularExpressions.Regex.Replace(technologicalRoute, @",+", ",");
                    technologicalRoute = technologicalRoute.Trim().Trim(',').Trim();
                }

                row["Технологический маршрут"] = technologicalRoute;
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
            IXLWorksheet worksheet = excelWorkbook.Worksheets.Add("Отчет");

            // Записываем заголовки
            for (int colIndex = 0; colIndex < dt.Columns.Count; colIndex++)
            {
                worksheet.Cell(1, colIndex + 1).Value = dt.Columns[colIndex].ColumnName;
            }

            // Записываем данные из DataTable (пропускаем колонку "Превью" с byte[])
            for (int rowIndex = 0; rowIndex < dt.Rows.Count; rowIndex++)
            {
                for (int colIndex = 0; colIndex < dt.Columns.Count; colIndex++)
                {
                    // Пропускаем колонку "Превью" - изображения вставим отдельно
                    if (dt.Columns[colIndex].ColumnName == "Превью")
                        continue;

                    var cellValue = dt.Rows[rowIndex][colIndex];
                    worksheet.Cell(rowIndex + 2, colIndex + 1).Value = cellValue?.ToString() ?? "";
                }
            }

            // Настройка стилей
            worksheet.RangeUsed().Style.NumberFormat.Format = "@";
            worksheet.Outline.SummaryVLocation = XLOutlineSummaryVLocation.Top;
            // Выравнивание текста по вертикали по центру для всех ячеек
            worksheet.RangeUsed().Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;

            // Создаем таблицу из диапазона данных
            var dataRange = worksheet.Range(1, 1, dt.Rows.Count + 1, dt.Columns.Count);
            IXLTable xLTable = dataRange.CreateTable();
            xLTable.Theme = XLTableTheme.TableStyleLight8;



            // Вставляем изображения из DataTable в Excel
            int previewColumnIndex = 1; // Колонка "Превью" - первая колонка
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                var previewData = dt.Rows[i]["Превью"];
                if (previewData != null && previewData != DBNull.Value)
                {
                    byte[] imageBytes = (byte[])previewData;
                    if (imageBytes != null && imageBytes.Length > 0)
                    {
                        try
                        {
                            int rowIndex = i + 2; // +2 потому что первая строка - заголовки

                            // Создаем временный файл для изображения
                            string tempImagePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString() + ".png");
                            File.WriteAllBytes(tempImagePath, imageBytes);

                            // Добавляем изображение в ячейку
                            var picture = worksheet.AddPicture(tempImagePath);
                            picture.MoveTo(worksheet.Cell(rowIndex, previewColumnIndex));
                            picture.Width = 150;
                            picture.Height = 150;

                            // Программно перемещаем изображение вниз (эмуляция двойного нажатия стрелки вниз)
                            // Смещение вниз на 2 шага (каждый шаг ~3-5 пикселей в Excel)
                            var topLeftCell = picture.TopLeftCell;
                            int offsetPixels = 3; // Подбираем оптимальное смещение
                            picture.MoveTo(topLeftCell, 0, offsetPixels);

                            // Удаляем временный файл
                            try { File.Delete(tempImagePath); } catch { }

                            // Устанавливаем высоту строки для изображения
                            worksheet.Row(rowIndex).Height = 115; // Высота ~150 пикселей в единицах Excel
                        }
                        catch (Exception ex)
                        {
                            // Игнорируем ошибки при вставке изображения
                            System.Diagnostics.Debug.WriteLine($"Ошибка вставки изображения: {ex.Message}");
                        }
                    }
                }
            }

            #region Группирование строк
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
            #endregion

            // Автоматическая подгонка ширины колонок
            worksheet.Columns().AdjustToContents();
            // Устанавливаем ширину колонки "Превью"
            worksheet.Column(previewColumnIndex).Width = 20.7; // Ширина для изображения 150px

            excelWorkbook.SaveAs(pathForExcel + root.Designation + " - " + root.Name + ".xlsx");

            // Теперь у вас есть DataTable dt с данными из treeListView
            // Можно, например, вывести его, открыть диалог с отчетом, привязать к GridView и т.д.            
        }

        private void toolStripOpenExplorer_Click(object sender, EventArgs e)
        {
            string tempPath = System.Reflection.Assembly.GetExecutingAssembly().Location.ToString();
            Process.Start(tempPath.Remove(tempPath.LastIndexOf(@"\")));
        }

        private void toolStripButtonPaint_Click(object sender, EventArgs e)
        {
            if (root == null)
            {
                MessageBox.Show("Сначала загрузите данные из Компас", "Предупреждение", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            Coating coatingForm = new Coating(root);
            coatingForm.Owner = this;
            coatingForm.StartPosition = FormStartPosition.CenterParent;
            coatingForm.ShowDialog(this);

            // Обновляем TreeListView после закрытия формы покрытий
            if (treeListView != null)
            {
                // Полностью перестраиваем TreeListView
                treeListView.RebuildAll(true);
            }
        }
    }
}
