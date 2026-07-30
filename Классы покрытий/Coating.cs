using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Windows.Forms;

namespace ReportKompas
{
    public partial class Coating : Form
    {
        private ObjectAssemblyKompas root;

        public Coating()
        {
            InitializeComponent();
        }

        public Coating(ObjectAssemblyKompas root) : this()
        {
            this.root = root;
            PopulateDataGrid();

            // Размер подгоняем в Load, когда размеры контролов (в т.ч. toolStrip)
            // уже рассчитаны, и сразу центрируем относительно родителя.
            this.Load += (s, e) =>
            {
                AdjustFormSize();
                CenterToParent();
            };

            // Выводим окно на передний план, чтобы оно не пряталось за родительским.
            this.Shown += (s, e) =>
            {
                this.Activate();
                this.BringToFront();
            };

            // Подписываемся на событие закрытия формы
            this.FormClosing += Coating_FormClosing;
        }

        private void Coating_FormClosing(object sender, FormClosingEventArgs e)
        {
            // Сохраняем данные из DataGridView обратно в объекты
            SaveDataFromGrid();

            // Главная форма сама обновит TreeListView после закрытия этой формы
        }

        private void SaveDataFromGrid()
        {
            // Завершаем редактирование текущей ячейки, иначе при нажатии на кнопку
            // ToolStrip введённое значение не коммитится и читается старое.
            dataGridView1.EndEdit();

            int savedCount = 0;
            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                if (row.Tag is ObjectAssemblyKompas node)
                {
                    // Сохраняем покрытие
                    var coatingValue = row.Cells["colCoating"].Value;
                    if (coatingValue != null && !string.IsNullOrWhiteSpace(coatingValue.ToString()))
                    {
                        string newCoating = coatingValue.ToString();
                        if (node.Coating != newCoating)
                        {
                            node.Coating = newCoating;
                            savedCount++;
                            System.Diagnostics.Debug.WriteLine($"Сохранено покрытие для {node.Designation}: {newCoating}");
                        }
                    }

                    // Площадь покрытия уже сохранена в node.CoverageArea при PopulateDataGrid
                    // Если пользователь изменил значение вручную, обновляем его
                    var cellValue = row.Cells["colCoverageArea"].Value;
                    if (cellValue != null)
                    {
                        // Проверяем, является ли значение уже числом double
                        if (cellValue is double area)
                        {
                            node.CoverageArea = area;
                        }
                        // Если это строка, парсим её
                        else if (double.TryParse(cellValue.ToString(), NumberStyles.Any, CultureInfo.InvariantCulture, out double parsedArea))
                        {
                            node.CoverageArea = parsedArea;
                        }
                    }
                }
            }
            System.Diagnostics.Debug.WriteLine($"Всего сохранено покрытий: {savedCount}");
        }

        private void PopulateDataGrid()
        {
            if (root == null)
                return;

            dataGridView1.Rows.Clear();

            // Получаем все узлы с IsPainted > 0
            List<ObjectAssemblyKompas> paintedNodes = GetPaintedNodes(root);

            // Заполняем DataGridView
            foreach (var node in paintedNodes)
            {
                int rowIndex = dataGridView1.Rows.Add();
                DataGridViewRow row = dataGridView1.Rows[rowIndex];

                row.Cells["colDesignation"].Value = node.Designation;
                row.Cells["colName"].Value = node.Name;
                row.Cells["colCoating"].Value = node.Coating;

                // Рассчитываем площадь покрытия как Area * IsPainted / 100
                double calculatedArea = 0;
                string areaNormalized = node.Area?.Replace(',', '.');
                string isPaintedNormalized = node.IsPainted?.Replace(',', '.');

                if (!string.IsNullOrEmpty(areaNormalized) && double.TryParse(areaNormalized, NumberStyles.Any, CultureInfo.InvariantCulture, out double area))
                {
                    if (!string.IsNullOrEmpty(isPaintedNormalized) && double.TryParse(isPaintedNormalized, NumberStyles.Any, CultureInfo.InvariantCulture, out double isPainted))
                    {
                        calculatedArea = (area * isPainted);
                        node.CoverageArea = calculatedArea;
                    }
                }
                row.Cells["colCoverageArea"].Value = calculatedArea;

                // Проверяем, что IsPainted больше 0
                bool isPaintedValue = false;
                if (!string.IsNullOrEmpty(node.IsPainted) && double.TryParse(node.IsPainted, NumberStyles.Any, CultureInfo.InvariantCulture, out double isPaintedCheck))
                {
                    isPaintedValue = isPaintedCheck > 0;
                }
                row.Cells["colIsPainted"].Value = isPaintedValue;

                // Сохраняем ссылку на объект в Tag строки
                row.Tag = node;
            }
        }

        /// <summary>
        /// Устанавливает размер формы строго по размеру таблицы.
        /// Горизонтальная прокрутка отключается, чтобы её полоса не перекрывала строки;
        /// окно подгоняется точно под ширину колонок. Высота ограничивается рабочей
        /// областью экрана; при нехватке места резервируется ширина вертикальной полосы.
        /// </summary>
        private void AdjustFormSize()
        {
            // Горизонтальная прокрутка не нужна — иначе её полоса перекрывает строки
            dataGridView1.ScrollBars = ScrollBars.Vertical;

            // Суммарная ширина видимых колонок
            int gridWidth = 0;
            foreach (DataGridViewColumn col in dataGridView1.Columns)
            {
                if (col.Visible)
                    gridWidth += col.Width;
            }

            // Высота заголовка + всех видимых строк (фактическая суммарная высота)
            int gridHeight = dataGridView1.ColumnHeadersHeight
                + dataGridView1.Rows.GetRowsHeight(DataGridViewElementStates.Visible);

            // Рамки DataGridView (по 1px) + небольшой запас, чтобы колонки точно умещались
            gridWidth += 4;
            gridHeight += 2;

            // Размеры рамок и заголовка окна
            int chromeWidth = this.Width - this.ClientSize.Width;
            int chromeHeight = this.Height - this.ClientSize.Height;

            Rectangle workingArea = Screen.FromControl(this).WorkingArea;
            int maxClientHeight = workingArea.Height - chromeHeight - 40;

            // Если таблица не помещается по высоте — появится вертикальная прокрутка,
            // под неё резервируем ширину (горизонтальной прокрутки не будет)
            int availableForGrid = maxClientHeight - toolStrip1.Height;
            if (gridHeight > availableForGrid)
            {
                gridHeight = availableForGrid;
                gridWidth += SystemInformation.VerticalScrollBarWidth;
            }

            int clientWidth = gridWidth;
            int clientHeight = gridHeight + toolStrip1.Height;

            // Ширина не должна выходить за рабочую область экрана
            int maxClientWidth = workingArea.Width - chromeWidth;
            if (clientWidth > maxClientWidth)
                clientWidth = maxClientWidth;

            this.ClientSize = new Size(clientWidth, clientHeight);
        }

        private List<ObjectAssemblyKompas> GetPaintedNodes(ObjectAssemblyKompas node)
        {
            List<ObjectAssemblyKompas> result = new List<ObjectAssemblyKompas>();

            if (node == null)
                return result;

            // Проверяем текущий узел
            bool shouldAdd = false;
            if (!string.IsNullOrEmpty(node.IsPainted))
            {
                // Попытка парсинга как число
                if (double.TryParse(node.IsPainted, NumberStyles.Any, CultureInfo.InvariantCulture, out double isPainted))
                {
                    shouldAdd = isPainted > 0;
                }
            }

            if (shouldAdd)
            {
                result.Add(node);
            }

            // Рекурсивно обходим детей
            if (node.Children != null)
            {
                foreach (var child in node.Children)
                {
                    result.AddRange(GetPaintedNodes(child));
                }
            }

            return result;
        }

        private void btnAssignCoating_Click(object sender, EventArgs e)
        {
            // Сохраняем данные из DataGridView в объекты
            SaveDataFromGrid();

            // Обновляем отображение в DataGridView
            dataGridView1.Refresh();

            MessageBox.Show("Покрытие успешно назначено для выбранных деталей.",
                "Информация",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);
        }
    }
}
