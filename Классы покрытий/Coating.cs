using System;
using System.Collections.Generic;
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
                if (!string.IsNullOrEmpty(node.Area) && double.TryParse(node.Area, NumberStyles.Any, CultureInfo.InvariantCulture, out double area))
                {
                    if (!string.IsNullOrEmpty(node.IsPainted) && double.TryParse(node.IsPainted, NumberStyles.Any, CultureInfo.InvariantCulture, out double isPainted))
                    {
                        calculatedArea = (area * isPainted) / 100;
                        // Сохраняем значение в объект
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
