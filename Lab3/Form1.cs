using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;
using Excel = Microsoft.Office.Interop.Excel;


namespace Lab3
{
    public partial class Form1 : Form
    {
        private bool filesCreated = false;

        public Form1()
        {
            InitializeComponent();
            InitializeEventHandlers();
            SetupDataGridViewColumns();

            // Автоматическое заполнение полей при инициализации формы
            textBox1.Text = "12345"; // Пример значения
            dateTimePicker1.Value = DateTime.Now;
            textBox3.Text = "Поставщик";
            textBox4.Text = "Покупатель";

            // Заполнение данных в DataGridView (пример данных, замените их своими)
            dataGridView1.Rows.Add("1", "Товар1", "10", "100", "");
            dataGridView1.Rows.Add("2", "Товар2", "5", "5550", "");

            // Вычисление и вывод суммы в каждой строке DataGridView
            for (int i = 0; i < dataGridView1.RowCount; i++)
            {
                UpdateTotalForRow(i);
            }

            // Вычисление и вывод итоговой суммы
            UpdateTotals();
        }

        private void UpdateTotals()
        {
            decimal totalSum = 0;

            // Проход по всем строкам в DataGridView и добавление суммы из столбца "Total"
            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                if (row.Cells["Total"].Value != null)
                {
                    totalSum += Convert.ToDecimal(row.Cells["Total"].Value);
                }
            }

            // Вывод общей суммы в ваш элемент управления, например, в какой-то TextBox
            textBoxTotalSum.Text = totalSum.ToString();
        }

        private void InitializeEventHandlers()
        {
            label1.Click += label1_Click;
            label2.Click += label2_Click;
            label4.Click += label4_Click;
            label5.Click += label5_Click;
            textBox1.TextChanged += textBox1_TextChanged;
            textBox4.TextChanged += textBox4_TextChanged;
            dataGridView1.CellContentClick += dataGridView1_CellContentClick;
            Load += Form1_Load;
            btnCreateFiles.Click += btnCreateFiles_Click;
            dateTimePicker1.ValueChanged += dateTimePicker1_ValueChanged;
        }

        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {
            textBox2.Text = dateTimePicker1.Value.ToShortDateString();
        }

        private void dataGridView1_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0 && (e.ColumnIndex == dataGridView1.Columns["Quantity"].Index || e.ColumnIndex == dataGridView1.Columns["Price"].Index))
            {
                UpdateTotalForRow(e.RowIndex);
            }
        }

        private void UpdateTotalForRow(int rowIndex)
        {
            object quantityValue = dataGridView1["Quantity", rowIndex].Value;
            object priceValue = dataGridView1["Price", rowIndex].Value;

            if (quantityValue != null && priceValue != null)
            {
                decimal quantity = Convert.ToDecimal(quantityValue);
                decimal price = Convert.ToDecimal(priceValue);
                decimal total = quantity * price;
                dataGridView1["Total", rowIndex].Value = total;
            }
        }

        private void btnCreateFiles_Click(object sender, EventArgs e)
        {
            Console.WriteLine("btnCreateFiles_Click called");

            CreateAndFillWordFile();
            CreateAndFillExcelFile();

        }

        private void CreateAndFillWordFile()
        {
            try
            {
                Word.Application wordApp = new Word.Application();
                wordApp.Visible = true;
                Word.Document doc = wordApp.Documents.Add();

                Word.Paragraph title = doc.Paragraphs.Add();
                title.Range.Text = "Отчет";
                title.Range.Font.Size = 16;
                title.Range.Font.Bold = 1;
                title.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                title.Range.InsertParagraphAfter();

                doc.Content.Text += $"Расходная накладная номер: {textBox1.Text}\n";
                doc.Content.Text += $"Дата: {dateTimePicker1.Value.ToShortDateString()}\n";
                doc.Content.Text += $"Поставщик: {textBox3.Text}\n";
                doc.Content.Text += $"Покупатель: {textBox4.Text}\n";

                int numRows = dataGridView1.RowCount + 1;
                int numColumns = dataGridView1.ColumnCount;

                Word.Table table = doc.Tables.Add(title.Range, numRows, numColumns);
                table.Borders.Enable = 1;

                // Заполнение заголовков таблицы
                for (int i = 0; i < numColumns; i++)
                {
                    table.Cell(1, i + 1).Range.Text = dataGridView1.Columns[i].HeaderText;
                }

                // Заполнение данных в таблицу
                for (int i = 0; i < dataGridView1.RowCount; i++)
                {
                    for (int j = 0; j < numColumns; j++)
                    {
                        table.Cell(i + 2, j + 1).Range.Text = dataGridView1[j, i].Value?.ToString() ?? "";
                    }
                }

                // Итоговая сумма
                table.Cell(numRows, 1).Range.Text = "Итоговая сумма:";
                table.Cell(numRows, 2).Range.Text = textBoxTotalSum.Text;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        private void CreateAndFillExcelFile()
        {
            try
            {
                Excel.Application excelApp = new Excel.Application();
                excelApp.Visible = true;
                Excel.Workbooks workbooks = excelApp.Workbooks;
                Excel.Workbook workbook = workbooks.Add();
                Excel.Worksheet worksheet = (Excel.Worksheet)workbook.Worksheets[1];

                worksheet.Cells[1, 1] = "Расходная накладная номер";
                worksheet.Cells[1, 2] = "Дата";
                worksheet.Cells[1, 3] = "Поставщик";
                worksheet.Cells[1, 4] = "Покупатель";

                worksheet.Cells[2, 1] = textBox1.Text;
                worksheet.Cells[2, 2] = textBox2.Text;
                worksheet.Cells[2, 3] = textBox3.Text;
                worksheet.Cells[2, 4] = textBox4.Text;

                for (int i = 0; i < dataGridView1.ColumnCount; i++)
                {
                    // Проверка на null перед присваиванием значения ячейки
                    object headerValue = dataGridView1.Columns[i].HeaderText;
                    worksheet.Cells[1, i + 5] = headerValue != null ? headerValue.ToString() : string.Empty;
                }

                for (int i = 0; i < dataGridView1.RowCount; i++)
                {
                    for (int j = 0; j < dataGridView1.ColumnCount; j++)
                    {
                        // Проверка на null перед присваиванием значения ячейки
                        object cellValue = dataGridView1[j, i].Value;
                        worksheet.Cells[i + 2, j + 5] = cellValue != null ? cellValue.ToString() : string.Empty;
                    }
                }

                // Добавление итоговой строки для столбца "Total"
                int rowCount = dataGridView1.RowCount + 2;
                worksheet.Cells[rowCount, 1] = "Итоговая сумма:";
                worksheet.Cells[rowCount, 5].Formula = $"SUM(I2:I{rowCount - 1})"; // Подставьте номер столбца "Total" вместо 5
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        private void SetupDataGridViewColumns()
        {
            dataGridView1.Columns.Add("Number", "номер");
            dataGridView1.Columns.Add("Product", "товар");
            dataGridView1.Columns.Add("Quantity", "количество");
            dataGridView1.Columns.Add("Price", "цена");
            dataGridView1.Columns.Add("Total", "сумма");

            for (int i = 0; i < dataGridView1.RowCount; i++)
            {
                UpdateTotalForRow(i);
            }
        }

        private void label1_Click(object sender, EventArgs e)
        {
            // Обработчик события для label1
        }

        private void label2_Click(object sender, EventArgs e)
        {
            // Обработчик события для label2
        }

        private void label4_Click(object sender, EventArgs e)
        {
            // Обработчик события для label4
        }

        private void label5_Click(object sender, EventArgs e)
        {
            // Обработчик события для label5
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            // Обработчик события для textBox1
        }

        private void textBox4_TextChanged(object sender, EventArgs e)
        {
            // Обработчик события для textBox4
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            // Обработчик события для dataGridView1
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            // Обработчик события для Form1_Load
        }
    }
}
