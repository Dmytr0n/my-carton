using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Text;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using FontAwesome.Sharp; // Це дозволить нам брати круті іконки

namespace moy_carton
{
    public partial class Form1 : Form
    {
        private string dataFilePath = "database.csv";
        // Ваші магазини
        private int[] shops = { 30, 31, 32, 33, 11, 20, 40, 13, 10 };

        public Form1()
        {
            InitializeComponent();
            this.FormClosing += new FormClosingEventHandler(Form1_FormClosing);
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            SetupModernLayout();
            SetupGrid();
            FillDates();
            LoadData();
            CalculateSums();
            ApplyEnterpriseStyleGrid(dataGridView1);
        }
        private void ApplyEnterpriseStyleGrid(DataGridView dgv)
        {
            // --- БАЗА ---
            dgv.BackgroundColor = Color.WhiteSmoke;
            dgv.BorderStyle = BorderStyle.None;

            // !!! ЗМІНА ТУТ: Single вмикає сітку і по вертикалі, і по горизонталі
            dgv.CellBorderStyle = DataGridViewCellBorderStyle.Single;

            dgv.ColumnHeadersBorderStyle = DataGridViewHeaderBorderStyle.Single; // Тонкі роздільники в шапці
            dgv.EnableHeadersVisualStyles = false;

            // --- КОЛІР СІТКИ (НАЙВАЖЛИВІШЕ) ---
            // Робимо лінії дуже світлими (LightGray або спеціальний світлий відтінок)
            // Це створює ефект "Excel", а не "старої таблиці"
            dgv.GridColor = Color.FromArgb(224, 224, 224);

            // --- ШАПКА (ТЕМНА) ---
            dgv.ColumnHeadersHeight = 45;
            dgv.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(41, 128, 185);
            dgv.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            dgv.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI Semibold", 10);
            dgv.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;

            // --- РЯДКИ ---
            dgv.DefaultCellStyle.Font = new Font("Segoe UI", 10);
            dgv.RowTemplate.Height = 35;

            // Стиль виділення
            dgv.DefaultCellStyle.SelectionBackColor = Color.FromArgb(255, 240, 150); // Пастельний жовтий
            dgv.DefaultCellStyle.SelectionForeColor = Color.Black;

            // --- ЕФЕКТ "ЗЕБРИ" ---
            dgv.RowsDefaultCellStyle.BackColor = Color.White;
            dgv.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(248, 250, 252); // Дуже легкий відтінок

            dgv.RowHeadersVisible = false;

            // --- ЗАБОРОНА ЗМІНИ РОЗМІРІВ (НОВЕ) ---
            dgv.AllowUserToResizeColumns = false; // Блокуємо ширину стовпців
            dgv.AllowUserToResizeRows = false;    // Блокуємо висоту рядків

            // --- ДОДАТКОВИЙ ТЮНІНГ ДЛЯ КОЛОНКИ "ВСЬОГО" ---
            // Якщо у вас є колонка "total_day", зробимо їй окремий акцент
            if (dgv.Columns.Contains("total_day"))
            {
                dgv.Columns["total_day"].DefaultCellStyle.Font = new Font("Segoe UI", 10, FontStyle.Bold);
                dgv.Columns["total_day"].DefaultCellStyle.BackColor = Color.FromArgb(245, 245, 245);
            }
        }
        private void SetupModernLayout()
        {
            // --- ЗАГАЛЬНИЙ ФОН ---
            this.Text = "Moy_Karton v 1.0.0";
            this.BackColor = Color.FromArgb(240, 242, 245);
            this.Padding = new Padding(0);
            this.Font = new Font("Segoe UI", 9.75F, FontStyle.Regular);

            // --- 1. ШАПКА (HEADER) ---
            Panel headerPanel = new Panel();
            headerPanel.Dock = DockStyle.Top;
            headerPanel.Height = 80; // Трохи вища шапка
            headerPanel.BackColor = Color.FromArgb(30, 39, 46); // Темний матовий колір (Dark Slate)
            headerPanel.Padding = new Padding(15);
            this.Controls.Add(headerPanel);

            // --- ЛОГОТИП (КРУТА ІКОНКА) ---
            // Використовуємо IconPictureBox з бібліотеки FontAwesome
            IconPictureBox logoBox = new IconPictureBox();
            logoBox.IconChar = IconChar.BoxOpen; // Іконка відкритої коробки
            logoBox.IconColor = Color.FromArgb(52, 152, 219); // Блакитний колір іконки
            logoBox.IconSize = 50;
            logoBox.Size = new Size(50, 50);
            logoBox.Location = new Point(20, 15);
            logoBox.BackColor = Color.Transparent;
            logoBox.SizeMode = PictureBoxSizeMode.CenterImage;
            headerPanel.Controls.Add(logoBox);

            // --- ЗАГОЛОВОК ---
            Label titleLabel = new Label();
            titleLabel.Text = "Moy_Karton"; // Англійська назва виглядає дорожче
            titleLabel.ForeColor = Color.White;
            titleLabel.Font = new Font("Segoe UI", 15, FontStyle.Bold);
            titleLabel.AutoSize = true;
            titleLabel.Location = new Point(80, 15);
            headerPanel.Controls.Add(titleLabel);

            // Підзаголовок
            Label subTitle = new Label();
            subTitle.Text = "Система розрахунку картону";
            subTitle.ForeColor = Color.Gray;
            subTitle.Font = new Font("Segoe UI", 9, FontStyle.Regular);
            subTitle.AutoSize = true;
            subTitle.Location = new Point(82, 45);
            headerPanel.Controls.Add(subTitle);

            // --- 2. КНОПКИ З ІКОНКАМИ ---
            // Передаємо конкретні іконки з бібліотеки (FileExcel, Camera)
            StylizeButtonWithIcon(button1, "Експорт Excel", headerPanel, 1, IconChar.FileExcel);
            StylizeButtonWithIcon(button2, "Зберегти Звіт", headerPanel, 2, IconChar.Camera);

            // --- 3. КАРТКА З ТАБЛИЦЕЮ ---
            Panel tableCard = new Panel();
            tableCard.BackColor = Color.White;
            tableCard.Padding = new Padding(15);
            tableCard.Location = new Point(20, 100);
            tableCard.Size = new Size(this.ClientSize.Width - 40, this.ClientSize.Height - 140);
            tableCard.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;

            // Тінь (декоративна панель)
            Panel shadow = new Panel();
            shadow.BackColor = Color.FromArgb(200, 200, 200);
            shadow.Location = new Point(24, 104);
            shadow.Size = tableCard.Size;
            shadow.Anchor = tableCard.Anchor;

            this.Controls.Add(tableCard);
            this.Controls.Add(shadow);
            shadow.SendToBack(); // Ховаємо тінь назад

            // Таблиця
            dataGridView1.Parent = tableCard;
            dataGridView1.Dock = DockStyle.Fill;
            dataGridView1.BringToFront();
        }

        // Допоміжний метод для красивих кнопок
        private void StylizeButtonWithIcon(Button btn, string text, Panel parent, int order, IconChar icon)
        {
            if (btn == null) return;

            btn.Parent = parent;
            btn.Text = "  " + text; // Відступ для тексту
            btn.Size = new Size(160, 40);

            // Генеруємо картинку з іконки
            btn.Image = icon.ToBitmap(Color.White, 24);
            btn.TextImageRelation = TextImageRelation.ImageBeforeText; // Іконка зліва від тексту
            btn.ImageAlign = ContentAlignment.MiddleLeft;
            btn.TextAlign = ContentAlignment.MiddleLeft;
            btn.Padding = new Padding(10, 0, 0, 0); // Відступ іконки від лівого краю

            // Позиція
            int rightMargin = 20;
            int gap = 15;
            btn.Location = new Point(parent.Width - rightMargin - (btn.Width * order) - (gap * (order - 1)), 20);
            btn.Anchor = AnchorStyles.Top | AnchorStyles.Right;

            // Стиль
            btn.FlatStyle = FlatStyle.Flat;
            btn.FlatAppearance.BorderSize = 0;
            btn.Font = new Font("Segoe UI", 9, FontStyle.Bold);
            btn.Cursor = Cursors.Hand;
            btn.ForeColor = Color.White;

            // Кольори кнопок
            if (icon == IconChar.FileExcel)
                btn.BackColor = Color.FromArgb(39, 174, 96); // Зелений
            else if (icon == IconChar.Camera)
                btn.BackColor = Color.FromArgb(41, 128, 185); // Синій
            else
                btn.BackColor = Color.Gray;

            // Анімація при наведенні
            Color baseColor = btn.BackColor;
            btn.MouseEnter += (s, e) => btn.BackColor = ControlPaint.Light(baseColor, 0.1f);
            btn.MouseLeave += (s, e) => btn.BackColor = baseColor;
        }
        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            SaveData();
        }

        private void SetupGrid()
        {
            dataGridView1.Columns.Clear();
            dataGridView1.Rows.Clear();

            // 1. Колонка Дата
            dataGridView1.Columns.Add("date", "Дата");
            dataGridView1.Columns[0].ReadOnly = true;

            // Важливе виправлення з попереднього кроку (ширина дати)
            dataGridView1.Columns[0].AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells;

            // 2. Колонки Магазинів
            foreach (int shop in shops)
            {
                int index = dataGridView1.Columns.Add($"shop_{shop}", shop.ToString());
                dataGridView1.Columns[index].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            }

            // 3. Колонка "Всього за день"
            int totalIndex = dataGridView1.Columns.Add("total_day", "Всього");
            dataGridView1.Columns[totalIndex].ReadOnly = true;
            dataGridView1.Columns[totalIndex].DefaultCellStyle.BackColor = Color.WhiteSmoke;
            dataGridView1.Columns[totalIndex].DefaultCellStyle.Font = new Font(dataGridView1.Font, FontStyle.Bold);
            dataGridView1.Columns[totalIndex].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;

            dataGridView1.AllowUserToAddRows = false;
            dataGridView1.AllowUserToDeleteRows = false;
            dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;

            // --- ГОЛОВНЕ: ПРИБИРАЄМО ТРИКУТНИКИ (СОРТУВАННЯ) ---
            foreach (DataGridViewColumn column in dataGridView1.Columns)
            {
                // Це прибирає іконку і блокує клік по заголовку
                column.SortMode = DataGridViewColumnSortMode.NotSortable;
            }
            // ----------------------------------------------------

            // Події
            dataGridView1.CellValueChanged += (s, e) => CalculateSums();

            dataGridView1.CurrentCellDirtyStateChanged += (s, e) =>
            {
                if (dataGridView1.IsCurrentCellDirty)
                    dataGridView1.CommitEdit(DataGridViewDataErrorContexts.Commit);
            };

            dataGridView1.EditingControlShowing += DataGridView1_EditingControlShowing;
        }

        private void DataGridView1_EditingControlShowing(object sender, DataGridViewEditingControlShowingEventArgs e)
        {
            e.Control.KeyPress -= new KeyPressEventHandler(Column_KeyPress);

            // Перевіряємо, чи це колонка магазинів (індекси від 1 до 7)
            int colIndex = dataGridView1.CurrentCell.ColumnIndex;
            if (colIndex >= 1 && colIndex <= shops.Length)
            {
                TextBox tb = e.Control as TextBox;
                if (tb != null)
                {
                    tb.KeyPress += new KeyPressEventHandler(Column_KeyPress);
                }
            }
        }

        private void Column_KeyPress(object sender, KeyPressEventArgs e)
        {
            // Дозволяємо лише цифри та Backspace
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar))
            {
                e.Handled = true;
            }
        }

        private void FillDates()
        {
            dataGridView1.Rows.Clear();
            DateTime now = DateTime.Now;
            int daysInMonth = DateTime.DaysInMonth(now.Year, now.Month);

            for (int day = 1; day <= daysInMonth; day++)
            {
                DateTime date = new DateTime(now.Year, now.Month, day);
                int rowIndex = dataGridView1.Rows.Add(date.ToString("dd.MM.yyyy"));

                // Підсвічуємо сьогоднішній день
                if (date.Date == DateTime.Now.Date)
                {
                    dataGridView1.Rows[rowIndex].DefaultCellStyle.BackColor = Color.LightYellow;
                }
            }

            // Рядок СУМА
            int sumRowIndex = dataGridView1.Rows.Add();
            dataGridView1.Rows[sumRowIndex].Cells[0].Value = "СУМА:";
            dataGridView1.Rows[sumRowIndex].ReadOnly = true;
            dataGridView1.Rows[sumRowIndex].DefaultCellStyle.BackColor = Color.LightGray;
            dataGridView1.Rows[sumRowIndex].DefaultCellStyle.Font = new Font(dataGridView1.Font, FontStyle.Bold);
        }

        bool isCalculating = false;

        private void CalculateSums()
        {
            if (isCalculating) return;
            isCalculating = true;

            try
            {
                if (dataGridView1.Rows.Count == 0) return;

                int sumRowIndex = dataGridView1.Rows.Count - 1;
                int shopColsCount = shops.Length;
                int totalDayColIndex = shopColsCount + 1; // Індекс колонки "Всього"

                // 1. Очистка рядка сум (вертикальних)
                for (int col = 1; col <= totalDayColIndex; col++)
                    dataGridView1.Rows[sumRowIndex].Cells[col].Value = 0;

                // 2. Проходимо по рядках (дніх)
                for (int row = 0; row < sumRowIndex; row++)
                {
                    int rowSum = 0;

                    // Проходимо по магазинах
                    for (int col = 1; col <= shopColsCount; col++)
                    {
                        var cellValue = dataGridView1.Rows[row].Cells[col].Value;

                        if (cellValue != null && int.TryParse(cellValue.ToString(), out int val))
                        {
                            // Додаємо до суми внизу (вертикально)
                            int prevTotal = Convert.ToInt32(dataGridView1.Rows[sumRowIndex].Cells[col].Value);
                            dataGridView1.Rows[sumRowIndex].Cells[col].Value = prevTotal + val;

                            // Додаємо до суми рядка (горизонтально)
                            rowSum += val;
                        }
                    }
                    // Записуємо суму дня
                    dataGridView1.Rows[row].Cells[totalDayColIndex].Value = rowSum > 0 ? rowSum.ToString() : "";
                }

                // 3. Гранд-тотал (правий нижній кут - сума всього за місяць)
                int grandTotal = 0;
                for (int col = 1; col <= shopColsCount; col++)
                {
                    grandTotal += Convert.ToInt32(dataGridView1.Rows[sumRowIndex].Cells[col].Value);
                }
                dataGridView1.Rows[sumRowIndex].Cells[totalDayColIndex].Value = grandTotal;
            }
            finally
            {
                isCalculating = false;
            }
        }

        private void SaveData()
        {
            try
            {
                using (StreamWriter sw = new StreamWriter(dataFilePath, false, Encoding.UTF8))
                {
                    // Зберігаємо всі рядки, крім останнього (СУМА)
                    for (int i = 0; i < dataGridView1.Rows.Count - 1; i++)
                    {
                        List<string> rowData = new List<string>();
                        // Зберігаємо всі колонки
                        for (int j = 0; j < dataGridView1.Columns.Count; j++)
                        {
                            var val = dataGridView1.Rows[i].Cells[j].Value;
                            rowData.Add(val != null ? val.ToString() : "");
                        }
                        sw.WriteLine(string.Join(";", rowData));
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Помилка збереження: " + ex.Message);
            }
        }

        private void LoadData()
        {
            if (!File.Exists(dataFilePath)) return;

            try
            {
                isCalculating = true;
                string[] lines = File.ReadAllLines(dataFilePath, Encoding.UTF8);

                foreach (string line in lines)
                {
                    string[] parts = line.Split(';');
                    if (parts.Length < 2) continue;

                    string savedDate = parts[0];

                    foreach (DataGridViewRow row in dataGridView1.Rows)
                    {
                        if (row.Index == dataGridView1.Rows.Count - 1) continue;

                        if (row.Cells[0].Value?.ToString() == savedDate)
                        {
                            // Завантажуємо дані (Магазини + Всього)
                            // col < parts.Length гарантує, що ми не вийдемо за межі, якщо файл старий
                            for (int col = 1; col < dataGridView1.Columns.Count && col < parts.Length; col++)
                            {
                                if (int.TryParse(parts[col], out int val))
                                {
                                    row.Cells[col].Value = val;
                                }
                            }
                            break;
                        }
                    }
                }
            }
            finally
            {
                isCalculating = false;
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            ExportToExcel(dataGridView1);
        }

        private void ExportToExcel(DataGridView dgv)
        {
            Excel.Application excelApp = null;
            Excel.Workbook workbook = null;
            Excel.Worksheet worksheet = null;

            try
            {
                excelApp = new Excel.Application();
                if (excelApp == null) { MessageBox.Show("Excel не знайдено!"); return; }

                workbook = excelApp.Workbooks.Add(Type.Missing);
                worksheet = (Excel.Worksheet)workbook.Sheets[1];
                worksheet.Name = "Місячні дані";

                for (int i = 0; i < dgv.Columns.Count; i++)
                    worksheet.Cells[1, i + 1] = dgv.Columns[i].HeaderText;

                for (int i = 0; i < dgv.Rows.Count; i++)
                {
                    for (int j = 0; j < dgv.Columns.Count; j++)
                    {
                        var val = dgv.Rows[i].Cells[j].Value;
                        worksheet.Cells[i + 2, j + 1] = val != null ? val.ToString() : "";
                    }
                }
                worksheet.Columns.AutoFit();
                excelApp.Visible = true;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Помилка: " + ex.Message);
                if (excelApp != null) excelApp.Visible = true;
            }
        }

        // Обробник натискання на кнопку (додайте цю подію до вашої нової кнопки в дизайнері)
        private void button2_Click(object sender, EventArgs e)
        {
            SaveFileDialog sfd = new SaveFileDialog();
            sfd.Filter = "PNG Image|*.png";
            sfd.Title = "Зберегти як зображення";
            sfd.FileName = "Звіт_" + DateTime.Now.ToString("yyyy-MM-dd") + ".png";

            if (sfd.ShowDialog() == DialogResult.OK)
            {
                ExportToPng(dataGridView1, sfd.FileName);
            }
        }
        private void ExportToPng(DataGridView dgv, string filename)
        {
            try
            {
                // 1. Розраховуємо загальні розміри зображення
                int totalWidth = 0;
                foreach (DataGridViewColumn col in dgv.Columns)
                {
                    if (col.Visible) totalWidth += col.Width;
                }

                int totalHeight = dgv.ColumnHeadersHeight;
                foreach (DataGridViewRow row in dgv.Rows)
                {
                    totalHeight += row.Height;
                }

                // Додаємо трохи відступу
                totalWidth += 2;
                totalHeight += 2;

                // 2. Створюємо "полотно" потрібного розміру
                using (Bitmap bmp = new Bitmap(totalWidth, totalHeight))
                {
                    using (Graphics g = Graphics.FromImage(bmp))
                    {
                        // Заливаємо фон білим
                        g.Clear(Color.White);

                        // Налаштування якості
                        g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.AntiAlias;

                        int currentX = 0;
                        int currentY = 0;

                        // --- МАЛЮЄМО ЗАГОЛОВКИ (ШАПКУ) ---
                        foreach (DataGridViewColumn col in dgv.Columns)
                        {
                            if (!col.Visible) continue;

                            Rectangle rect = new Rectangle(currentX, currentY, col.Width, dgv.ColumnHeadersHeight);

                            // Малюємо фон і рамку заголовка
                            g.FillRectangle(Brushes.LightGray, rect);
                            g.DrawRectangle(Pens.Gray, rect);

                            // Пишемо текст заголовка
                            TextRenderer.DrawText(g, col.HeaderText, dgv.ColumnHeadersDefaultCellStyle.Font,
                                rect, Color.Black, TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter);

                            currentX += col.Width;
                        }

                        currentY += dgv.ColumnHeadersHeight;

                        // --- МАЛЮЄМО РЯДКИ З ДАНИМИ ---
                        foreach (DataGridViewRow row in dgv.Rows)
                        {
                            currentX = 0;
                            foreach (DataGridViewCell cell in row.Cells)
                            {
                                if (!cell.OwningColumn.Visible) continue;

                                Rectangle rect = new Rectangle(currentX, currentY, cell.OwningColumn.Width, row.Height);

                                // Отримуємо колір фону (враховуючи ваші жовті дні та сірі суми)
                                Color backColor = cell.InheritedStyle.BackColor;
                                if (backColor.Name == "0" || backColor == Color.Empty) backColor = Color.White;

                                // Малюємо фон клітинки
                                using (SolidBrush brush = new SolidBrush(backColor))
                                {
                                    g.FillRectangle(brush, rect);
                                }

                                // Малюємо рамку
                                g.DrawRectangle(Pens.Gray, rect);

                                // Малюємо значення
                                if (cell.Value != null)
                                {
                                    string text = cell.Value.ToString();
                                    // Використовуємо шрифт клітинки (наприклад, жирний для суми)
                                    Font font = cell.InheritedStyle.Font ?? dgv.DefaultCellStyle.Font;

                                    TextRenderer.DrawText(g, text, font, rect, Color.Black,
                                        TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter);
                                }

                                currentX += cell.OwningColumn.Width;
                            }
                            currentY += row.Height;
                        }
                    }

                    // Зберігаємо файл
                    bmp.Save(filename, System.Drawing.Imaging.ImageFormat.Png);
                }

                MessageBox.Show("Зображення успішно збережено!", "Успіх", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Помилка при збереженні: " + ex.Message);
            }
        }


    }
}