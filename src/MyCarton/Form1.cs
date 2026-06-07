using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Text;
using System.Windows.Forms;
using ClosedXML.Excel; // ДОДАТИ
using System.Diagnostics; // Знадобиться для відкриття файлу після збереження
using FontAwesome.Sharp; // Це дозволить нам брати круті іконки

namespace moy_carton
{
    public partial class Form1 : Form
    {
        // --- ЗМІННІ ДЛЯ НАВІГАЦІЇ ПО МІСЯЦЯХ ---
        private DateTime currentMonthDate = DateTime.Now; // Зберігає поточний відкритий місяць
                                                          // Динамічний шлях: створює окремий файл для кожного місяця (напр. data_05_2026.csv)
        private string GetCurrentDataFilePath() => $"data_{currentMonthDate:MM_yyyy}.csv";

        private Label lblMonthDisplay; // Текст з назвою місяця
        // Ваші магазини
        private int[] shops = { 30, 31, 32, 33, 11, 20, 40, 13, 10 };

        private string pricesFilePath = "prices.csv"; // Файл для збереження цін
        private Dictionary<int, decimal> shopPrices = new Dictionary<int, decimal>(); // Зберігає ціни
        private Panel pnlSettings; // Панель налаштувань
        private DataGridView dgvPrices; // Таблиця для цін
        private Button btnSettings; // Кнопка налаштувань

        public Form1()
        {
            InitializeComponent();
            this.FormClosing += new FormClosingEventHandler(Form1_FormClosing);
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            LoadPrices(); // <--- ДОДАЙ ЦЕ ПЕРШИМ

            SetupModernLayout();
            SetupSettingsPanel(); // <--- ДОДАЙ ЦЕ ТУТ

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

            if (dgv.Columns.Contains("total_money"))
            {
                dgv.Columns["total_money"].DefaultCellStyle.Font = new Font("Segoe UI", 10, FontStyle.Bold);
                dgv.Columns["total_money"].DefaultCellStyle.BackColor = Color.FromArgb(230, 250, 230); // Світло-зелений
                dgv.Columns["total_money"].DefaultCellStyle.ForeColor = Color.DarkGreen;
            }

            // Стиль для Годин
            if (dgv.Columns.Contains("hours"))
            {
                dgv.Columns["hours"].DefaultCellStyle.Font = new Font("Segoe UI", 10, FontStyle.Bold);
                dgv.Columns["hours"].DefaultCellStyle.BackColor = Color.FromArgb(255, 250, 240); // Легкий теплий відтінок
            }
            // Стиль для Середнього за годину
            if (dgv.Columns.Contains("avg_hour"))
            {
                dgv.Columns["avg_hour"].DefaultCellStyle.Font = new Font("Segoe UI", 10, FontStyle.Bold);
                dgv.Columns["avg_hour"].DefaultCellStyle.BackColor = Color.FromArgb(240, 248, 255); // Світло-синій
                dgv.Columns["avg_hour"].DefaultCellStyle.ForeColor = Color.DarkBlue;
            }
        }
        private void SetupModernLayout()
        {
            // --- ЗАГАЛЬНИЙ ФОН ТА РОЗМІР ВІКНА ---
            this.Text = "Moy_Karton v 1.1.0";
            this.BackColor = Color.FromArgb(240, 242, 245);
            this.Padding = new Padding(0);
            this.Font = new Font("Segoe UI", 9.75F, FontStyle.Regular);

            // ІДЕАЛЬНІ НАЛАШТУВАННЯ ВІКНА:
            this.Size = new Size(1130, 750); // Ті самі ідеальні розміри зі скріншоту
            this.FormBorderStyle = FormBorderStyle.FixedSingle; // Забороняємо тягнути вікно мишкою за краї
            this.MaximizeBox = true; // ДОЗВОЛЯЄМО розгортати вікно на весь екран
            this.StartPosition = FormStartPosition.CenterScreen; // Запуск по центру екрана

            // --- 1. ШАПКА (HEADER) ---
            Panel headerPanel = new Panel();
            headerPanel.Dock = DockStyle.Top;
            headerPanel.Height = 80;
            headerPanel.BackColor = Color.FromArgb(30, 39, 46); // Темний матовий колір (Dark Slate)
            headerPanel.Padding = new Padding(15);
            this.Controls.Add(headerPanel);

            // --- ЛОГОТИП (КРУТА ІКОНКА) ---
            IconPictureBox logoBox = new IconPictureBox();
            logoBox.IconChar = IconChar.BoxOpen;
            logoBox.IconColor = Color.FromArgb(52, 152, 219);
            logoBox.IconSize = 50;
            logoBox.Size = new Size(50, 50);
            logoBox.Location = new Point(20, 15);
            logoBox.BackColor = Color.Transparent;
            logoBox.SizeMode = PictureBoxSizeMode.CenterImage;
            headerPanel.Controls.Add(logoBox);

            // --- ЗАГОЛОВОК ---
            Label titleLabel = new Label();
            titleLabel.Text = "Moy_Karton";
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

            // --- 2. КНОПКИ З ІКОНКАМИ (ПРАВА ЧАСТИНА) ---
            StylizeButtonWithIcon(button1, "Експорт Excel", headerPanel, 1, IconChar.FileExcel);
            StylizeButtonWithIcon(button2, "Зберегти Звіт", headerPanel, 2, IconChar.Camera);

            btnSettings = new Button();
            StylizeButtonWithIcon(btnSettings, "Ціни", headerPanel, 3, IconChar.Cogs);
            btnSettings.Click += (s, ev) => { pnlSettings.Visible = true; pnlSettings.BringToFront(); };

            // --- ПАНЕЛЬ НАВІГАЦІЇ МІСЯЦІВ (ІДЕАЛЬНА СТИЛІЗАЦІЯ) ---
            Panel monthNavPanel = new Panel();
            monthNavPanel.Size = new Size(250, 40);
            // Розміщуємо по X=320, щоб бути точно між заголовком і кнопками
            monthNavPanel.Location = new Point(320, 20);
            monthNavPanel.BackColor = Color.FromArgb(40, 50, 60); // Темний акцентний фон

            // Робимо закруглені кути (переконайся, що метод CreateRoundRectRgn є в твоєму класі)
            monthNavPanel.Region = Region.FromHrgn(CreateRoundRectRgn(0, 0, monthNavPanel.Width, monthNavPanel.Height, 8, 8));
            headerPanel.Controls.Add(monthNavPanel);

            // Кнопка ВЛІВО (<) з ефектом наведення
            Button btnPrevMonth = new Button { Size = new Size(45, 40), Dock = DockStyle.Left, FlatStyle = FlatStyle.Flat, Cursor = Cursors.Hand, BackColor = Color.Transparent };
            btnPrevMonth.FlatAppearance.BorderSize = 0;
            btnPrevMonth.FlatAppearance.MouseOverBackColor = Color.FromArgb(60, 75, 90); // Світлішає при наведенні
            btnPrevMonth.Image = IconChar.ChevronLeft.ToBitmap(Color.White, 20);
            btnPrevMonth.Click += (s, ev) => ChangeMonth(-1);

            // Кнопка ВПРАВО (>) з ефектом наведення
            Button btnNextMonth = new Button { Size = new Size(45, 40), Dock = DockStyle.Right, FlatStyle = FlatStyle.Flat, Cursor = Cursors.Hand, BackColor = Color.Transparent };
            btnNextMonth.FlatAppearance.BorderSize = 0;
            btnNextMonth.FlatAppearance.MouseOverBackColor = Color.FromArgb(60, 75, 90); // Світлішає при наведенні
            btnNextMonth.Image = IconChar.ChevronRight.ToBitmap(Color.White, 20);
            btnNextMonth.Click += (s, ev) => ChangeMonth(1);

            // Назва місяця
            lblMonthDisplay = new Label
            {
                Dock = DockStyle.Fill,
                ForeColor = Color.White,
                Font = new Font("Segoe UI", 11, FontStyle.Bold), // Шрифт трохи менший, щоб ідеально влізло
                TextAlign = ContentAlignment.MiddleCenter
            };

            // ВАЖЛИВО: Порядок додавання! Спочатку краї (кнопки), потім центр (текст)
            monthNavPanel.Controls.Add(btnPrevMonth);
            monthNavPanel.Controls.Add(btnNextMonth);
            monthNavPanel.Controls.Add(lblMonthDisplay);

            UpdateMonthLabel();

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

            dataGridView1.Columns.Add("date", "Дата");
            dataGridView1.Columns[0].ReadOnly = true;
            dataGridView1.Columns[0].AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells;

            // Колонки Магазинів
            foreach (int shop in shops)
            {
                int index = dataGridView1.Columns.Add($"shop_{shop}", shop.ToString());
                dataGridView1.Columns[index].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            }

            // НОВА КОЛОНКА: Години роботи
            int hoursIndex = dataGridView1.Columns.Add("hours", "Години");
            dataGridView1.Columns[hoursIndex].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dataGridView1.Columns[hoursIndex].AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells; // ФІКСАЦІЯ ШИРИНИ

            // Колонка "Всього коробок"
            int totalIndex = dataGridView1.Columns.Add("total_day", "Всього");
            dataGridView1.Columns[totalIndex].ReadOnly = true;
            dataGridView1.Columns[totalIndex].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dataGridView1.Columns[totalIndex].AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells; // ФІКСАЦІЯ ШИРИНИ

            // Колонка "Сума (Гроші)"
            int moneyIndex = dataGridView1.Columns.Add("total_money", "Сума PLN");
            dataGridView1.Columns[moneyIndex].ReadOnly = true;
            dataGridView1.Columns[moneyIndex].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dataGridView1.Columns[moneyIndex].AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells; // ФІКСАЦІЯ ШИРИНИ

            // НОВА КОЛОНКА: Середнє за годину
            int avgIndex = dataGridView1.Columns.Add("avg_hour", "Сер./год");
            dataGridView1.Columns[avgIndex].ReadOnly = true;
            dataGridView1.Columns[avgIndex].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dataGridView1.Columns[avgIndex].AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells; // ФІКСАЦІЯ ШИРИНИ

            dataGridView1.AllowUserToAddRows = false;
            dataGridView1.AllowUserToDeleteRows = false;
            // Магазини розтягуються рівномірно на весь вільний простір
            dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;

            foreach (DataGridViewColumn column in dataGridView1.Columns)
                column.SortMode = DataGridViewColumnSortMode.NotSortable;

            dataGridView1.CellValueChanged += (s, e) => CalculateSums();
            dataGridView1.CurrentCellDirtyStateChanged += (s, e) =>
            {
                if (dataGridView1.IsCurrentCellDirty) dataGridView1.CommitEdit(DataGridViewDataErrorContexts.Commit);
            };

            dataGridView1.EditingControlShowing += DataGridView1_EditingControlShowing;
            dataGridView1.CellEnter += DataGridView1_CellEnter;
            dataGridView1.CellLeave += DataGridView1_CellLeave;
        }

        private void DataGridView1_EditingControlShowing(object sender, DataGridViewEditingControlShowingEventArgs e)
        {
            e.Control.KeyPress -= new KeyPressEventHandler(Column_KeyPress);
            int colIndex = dataGridView1.CurrentCell.ColumnIndex;

            // Дозволяємо ввід для магазинів ТА колонки годин (індекс shops.Length + 1)
            if (colIndex >= 1 && colIndex <= shops.Length + 1)
            {
                if (e.Control is TextBox tb) tb.KeyPress += new KeyPressEventHandler(Column_KeyPress);
            }
        }

        private void Column_KeyPress(object sender, KeyPressEventArgs e)
        {
            // Дозволяємо цифри, Backspace, та кому/крапку для годин
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) && e.KeyChar != ',' && e.KeyChar != '.')
                e.Handled = true;

            // Автоматично міняємо крапку на кому для правильності формату в C#
            if (e.KeyChar == '.') e.KeyChar = ',';

            // Блокуємо введення другої коми підряд
            if (sender is TextBox tb && e.KeyChar == ',' && tb.Text.Contains(","))
                e.Handled = true;
        }

        private void FillDates()
        {
            dataGridView1.Rows.Clear();

            // Використовуємо наш вибраний місяць замість DateTime.Now
            int daysInMonth = DateTime.DaysInMonth(currentMonthDate.Year, currentMonthDate.Month);

            for (int day = 1; day <= daysInMonth; day++)
            {
                DateTime date = new DateTime(currentMonthDate.Year, currentMonthDate.Month, day);
                int rowIndex = dataGridView1.Rows.Add(date.ToString("dd.MM.yyyy"));

                // Підсвічуємо сьогоднішній день ТІЛЬКИ якщо ми дивимося поточний місяць і рік
                if (date.Date == DateTime.Now.Date)
                {
                    dataGridView1.Rows[rowIndex].DefaultCellStyle.BackColor = Color.LightYellow;
                }
            }

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
                int hoursColIndex = shopColsCount + 1;
                int totalDayColIndex = shopColsCount + 2;
                int totalMoneyColIndex = shopColsCount + 3;
                int avgHourColIndex = shopColsCount + 4;

                // Очистка нижнього рядка СУМА
                for (int col = 1; col <= avgHourColIndex; col++)
                    dataGridView1.Rows[sumRowIndex].Cells[col].Value = 0;

                for (int row = 0; row < sumRowIndex; row++)
                {
                    int rowSumBoxes = 0;
                    decimal rowMoney = 0m;
                    decimal rowHours = 0m;

                    // 1. Рахуємо коробки та гроші
                    for (int col = 1; col <= shopColsCount; col++)
                    {
                        var cellValue = dataGridView1.Rows[row].Cells[col].Value;
                        if (cellValue != null && int.TryParse(cellValue.ToString(), out int val))
                        {
                            int prevTotal = Convert.ToInt32(dataGridView1.Rows[sumRowIndex].Cells[col].Value);
                            dataGridView1.Rows[sumRowIndex].Cells[col].Value = prevTotal + val;

                            rowSumBoxes += val;
                            int shopId = shops[col - 1];
                            rowMoney += val * shopPrices[shopId];
                        }
                    }

                    // 2. Читаємо введені години
                    var hoursCell = dataGridView1.Rows[row].Cells[hoursColIndex].Value;
                    if (hoursCell != null && decimal.TryParse(hoursCell.ToString().Replace(".", ","), out decimal h))
                    {
                        rowHours = h;
                        decimal prevTotalHours = Convert.ToDecimal(dataGridView1.Rows[sumRowIndex].Cells[hoursColIndex].Value);
                        dataGridView1.Rows[sumRowIndex].Cells[hoursColIndex].Value = prevTotalHours + h;
                    }

                    // Записуємо підсумки дня
                    dataGridView1.Rows[row].Cells[totalDayColIndex].Value = rowSumBoxes > 0 ? rowSumBoxes.ToString() : "";
                    dataGridView1.Rows[row].Cells[totalMoneyColIndex].Value = rowMoney > 0 ? rowMoney.ToString("0.00") : "";

                    // 3. Рахуємо середнє за годину
                    if (rowHours > 0)
                        dataGridView1.Rows[row].Cells[avgHourColIndex].Value = (rowMoney / rowHours).ToString("0.00");
                    else
                        dataGridView1.Rows[row].Cells[avgHourColIndex].Value = "";
                }

                // Гранд-тотал (Всього за місяць)
                int grandTotalBoxes = 0;
                decimal grandTotalMoney = 0m;

                for (int col = 1; col <= shopColsCount; col++)
                {
                    int colTotal = Convert.ToInt32(dataGridView1.Rows[sumRowIndex].Cells[col].Value);
                    grandTotalBoxes += colTotal;
                    int shopId = shops[col - 1];
                    grandTotalMoney += colTotal * shopPrices[shopId];
                }

                dataGridView1.Rows[sumRowIndex].Cells[totalDayColIndex].Value = grandTotalBoxes;
                dataGridView1.Rows[sumRowIndex].Cells[totalMoneyColIndex].Value = grandTotalMoney.ToString("0.00");

                decimal grandTotalHours = Convert.ToDecimal(dataGridView1.Rows[sumRowIndex].Cells[hoursColIndex].Value);
                if (grandTotalHours > 0)
                    dataGridView1.Rows[sumRowIndex].Cells[avgHourColIndex].Value = (grandTotalMoney / grandTotalHours).ToString("0.00");
            }
            finally { isCalculating = false; }
        }

        private void SaveData()
        {
            try
            {
                using (StreamWriter sw = new StreamWriter(GetCurrentDataFilePath(), false, Encoding.UTF8))
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
            string currentFile = GetCurrentDataFilePath();
            if (!File.Exists(currentFile)) return;

            try
            {
                isCalculating = true;
                string[] lines = File.ReadAllLines(currentFile, Encoding.UTF8);

                foreach (string line in lines)
                {
                    string[] parts = line.Split(';');
                    if (parts.Length < 2) continue;

                    string savedDate = parts[0];
                    foreach (DataGridViewRow row in dataGridView1.Rows)
                    {
                        // Пропускаємо останній рядок СУМА
                        if (row.Index == dataGridView1.Rows.Count - 1) continue;

                        if (row.Cells[0].Value?.ToString() == savedDate)
                        {
                            // 1. Завантажуємо дані магазинів
                            for (int col = 1; col <= shops.Length && col < parts.Length; col++)
                            {
                                if (int.TryParse(parts[col], out int val))
                                    row.Cells[col].Value = val;
                            }

                            // 2. Завантажуємо години (індекс стовпця йде відразу після магазинів)
                            if (parts.Length > shops.Length + 1)
                            {
                                string hoursValue = parts[shops.Length + 1];
                                if (decimal.TryParse(hoursValue.Replace(".", ","), out decimal h))
                                    row.Cells[shops.Length + 1].Value = h;
                            }
                            break;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Помилка завантаження даних місяця: " + ex.Message);
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
            try
            {
                using (var workbook = new XLWorkbook())
                {
                    var worksheet = workbook.Worksheets.Add("Місячні дані");

                    // 1. Заголовки стовпців
                    for (int i = 0; i < dgv.Columns.Count; i++)
                    {
                        worksheet.Cell(1, i + 1).Value = dgv.Columns[i].HeaderText;
                        worksheet.Cell(1, i + 1).Style.Font.Bold = true;
                        worksheet.Cell(1, i + 1).Style.Fill.BackgroundColor = XLColor.FromHtml("#2980b9");
                        worksheet.Cell(1, i + 1).Style.Font.FontColor = XLColor.White;
                    }

                    // 2. Дані таблиці
                    for (int i = 0; i < dgv.Rows.Count; i++)
                    {
                        for (int j = 0; j < dgv.Columns.Count; j++)
                        {
                            var val = dgv.Rows[i].Cells[j].Value;
                            if (val != null)
                            {
                                // Перевіряємо, чи це число, щоб у Excel можна було рахувати формули
                                if (double.TryParse(val.ToString(), out double numericVal))
                                    worksheet.Cell(i + 2, j + 1).Value = numericVal;
                                else
                                    worksheet.Cell(i + 2, j + 1).Value = val.ToString();
                            }
                        }
                    }

                    // 3. Автопідбір ширини колонок
                    worksheet.Columns().AdjustToContents();

                    // 4. Збереження через діалогове вікно
                    SaveFileDialog sfd = new SaveFileDialog();
                    sfd.Filter = "Excel Workbook|*.xlsx";
                    sfd.FileName = "Звіт_картон_" + DateTime.Now.ToString("MM_yyyy");

                    if (sfd.ShowDialog() == DialogResult.OK)
                    {
                        workbook.SaveAs(sfd.FileName);
                        MessageBox.Show("Дані успішно експортовано в Excel!", "Успіх", MessageBoxButtons.OK, MessageBoxIcon.Information);

                        // Опціонально: відкрити файл після збереження
                        Process.Start(new ProcessStartInfo(sfd.FileName) { UseShellExecute = true });
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Помилка при експорті в Excel: " + ex.Message, "Помилка", MessageBoxButtons.OK, MessageBoxIcon.Error);
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
        // --- МЕТОДИ ДЛЯ ПІДСВІЧУВАННЯ РЯДКА ---
        private void DataGridView1_CellEnter(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0 && e.RowIndex < dataGridView1.Rows.Count - 1)
                dataGridView1.Rows[e.RowIndex].DefaultCellStyle.BackColor = Color.FromArgb(226, 240, 255); // Ніжно-блакитний
        }

        private void DataGridView1_CellLeave(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0 && e.RowIndex < dataGridView1.Rows.Count - 1)
            {
                if (e.RowIndex % 2 == 0) dataGridView1.Rows[e.RowIndex].DefaultCellStyle.BackColor = Color.White;
                else dataGridView1.Rows[e.RowIndex].DefaultCellStyle.BackColor = Color.FromArgb(248, 250, 252);

                if (dataGridView1.Rows[e.RowIndex].Cells[0].Value?.ToString() == DateTime.Now.ToString("dd.MM.yyyy"))
                    dataGridView1.Rows[e.RowIndex].DefaultCellStyle.BackColor = Color.LightYellow;
            }
        }

        // --- МЕТОДИ ДЛЯ ЦІН ---
        private void LoadPrices()
        {
            foreach (int shop in shops) shopPrices[shop] = 0m;
            if (!File.Exists(pricesFilePath)) return;
            try
            {
                string[] lines = File.ReadAllLines(pricesFilePath, Encoding.UTF8);
                foreach (string line in lines)
                {
                    string[] parts = line.Split(';');
                    if (parts.Length == 2 && int.TryParse(parts[0], out int shopId) && decimal.TryParse(parts[1], out decimal price))
                        shopPrices[shopId] = price;
                }
            }
            catch { }
        }

        private void SavePrices()
        {
            try
            {
                using (StreamWriter sw = new StreamWriter(pricesFilePath, false, Encoding.UTF8))
                {
                    foreach (DataGridViewRow row in dgvPrices.Rows)
                    {
                        if (row.Cells[0].Value != null && row.Cells[1].Value != null)
                        {
                            string shopStr = row.Cells[0].Value.ToString().Replace("Mag ", "");
                            if (int.TryParse(shopStr, out int shopId) && decimal.TryParse(row.Cells[1].Value.ToString().Replace(".", ","), out decimal price))
                            {
                                shopPrices[shopId] = price;
                                sw.WriteLine($"{shopId};{price}");
                            }
                        }
                    }
                }
                CalculateSums();
                MessageBox.Show("Ціни успішно збережено!", "Успіх", MessageBoxButtons.OK, MessageBoxIcon.Information);
                pnlSettings.Visible = false;
            }
            catch (Exception ex) { MessageBox.Show("Помилка: " + ex.Message); }
        }

        private void SetupSettingsPanel()
        {
            // 1. Головна панель (ЗРОБИЛИ КОМПАКТНІШОЮ: 360x430)
            pnlSettings = new Panel
            {
                Size = new Size(360, 430),
                BackColor = Color.White,
                BorderStyle = BorderStyle.FixedSingle,
                Visible = false
            };

            // Центруємо на екрані
            pnlSettings.Location = new Point((this.ClientSize.Width - pnlSettings.Width) / 2, (this.ClientSize.Height - pnlSettings.Height) / 2);
            pnlSettings.Anchor = AnchorStyles.None;

            // 2. Красива шапка панелі (Header)
            Panel headerSettings = new Panel
            {
                Dock = DockStyle.Top,
                Height = 50,
                BackColor = Color.FromArgb(41, 128, 185)
            };
            pnlSettings.Controls.Add(headerSettings);

            // Заголовок у шапці
            Label lblTitle = new Label
            {
                Text = "Налаштування цін (Stawka)",
                ForeColor = Color.White,
                Font = new Font("Segoe UI", 12, FontStyle.Bold),
                AutoSize = true,
                Location = new Point(15, 12)
            };
            headerSettings.Controls.Add(lblTitle);

            // БЛОК З ХРЕСТИКОМ ВИДАЛЕНО

            // 3. Таблиця цін (Зменшено висоту, бо є прокрутка)
            dgvPrices = new DataGridView
            {
                Location = new Point(20, 65),
                Size = new Size(320, 270),    // Висота тепер 270 замість 350
                AllowUserToAddRows = false,
                RowHeadersVisible = false,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                BackgroundColor = Color.White,
                BorderStyle = BorderStyle.None,
                CellBorderStyle = DataGridViewCellBorderStyle.SingleHorizontal,
                EnableHeadersVisualStyles = false
            };

            // Стиль шапки таблиці
            dgvPrices.ColumnHeadersHeight = 40;
            dgvPrices.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(240, 242, 245);
            dgvPrices.ColumnHeadersDefaultCellStyle.ForeColor = Color.Black;
            dgvPrices.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI Semibold", 10);
            dgvPrices.ColumnHeadersBorderStyle = DataGridViewHeaderBorderStyle.None;

            // Стиль рядків таблиці
            dgvPrices.DefaultCellStyle.Font = new Font("Segoe UI", 10);
            dgvPrices.RowTemplate.Height = 35;
            dgvPrices.DefaultCellStyle.SelectionBackColor = Color.FromArgb(226, 240, 255);
            dgvPrices.DefaultCellStyle.SelectionForeColor = Color.Black;

            // Колонки
            dgvPrices.Columns.Add("shop", "Magazyn");
            dgvPrices.Columns[0].ReadOnly = true;
            dgvPrices.Columns.Add("price", "Ціна (PLN)");

            // Заповнення даними
            foreach (int shop in shops)
            {
                dgvPrices.Rows.Add($"Mag {shop}", shopPrices[shop].ToString("0.0000"));
            }

            pnlSettings.Controls.Add(dgvPrices);

            // 4. Нижні кнопки (Підтягнуті вище)
            Button btnSave = new Button
            {
                Text = "Зберегти",
                Location = new Point(20, 355), // Підняли Y-координату
                Size = new Size(150, 45),
                BackColor = Color.FromArgb(39, 174, 96),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 10, FontStyle.Bold),
                Cursor = Cursors.Hand
            };
            btnSave.FlatAppearance.BorderSize = 0;
            btnSave.Click += (s, e) => SavePrices();
            pnlSettings.Controls.Add(btnSave);

            Button btnClose = new Button
            {
                Text = "Скасувати",
                Location = new Point(190, 355), // Підняли Y-координату
                Size = new Size(150, 45),
                BackColor = Color.FromArgb(149, 165, 166),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 10, FontStyle.Bold),
                Cursor = Cursors.Hand
            };
            btnClose.FlatAppearance.BorderSize = 0;
            btnClose.Click += (s, e) => pnlSettings.Visible = false;
            pnlSettings.Controls.Add(btnClose);

            // Додаємо панель на форму
            this.Controls.Add(pnlSettings);
            pnlSettings.BringToFront();
        }
        private void UpdateMonthLabel()
        {
            // Робимо красиву назву українською мовою з великої літери (напр. "Травень 2026")
            string monthName = currentMonthDate.ToString("MMMM yyyy", new System.Globalization.CultureInfo("uk-UA"));
            lblMonthDisplay.Text = char.ToUpper(monthName[0]) + monthName.Substring(1);
        }

        private void ChangeMonth(int offset)
        {
            // Прораховуємо, на який місяць користувач хоче перейти
            DateTime targetDate = currentMonthDate.AddMonths(offset);

            // ОБМЕЖЕННЯ: Перевіряємо, чи збігається рік з поточним календарним роком
            if (targetDate.Year != DateTime.Now.Year)
            {
                // Не пускаємо далі і, за бажанням, можемо вивести попередження
                // MessageBox.Show("Перегляд даних доступний лише в межах поточного року.", "Обмеження", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // 1. Зберігаємо дані поточного відкритого місяця!
            SaveData();

            // 2. Змінюємо дату на дозволену
            currentMonthDate = targetDate;
            UpdateMonthLabel();

            // 3. Оновлюємо таблицю (генеруємо нові дати і вантажимо інший файл)
            FillDates();
            LoadData();
            CalculateSums();
        }
        [System.Runtime.InteropServices.DllImport("Gdi32.dll", EntryPoint = "CreateRoundRectRgn")]
        private static extern IntPtr CreateRoundRectRgn
(
    int nLeftRect, int nTopRect, int nRightRect, int nBottomRect, int nWidthEllipse, int nHeightEllipse
);
    }

}