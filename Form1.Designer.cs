namespace Optimizations;

partial class Form1
{
    /// <summary>
    ///  Required designer variable.
    /// </summary>
    private System.ComponentModel.IContainer components = null;

    /// <summary>
    ///  Clean up any resources being used.
    /// </summary>
    /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
    protected override void Dispose(bool disposing)
    {
        if (disposing && (components != null))
        {
            components.Dispose();
        }
        base.Dispose(disposing);
    }

    #region Windows Form Designer generated code

    /// <summary>
    ///  Required method for Designer support - do not modify
    ///  the contents of this method with the code editor.
    /// </summary>
    private void InitializeComponent()
    {
        button1 = new Button();
        button2 = new Button();
        button3 = new Button();
        dataGridView1 = new DataGridView();
        columnName = new DataGridViewTextBoxColumn();
        columnWorkload = new DataGridViewTextBoxColumn();
        columnSignificance = new DataGridViewTextBoxColumn();
        columnSemester = new DataGridViewTextBoxColumn();
        titleLabel = new Label();
        statusPanel = new Panel();
        progressBar = new ProgressBar();
        statusLabel = new Label();
        mainPanel = new Panel();
        buttonPanel = new Panel();
        ((System.ComponentModel.ISupportInitialize)dataGridView1).BeginInit();
        statusPanel.SuspendLayout();
        mainPanel.SuspendLayout();
        buttonPanel.SuspendLayout();
        SuspendLayout();
        // 
        // titleLabel
        // 
        titleLabel.BackColor = Color.FromArgb(45, 66, 91);
        titleLabel.Dock = DockStyle.Top;
        titleLabel.Font = new Font("Segoe UI", 18F, FontStyle.Bold);
        titleLabel.ForeColor = Color.White;
        titleLabel.Location = new Point(0, 0);
        titleLabel.Name = "titleLabel";
        titleLabel.Size = new Size(1024, 60);
        titleLabel.TabIndex = 0;
        titleLabel.Text = "Система оптимизации учебного плана";
        titleLabel.TextAlign = ContentAlignment.MiddleCenter;
        // 
        // buttonPanel
        // 
        buttonPanel.BackColor = Color.FromArgb(248, 249, 250);
        buttonPanel.Controls.Add(button1);
        buttonPanel.Controls.Add(button3);
        buttonPanel.Controls.Add(button2);
        buttonPanel.Dock = DockStyle.Top;
        buttonPanel.Location = new Point(0, 60);
        buttonPanel.Name = "buttonPanel";
        buttonPanel.Padding = new Padding(20, 15, 20, 15);
        buttonPanel.Size = new Size(1024, 70);
        buttonPanel.TabIndex = 1;
        // 
        // button1
        // 
        button1.BackColor = Color.FromArgb(40, 167, 69);
        button1.FlatAppearance.BorderSize = 0;
        button1.FlatAppearance.MouseDownBackColor = Color.FromArgb(33, 136, 56);
        button1.FlatAppearance.MouseOverBackColor = Color.FromArgb(52, 183, 85);
        button1.FlatStyle = FlatStyle.Flat;
        button1.Font = new Font("Segoe UI", 11F, FontStyle.Bold);
        button1.ForeColor = Color.White;
        button1.Location = new Point(20, 15);
        button1.Name = "button1";
        button1.Size = new Size(160, 40);
        button1.TabIndex = 0;
        button1.Text = "📁 Загрузить данные";
        button1.UseVisualStyleBackColor = false;
        button1.Click += button1_Click;
        // 
        // button3
        // 
        button3.BackColor = Color.FromArgb(0, 123, 255);
        button3.FlatAppearance.BorderSize = 0;
        button3.FlatAppearance.MouseDownBackColor = Color.FromArgb(0, 86, 179);
        button3.FlatAppearance.MouseOverBackColor = Color.FromArgb(38, 143, 255);
        button3.FlatStyle = FlatStyle.Flat;
        button3.Font = new Font("Segoe UI", 11F, FontStyle.Bold);
        button3.ForeColor = Color.White;
        button3.Location = new Point(200, 15);
        button3.Name = "button3";
        button3.Size = new Size(180, 40);
        button3.TabIndex = 2;
        button3.Text = "💾 Сохранить результаты";
        button3.UseVisualStyleBackColor = false;
        button3.Click += button3_Click_1;
        // 
        // button2
        // 
        button2.Anchor = AnchorStyles.Top | AnchorStyles.Right;
        button2.BackColor = Color.FromArgb(220, 53, 69);
        button2.FlatAppearance.BorderSize = 0;
        button2.FlatAppearance.MouseDownBackColor = Color.FromArgb(200, 35, 51);
        button2.FlatAppearance.MouseOverBackColor = Color.FromArgb(229, 83, 98);
        button2.FlatStyle = FlatStyle.Flat;
        button2.Font = new Font("Segoe UI", 11F, FontStyle.Bold);
        button2.ForeColor = Color.White;
        button2.Location = new Point(844, 15);
        button2.Name = "button2";
        button2.Size = new Size(160, 40);
        button2.TabIndex = 1;
        button2.Text = "❌ Закрыть";
        button2.UseVisualStyleBackColor = false;        button2.Click += button2_Click;
        // 
        // mainPanel
        // 
        mainPanel.BackColor = Color.White;
        mainPanel.Controls.Add(dataGridView1);
        mainPanel.Dock = DockStyle.Fill;
        mainPanel.Location = new Point(0, 130);
        mainPanel.Name = "mainPanel";
        mainPanel.Padding = new Padding(20);
        mainPanel.Size = new Size(1024, 490);
        mainPanel.TabIndex = 2;
        // 
        // dataGridView1
        // 
        dataGridView1.AllowUserToAddRows = false;
        dataGridView1.AllowUserToDeleteRows = false;
        dataGridView1.AllowUserToResizeRows = false;
        dataGridView1.BackgroundColor = Color.White;
        dataGridView1.BorderStyle = BorderStyle.None;
        dataGridView1.CellBorderStyle = DataGridViewCellBorderStyle.SingleHorizontal;
        dataGridView1.ColumnHeadersBorderStyle = DataGridViewHeaderBorderStyle.None;
        dataGridView1.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
        dataGridView1.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(52, 58, 64);
        dataGridView1.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 11F, FontStyle.Bold);
        dataGridView1.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
        dataGridView1.ColumnHeadersDefaultCellStyle.Padding = new Padding(10, 0, 10, 0);
        dataGridView1.ColumnHeadersDefaultCellStyle.SelectionBackColor = Color.FromArgb(52, 58, 64);
        dataGridView1.ColumnHeadersDefaultCellStyle.SelectionForeColor = Color.White;
        dataGridView1.ColumnHeadersDefaultCellStyle.WrapMode = DataGridViewTriState.False;
        dataGridView1.ColumnHeadersHeight = 45;
        dataGridView1.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing;
        dataGridView1.Columns.AddRange(new DataGridViewColumn[] { columnName, columnWorkload, columnSignificance, columnSemester });
        dataGridView1.DefaultCellStyle.BackColor = Color.White;
        dataGridView1.DefaultCellStyle.Font = new Font("Segoe UI", 10F);
        dataGridView1.DefaultCellStyle.ForeColor = Color.FromArgb(73, 80, 87);
        dataGridView1.DefaultCellStyle.Padding = new Padding(10, 5, 10, 5);        dataGridView1.DefaultCellStyle.SelectionBackColor = Color.FromArgb(0, 123, 255, 25);
        dataGridView1.DefaultCellStyle.SelectionForeColor = Color.FromArgb(73, 80, 87);
        dataGridView1.EnableHeadersVisualStyles = false;
        dataGridView1.GridColor = Color.FromArgb(222, 226, 230);
        dataGridView1.Location = new Point(20, 20);
        dataGridView1.MultiSelect = false;
        dataGridView1.Name = "dataGridView1";
        dataGridView1.ReadOnly = true;
        dataGridView1.RowHeadersVisible = false;
        dataGridView1.RowTemplate.Height = 35;
        dataGridView1.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
        dataGridView1.Size = new Size(984, 420);
        dataGridView1.TabIndex = 4;
        // 
        // columnName
        // 
        columnName.HeaderText = "Название дисциплины";
        columnName.Name = "columnName";
        columnName.ReadOnly = true;
        columnName.Width = 400;
        // 
        // columnWorkload
        // 
        columnWorkload.HeaderText = "Трудоемкость";
        columnWorkload.Name = "columnWorkload";
        columnWorkload.ReadOnly = true;
        columnWorkload.Width = 150;
        // 
        // columnSignificance
        // 
        columnSignificance.HeaderText = "Коэффициент значимости";
        columnSignificance.Name = "columnSignificance";
        columnSignificance.ReadOnly = true;
        columnSignificance.Width = 200;
        // 
        // columnSemester
        // 
        columnSemester.HeaderText = "Семестр";
        columnSemester.Name = "columnSemester";
        columnSemester.ReadOnly = true;
        columnSemester.Width = 100;
        // 
        // statusPanel
        // 
        statusPanel.BackColor = Color.FromArgb(248, 249, 250);
        statusPanel.Controls.Add(progressBar);
        statusPanel.Controls.Add(statusLabel);
        statusPanel.Dock = DockStyle.Bottom;
        statusPanel.Location = new Point(0, 620);
        statusPanel.Name = "statusPanel";
        statusPanel.Padding = new Padding(20, 10, 20, 10);
        statusPanel.Size = new Size(1024, 44);
        statusPanel.TabIndex = 3;
        // 
        // statusLabel
        // 
        statusLabel.AutoSize = true;
        statusLabel.Font = new Font("Segoe UI", 9F);
        statusLabel.ForeColor = Color.FromArgb(108, 117, 125);
        statusLabel.Location = new Point(20, 15);
        statusLabel.Name = "statusLabel";
        statusLabel.Size = new Size(122, 15);
        statusLabel.TabIndex = 0;
        statusLabel.Text = "Готов к работе";
        // 
        // progressBar
        // 
        progressBar.Anchor = AnchorStyles.Top | AnchorStyles.Right;
        progressBar.Location = new Point(844, 12);
        progressBar.Name = "progressBar";
        progressBar.Size = new Size(160, 20);
        progressBar.Style = ProgressBarStyle.Continuous;
        progressBar.TabIndex = 1;
        progressBar.Visible = false;
        // 
        // Form1
        // 
        AutoScaleDimensions = new SizeF(96F, 96F);
        AutoScaleMode = AutoScaleMode.Dpi;
        BackColor = Color.White;
        ClientSize = new Size(1024, 664);
        Controls.Add(mainPanel);
        Controls.Add(statusPanel);
        Controls.Add(buttonPanel);
        Controls.Add(titleLabel);
        Font = new Font("Segoe UI", 9F);
        MinimumSize = new Size(800, 600);
        Name = "Form1";
        StartPosition = FormStartPosition.CenterScreen;
        Text = "Система оптимизации учебного плана";        ((System.ComponentModel.ISupportInitialize)dataGridView1).EndInit();
        statusPanel.ResumeLayout(false);
        statusPanel.PerformLayout();
        mainPanel.ResumeLayout(false);
        buttonPanel.ResumeLayout(false);
        ResumeLayout(false);
    }

    #endregion

    private Button button1;
    private Button button2;
    private Button button3;
    private DataGridView dataGridView1;
    private DataGridViewTextBoxColumn columnName;
    private DataGridViewTextBoxColumn columnWorkload;
    private DataGridViewTextBoxColumn columnSignificance;
    private DataGridViewTextBoxColumn columnSemester;
    private Label titleLabel;
    private Panel statusPanel;
    private ProgressBar progressBar;
    private Label statusLabel;
    private Panel mainPanel;
    private Panel buttonPanel;
}
