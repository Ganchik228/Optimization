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
        label1 = new Label();
        dataGridView1 = new DataGridView();
        columnName = new DataGridViewTextBoxColumn();
        columnWorkload = new DataGridViewTextBoxColumn();
        columnSignificance = new DataGridViewTextBoxColumn();
        columnSemester = new DataGridViewTextBoxColumn();
        ((System.ComponentModel.ISupportInitialize)dataGridView1).BeginInit();
        SuspendLayout();
        // 
        // button1
        // 
        button1.Location = new Point(12, 12);
        button1.Name = "button1";
        button1.Size = new Size(120, 30);
        button1.TabIndex = 0;
        button1.Text = "Загрузить данные";
        button1.UseVisualStyleBackColor = true;
        button1.Click += button1_Click;
        // 
        // button2
        // 
        button2.Location = new Point(668, 12);
        button2.Name = "button2";
        button2.Size = new Size(120, 30);
        button2.TabIndex = 1;
        button2.Text = "Закрыть";
        button2.UseVisualStyleBackColor = true;
        button2.Click += button2_Click;
        // 
        // button3
        // 
        button3.Location = new Point(138, 12);
        button3.Name = "button3";
        button3.Size = new Size(140, 30);
        button3.TabIndex = 2;
        button3.Text = "Сохранить результаты";
        button3.UseVisualStyleBackColor = true;
        button3.Click += button3_Click_1;
        // 
        // label1
        // 
        label1.AutoSize = true;
        label1.Location = new Point(12, 55);
        label1.Name = "label1";
        label1.Size = new Size(120, 15);
        label1.TabIndex = 3;
        label1.Text = "Целевая функция: -";
        // 
        // dataGridView1
        // 
        dataGridView1.AllowUserToAddRows = false;
        dataGridView1.AllowUserToDeleteRows = false;
        dataGridView1.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
        dataGridView1.Columns.AddRange(new DataGridViewColumn[] { columnName, columnWorkload, columnSignificance, columnSemester });
        dataGridView1.Location = new Point(12, 85);
        dataGridView1.Name = "dataGridView1";
        dataGridView1.ReadOnly = true;
        dataGridView1.Size = new Size(776, 350);
        dataGridView1.TabIndex = 4;
        // 
        // columnName
        // 
        columnName.HeaderText = "Название дисциплины";
        columnName.Name = "columnName";
        columnName.ReadOnly = true;
        columnName.Width = 300;
        // 
        // columnWorkload
        // 
        columnWorkload.HeaderText = "Трудоемкость";
        columnWorkload.Name = "columnWorkload";
        columnWorkload.ReadOnly = true;
        columnWorkload.Width = 120;
        // 
        // columnSignificance
        // 
        columnSignificance.HeaderText = "Коэффициент значимости";
        columnSignificance.Name = "columnSignificance";
        columnSignificance.ReadOnly = true;
        columnSignificance.Width = 150;
        // 
        // columnSemester
        // 
        columnSemester.HeaderText = "Семестр";
        columnSemester.Name = "columnSemester";
        columnSemester.ReadOnly = true;
        columnSemester.Width = 80;
        // 
        // Form1
        // 
        AutoScaleDimensions = new SizeF(7F, 15F);
        AutoScaleMode = AutoScaleMode.Font;
        ClientSize = new Size(800, 450);
        Controls.Add(dataGridView1);
        Controls.Add(label1);
        Controls.Add(button3);
        Controls.Add(button2);
        Controls.Add(button1);
        Name = "Form1";
        Text = "Оптимизация учебного плана";
        ((System.ComponentModel.ISupportInitialize)dataGridView1).EndInit();
        ResumeLayout(false);
        PerformLayout();
    }

    #endregion

    private Button button1;
    private Button button2;
    private Button button3;
    private Label label1;
    private DataGridView dataGridView1;
    private DataGridViewTextBoxColumn columnName;
    private DataGridViewTextBoxColumn columnWorkload;
    private DataGridViewTextBoxColumn columnSignificance;
    private DataGridViewTextBoxColumn columnSemester;
}
