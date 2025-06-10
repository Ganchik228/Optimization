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
        progressBar = new ProgressBar();
        statusLabel = new Label();
        titleLabel = new Label();
        headerPanel = new Panel();
        mainPanel = new Panel();
        footerPanel = new Panel();
        infoLabel = new Label();
        targetSumsPanel = new Panel();
        targetSumsLabel = new Label();
        sem12Label = new Label();
        sem34Label = new Label();
        sem56Label = new Label();
        sem78Label = new Label();
        textBoxSem12 = new TextBox();
        textBoxSem34 = new TextBox();
        textBoxSem56 = new TextBox();
        textBoxSem78 = new TextBox();
        SuspendLayout();
          // 
        // headerPanel
        // 
        headerPanel.BackColor = Color.FromArgb(45, 125, 255);
        headerPanel.Dock = DockStyle.Top;
        headerPanel.Height = 100;
        headerPanel.Controls.Add(titleLabel);
        headerPanel.Controls.Add(infoLabel);
          // 
        // titleLabel
        // 
        titleLabel.Font = new Font("Segoe UI", 18F, FontStyle.Bold);
        titleLabel.ForeColor = Color.White;
        titleLabel.Location = new Point(30, 20);
        titleLabel.Size = new Size(520, 35);
        titleLabel.Text = "🎓 Система оптимизации учебного плана";
        titleLabel.TextAlign = ContentAlignment.MiddleLeft;
          // 
        // infoLabel
        // 
        infoLabel.Font = new Font("Segoe UI", 10F);
        infoLabel.ForeColor = Color.FromArgb(220, 220, 220);
        infoLabel.Location = new Point(30, 55);
        infoLabel.Size = new Size(500, 20);
        infoLabel.Text = "Загрузите Excel файл для оптимизации расписания";
        infoLabel.TextAlign = ContentAlignment.MiddleLeft;
        
        // 
        // mainPanel
        // 
        mainPanel.BackColor = Color.FromArgb(248, 249, 250);
        mainPanel.Dock = DockStyle.Fill;
        mainPanel.Padding = new Padding(30);
        mainPanel.Controls.Add(button1);
        mainPanel.Controls.Add(button3);
        mainPanel.Controls.Add(button2);
        mainPanel.Controls.Add(targetSumsPanel);
        
        // 
        // targetSumsPanel
        // 
        targetSumsPanel.BackColor = Color.White;
        targetSumsPanel.BorderStyle = BorderStyle.FixedSingle;
        targetSumsPanel.Location = new Point(40, 120); // Adjusted Y to be below buttons
        targetSumsPanel.Size = new Size(560, 120); // Width to fit in mainPanel padding
        targetSumsPanel.Controls.Add(targetSumsLabel);
        targetSumsPanel.Controls.Add(sem12Label);
        targetSumsPanel.Controls.Add(sem34Label);
        targetSumsPanel.Controls.Add(sem56Label);
        targetSumsPanel.Controls.Add(sem78Label);
        targetSumsPanel.Controls.Add(textBoxSem12);
        targetSumsPanel.Controls.Add(textBoxSem34);
        targetSumsPanel.Controls.Add(textBoxSem56);
        targetSumsPanel.Controls.Add(textBoxSem78);
        
        // 
        // targetSumsLabel
        // 
        targetSumsLabel.Font = new Font("Segoe UI", 11F, FontStyle.Bold);
        targetSumsLabel.ForeColor = Color.FromArgb(50, 50, 50);
        targetSumsLabel.Location = new Point(15, 10);
        targetSumsLabel.Size = new Size(400, 25);
        targetSumsLabel.Text = "Целевые суммы для пар семестров:";
        targetSumsLabel.TextAlign = ContentAlignment.MiddleLeft;
        
        // 
        // sem12Label
        // 
        sem12Label.Font = new Font("Segoe UI", 10F);
        sem12Label.ForeColor = Color.FromArgb(100, 100, 100);
        sem12Label.Location = new Point(15, 45);
        sem12Label.Size = new Size(120, 25);
        sem12Label.Text = "Семестры 1-2:";
        sem12Label.TextAlign = ContentAlignment.MiddleLeft;
        
        // 
        // textBoxSem12
        // 
        textBoxSem12.Font = new Font("Segoe UI", 10F);
        textBoxSem12.Location = new Point(140, 43);
        textBoxSem12.Size = new Size(80, 25);
        textBoxSem12.Text = "60.5";
        textBoxSem12.TextAlign = HorizontalAlignment.Center;
        
        // 
        // sem34Label
        // 
        sem34Label.Font = new Font("Segoe UI", 10F);
        sem34Label.ForeColor = Color.FromArgb(100, 100, 100);
        sem34Label.Location = new Point(240, 45);
        sem34Label.Size = new Size(120, 25);
        sem34Label.Text = "Семестры 3-4:";
        sem34Label.TextAlign = ContentAlignment.MiddleLeft;
        
        // 
        // textBoxSem34
        // 
        textBoxSem34.Font = new Font("Segoe UI", 10F);
        textBoxSem34.Location = new Point(365, 43);
        textBoxSem34.Size = new Size(80, 25);
        textBoxSem34.Text = "59.5";
        textBoxSem34.TextAlign = HorizontalAlignment.Center;
        
        // 
        // sem56Label
        // 
        sem56Label.Font = new Font("Segoe UI", 10F);
        sem56Label.ForeColor = Color.FromArgb(100, 100, 100);
        sem56Label.Location = new Point(15, 80);
        sem56Label.Size = new Size(120, 25);
        sem56Label.Text = "Семестры 5-6:";
        sem56Label.TextAlign = ContentAlignment.MiddleLeft;
        
        // 
        // textBoxSem56
        // 
        textBoxSem56.Font = new Font("Segoe UI", 10F);
        textBoxSem56.Location = new Point(140, 78);
        textBoxSem56.Size = new Size(80, 25);
        textBoxSem56.Text = "60";
        textBoxSem56.TextAlign = HorizontalAlignment.Center;
        
        // 
        // sem78Label
        // 
        sem78Label.Font = new Font("Segoe UI", 10F);
        sem78Label.ForeColor = Color.FromArgb(100, 100, 100);
        sem78Label.Location = new Point(240, 80);
        sem78Label.Size = new Size(120, 25);
        sem78Label.Text = "Семестры 7-8:";
        sem78Label.TextAlign = ContentAlignment.MiddleLeft;
        
        // 
        // textBoxSem78
        // 
        textBoxSem78.Font = new Font("Segoe UI", 10F);
        textBoxSem78.Location = new Point(365, 78);
        textBoxSem78.Size = new Size(80, 25);
        textBoxSem78.Text = "60";
        textBoxSem78.TextAlign = HorizontalAlignment.Center;
        
        // 
        // footerPanel
        // 
        footerPanel.BackColor = Color.FromArgb(255, 255, 255); // White background
        footerPanel.Dock = DockStyle.Bottom;
        footerPanel.Height = 70;
        footerPanel.Padding = new Padding(30, 0, 30, 0); // Add padding for controls
        footerPanel.Controls.Add(progressBar);
        footerPanel.Controls.Add(statusLabel);
        // 
        // button1 (Загрузить)
        // 
        button1.BackColor = Color.FromArgb(52, 199, 89);
        button1.FlatAppearance.BorderSize = 0;
        button1.FlatAppearance.MouseDownBackColor = Color.FromArgb(36, 160, 67);
        button1.FlatAppearance.MouseOverBackColor = Color.FromArgb(76, 217, 100);
        button1.FlatStyle = FlatStyle.Flat;
        button1.Font = new Font("Segoe UI", 12F, FontStyle.Bold);
        button1.ForeColor = Color.White;
        button1.Location = new Point(40, 40); // Position within mainPanel
        button1.Name = "button1";
        button1.Size = new Size(200, 60);
        button1.TabIndex = 0;
        button1.Text = "📁 Загрузить файл";
        button1.UseVisualStyleBackColor = false;
        button1.Cursor = Cursors.Hand;
        button1.Click += button1_Click;
        // 
        // button3 (Сохранить)
        // 
        button3.BackColor = Color.FromArgb(0, 122, 255);
        button3.FlatAppearance.BorderSize = 0;
        button3.FlatAppearance.MouseDownBackColor = Color.FromArgb(0, 86, 179);
        button3.FlatAppearance.MouseOverBackColor = Color.FromArgb(38, 143, 255);
        button3.FlatStyle = FlatStyle.Flat;
        button3.Font = new Font("Segoe UI", 12F, FontStyle.Bold);
        button3.ForeColor = Color.White;
        button3.Location = new Point(260, 40); // Position within mainPanel
        button3.Name = "button3";
        button3.Size = new Size(180, 60);
        button3.TabIndex = 2; // Keep TabIndex logical
        button3.Text = "💾 Сохранить";
        button3.UseVisualStyleBackColor = false;
        button3.Cursor = Cursors.Hand;
        button3.Click += button3_Click_1;
        // 
        // button2 (Закрыть)
        // 
        button2.BackColor = Color.FromArgb(255, 59, 48);
        button2.FlatAppearance.BorderSize = 0;
        button2.FlatAppearance.MouseDownBackColor = Color.FromArgb(200, 35, 51);
        button2.FlatAppearance.MouseOverBackColor = Color.FromArgb(255, 105, 97);
        button2.FlatStyle = FlatStyle.Flat;
        button2.Font = new Font("Segoe UI", 12F, FontStyle.Bold);
        button2.ForeColor = Color.White;
        button2.Location = new Point(460, 40); // Position within mainPanel
        button2.Name = "button2";
        button2.Size = new Size(140, 60);
        button2.TabIndex = 1; // Keep TabIndex logical
        button2.Text = "❌ Закрыть";
        button2.UseVisualStyleBackColor = false;
        button2.Cursor = Cursors.Hand;
        button2.Click += button2_Click;
        
        // 
        // progressBar
        // 
        progressBar.Location = new Point(30, 15); // Position within footerPanel
        progressBar.Size = new Size(520, 20); // Adjust size based on footerPanel padding
        progressBar.Style = ProgressBarStyle.Continuous;
        progressBar.TabIndex = 1; // TabIndex within footerPanel
        progressBar.ForeColor = Color.FromArgb(0, 122, 255);
        progressBar.BackColor = Color.FromArgb(240, 240, 240);
        progressBar.Visible = false;
        
        // 
        // statusLabel
        // 
        statusLabel.Font = new Font("Segoe UI", 10F, FontStyle.Bold);
        statusLabel.ForeColor = Color.FromArgb(100, 100, 100);
        statusLabel.Location = new Point(30, 40); // Position within footerPanel
        statusLabel.Size = new Size(520, 20); // Adjust size based on footerPanel padding
        statusLabel.TextAlign = ContentAlignment.MiddleLeft;
        statusLabel.Text = "✅ Готов к работе";
        // 
        // Form1
        // 
        AutoScaleDimensions = new SizeF(96F, 96F); // Standard DPI scaling
        AutoScaleMode = AutoScaleMode.Dpi;
        BackColor = Color.FromArgb(245, 247, 250); // Light gray background
        ClientSize = new Size(640, 420); // Fixed client size
        Controls.Add(mainPanel); // Add mainPanel first so it's behind header/footer
        Controls.Add(headerPanel);
        Controls.Add(footerPanel);
        Font = new Font("Segoe UI", 9F); // Default font
        FormBorderStyle = FormBorderStyle.FixedSingle;
        MaximizeBox = false;
        MinimumSize = new Size(640, 420); // Prevent resizing smaller
        Name = "Form1";
        StartPosition = FormStartPosition.CenterScreen;
        Text = "🎓 Система оптимизации учебного плана"; // Window title
        Icon = null; // Set your icon here if you have one
        
        headerPanel.ResumeLayout(false);
        mainPanel.ResumeLayout(false);
        targetSumsPanel.ResumeLayout(false);
        targetSumsPanel.PerformLayout();
        footerPanel.ResumeLayout(false);
        ResumeLayout(false);
    }

    #endregion

    private Button button1;
    private Button button2;
    private Button button3;
    private ProgressBar progressBar;
    private Label statusLabel;
    private Label titleLabel;
    private Panel headerPanel;
    private Panel mainPanel;
    private Panel footerPanel;
    private Label infoLabel;
    private Panel targetSumsPanel;
    private Label targetSumsLabel;
    private Label sem12Label;
    private Label sem34Label;
    private Label sem56Label;
    private Label sem78Label;
    private TextBox textBoxSem12;
    private TextBox textBoxSem34;
    private TextBox textBoxSem56;
    private TextBox textBoxSem78;
}
