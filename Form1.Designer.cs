namespace TestOpt;

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

    #region Windows Form Designer generated code    /// <summary>
    ///  Required method for Designer support - do not modify
    ///  the contents of this method with the code editor.
    /// </summary>
    private void InitializeComponent()
    {
        this.btnOpenFile = new System.Windows.Forms.Button();
        this.dataGridView1 = new System.Windows.Forms.DataGridView();
        this.lblFileName = new System.Windows.Forms.Label();
        ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
        this.SuspendLayout();
        // 
        // btnOpenFile
        // 
        this.btnOpenFile.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point);
        this.btnOpenFile.Location = new System.Drawing.Point(12, 12);
        this.btnOpenFile.Name = "btnOpenFile";
        this.btnOpenFile.Size = new System.Drawing.Size(150, 35);
        this.btnOpenFile.TabIndex = 0;
        this.btnOpenFile.Text = "Открыть Excel файл";
        this.btnOpenFile.UseVisualStyleBackColor = true;
        this.btnOpenFile.Click += new System.EventHandler(this.btnOpenFile_Click);
        // 
        // dataGridView1
        // 
        this.dataGridView1.AllowUserToAddRows = false;
        this.dataGridView1.AllowUserToDeleteRows = false;
        this.dataGridView1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
        | System.Windows.Forms.AnchorStyles.Left) 
        | System.Windows.Forms.AnchorStyles.Right)));
        this.dataGridView1.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
        this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
        this.dataGridView1.Location = new System.Drawing.Point(12, 80);
        this.dataGridView1.Name = "dataGridView1";
        this.dataGridView1.RowHeadersWidth = 51;
        this.dataGridView1.Size = new System.Drawing.Size(776, 358);
        this.dataGridView1.TabIndex = 1;
        this.dataGridView1.CellValueChanged += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridView1_CellValueChanged);
        // 
        // lblFileName
        // 
        this.lblFileName.AutoSize = true;
        this.lblFileName.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point);
        this.lblFileName.Location = new System.Drawing.Point(180, 22);
        this.lblFileName.Name = "lblFileName";
        this.lblFileName.Size = new System.Drawing.Size(95, 15);
        this.lblFileName.TabIndex = 2;
        this.lblFileName.Text = "Файл не выбран";
        // 
        // Form1
        // 
        this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
        this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
        this.ClientSize = new System.Drawing.Size(800, 450);
        this.Controls.Add(this.lblFileName);
        this.Controls.Add(this.dataGridView1);
        this.Controls.Add(this.btnOpenFile);
        this.Name = "Form1";
        this.Text = "Excel Reader & Editor";
        ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
        this.ResumeLayout(false);
        this.PerformLayout();
    }

    #endregion

    private System.Windows.Forms.Button btnOpenFile;
    private System.Windows.Forms.DataGridView dataGridView1;
    private System.Windows.Forms.Label lblFileName;
}
