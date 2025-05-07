namespace WinFormsApp1
{
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
            this.knopkaOtkryt = new System.Windows.Forms.Button();
            this.knopkaZapisat = new System.Windows.Forms.Button();
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.textBoxText = new System.Windows.Forms.TextBox();
            this.textBoxYacheika = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.SuspendLayout();
            // 
            // knopkaOtkryt
            // 
            this.knopkaOtkryt.Location = new System.Drawing.Point(20, 20);
            this.knopkaOtkryt.Name = "knopkaOtkryt";
            this.knopkaOtkryt.Size = new System.Drawing.Size(120, 30);
            this.knopkaOtkryt.TabIndex = 0;
            this.knopkaOtkryt.Text = "Открыть Excel";
            this.knopkaOtkryt.UseVisualStyleBackColor = true;
            this.knopkaOtkryt.Click += new System.EventHandler(this.OpenFile_Click);
            // 
            // knopkaZapisat
            // 
            this.knopkaZapisat.Location = new System.Drawing.Point(240, 310);
            this.knopkaZapisat.Name = "knopkaZapisat";
            this.knopkaZapisat.Size = new System.Drawing.Size(120, 30);
            this.knopkaZapisat.TabIndex = 1;
            this.knopkaZapisat.Text = "Записать в Excel";
            this.knopkaZapisat.UseVisualStyleBackColor = true;
            this.knopkaZapisat.Click += new System.EventHandler(this.Change_Click);
            // 
            // dataGridView1
            // 
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Location = new System.Drawing.Point(20, 60);
            dataGridView1.EditMode = DataGridViewEditMode.EditOnKeystroke;
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.Size = new System.Drawing.Size(500, 200);
            this.dataGridView1.TabIndex = 2;
            // 
            // textBoxText
            // 
            this.textBoxText.Location = new System.Drawing.Point(120, 280);
            this.textBoxText.Name = "textBoxText";
            this.textBoxText.Size = new System.Drawing.Size(200, 22);
            this.textBoxText.TabIndex = 3;
            // 
            // textBoxYacheika
            // 
            this.textBoxYacheika.Location = new System.Drawing.Point(120, 310);
            this.textBoxYacheika.Name = "textBoxYacheika";
            this.textBoxYacheika.Size = new System.Drawing.Size(100, 22);
            this.textBoxYacheika.TabIndex = 4;
            // 
            // label1
            // 
            this.label1.Location = new System.Drawing.Point(20, 280);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(100, 22);
            this.label1.TabIndex = 5;
            this.label1.Text = "Текст:";
            this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // label2
            // 
            this.label2.Location = new System.Drawing.Point(20, 310);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(100, 22);
            this.label2.TabIndex = 6;
            this.label2.Text = "Ячейка:";
            this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(550, 370);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.textBoxYacheika);
            this.Controls.Add(this.textBoxText);
            this.Controls.Add(this.dataGridView1);
            this.Controls.Add(this.knopkaZapisat);
            this.Controls.Add(this.knopkaOtkryt);
            this.Name = "Form1";
            this.Text = "Работа с Excel";
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();
        }

        #endregion

        private System.Windows.Forms.Button knopkaOtkryt;
        private System.Windows.Forms.Button knopkaZapisat;
        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.TextBox textBoxText;
        private System.Windows.Forms.TextBox textBoxYacheika;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
    }
}

