namespace TES_FUN2
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
            formsPlot1 = new ScottPlot.WinForms.FormsPlot();
            dateTimePicker1 = new DateTimePicker();
            dateTimePicker2 = new DateTimePicker();
            DateFilterBtn = new Button();
            resetBtn = new Button();
            SuspendLayout();
            // 
            // formsPlot1
            // 
            formsPlot1.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
            formsPlot1.DisplayScale = 1F;
            formsPlot1.Location = new Point(12, 99);
            formsPlot1.Name = "formsPlot1";
            formsPlot1.Size = new Size(796, 280);
            formsPlot1.TabIndex = 0;
            formsPlot1.Load += formsPlot1_Load;
            // 
            // dateTimePicker1
            // 
            dateTimePicker1.Location = new Point(12, 12);
            dateTimePicker1.Name = "dateTimePicker1";
            dateTimePicker1.Size = new Size(200, 23);
            dateTimePicker1.TabIndex = 1;
            dateTimePicker1.ValueChanged += dateTimePicker1_ValueChanged;
            // 
            // dateTimePicker2
            // 
            dateTimePicker2.Location = new Point(218, 12);
            dateTimePicker2.Name = "dateTimePicker2";
            dateTimePicker2.Size = new Size(200, 23);
            dateTimePicker2.TabIndex = 2;
            dateTimePicker2.ValueChanged += dateTimePicker2_ValueChanged;
            // 
            // DateFilterBtn
            // 
            DateFilterBtn.Location = new Point(424, 12);
            DateFilterBtn.Name = "DateFilterBtn";
            DateFilterBtn.Size = new Size(200, 23);
            DateFilterBtn.TabIndex = 3;
            DateFilterBtn.Text = "Fitrer";
            DateFilterBtn.UseVisualStyleBackColor = true;
            DateFilterBtn.Click += DateFilterBtn_Click;
            // 
            // resetBtn
            // 
            resetBtn.Location = new Point(630, 12);
            resetBtn.Name = "resetBtn";
            resetBtn.Size = new Size(178, 23);
            resetBtn.TabIndex = 4;
            resetBtn.Text = "reset";
            resetBtn.UseVisualStyleBackColor = true;
            resetBtn.Click += resetBtn_Click;
            // 
            // Form1
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(820, 469);
            Controls.Add(resetBtn);
            Controls.Add(DateFilterBtn);
            Controls.Add(dateTimePicker2);
            Controls.Add(dateTimePicker1);
            Controls.Add(formsPlot1);
            Name = "Form1";
            Text = "Form1";
            Load += Form1_Load;
            ResumeLayout(false);
        }

        #endregion

        private ScottPlot.WinForms.FormsPlot formsPlot1;
        private DateTimePicker dateTimePicker1;
        private DateTimePicker dateTimePicker2;
        private Button DateFilterBtn;
        private Button resetBtn;
    }
}
