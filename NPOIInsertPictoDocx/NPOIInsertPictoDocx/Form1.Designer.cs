namespace NPOIInsertPictoDocx
{
    partial class Form1
    {
        /// <summary>
        /// 必需的设计器变量。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 清理所有正在使用的资源。
        /// </summary>
        /// <param name="disposing">如果应释放托管资源，为 true；否则为 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows 窗体设计器生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            this.button1 = new System.Windows.Forms.Button();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.button2_Besides = new System.Windows.Forms.Button();
            this.button2_Largest = new System.Windows.Forms.Button();
            this.button2_Right = new System.Windows.Forms.Button();
            this.button2_Left = new System.Windows.Forms.Button();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.button3_Largest = new System.Windows.Forms.Button();
            this.button3_Right = new System.Windows.Forms.Button();
            this.button3_Left = new System.Windows.Forms.Button();
            this.button3_Besides = new System.Windows.Forms.Button();
            this.groupBox3 = new System.Windows.Forms.GroupBox();
            this.button4_Largest = new System.Windows.Forms.Button();
            this.button4_Right = new System.Windows.Forms.Button();
            this.button4_Left = new System.Windows.Forms.Button();
            this.button4_Besides = new System.Windows.Forms.Button();
            this.button2 = new System.Windows.Forms.Button();
            this.button3 = new System.Windows.Forms.Button();
            this.button4 = new System.Windows.Forms.Button();
            this.groupBox1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.groupBox3.SuspendLayout();
            this.SuspendLayout();
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(18, 2);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(102, 32);
            this.button1.TabIndex = 0;
            this.button1.Text = "inline";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.button2_Besides);
            this.groupBox1.Controls.Add(this.button2_Largest);
            this.groupBox1.Controls.Add(this.button2_Right);
            this.groupBox1.Controls.Add(this.button2_Left);
            this.groupBox1.Location = new System.Drawing.Point(12, 78);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(250, 91);
            this.groupBox1.TabIndex = 4;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "四周(Square)";
            // 
            // button2_Besides
            // 
            this.button2_Besides.Location = new System.Drawing.Point(6, 20);
            this.button2_Besides.Name = "button2_Besides";
            this.button2_Besides.Size = new System.Drawing.Size(104, 23);
            this.button2_Besides.TabIndex = 6;
            this.button2_Besides.Text = "Square-Besides";
            this.button2_Besides.UseVisualStyleBackColor = true;
            this.button2_Besides.Click += new System.EventHandler(this.button2_Besides_Click);
            // 
            // button2_Largest
            // 
            this.button2_Largest.Location = new System.Drawing.Point(131, 60);
            this.button2_Largest.Name = "button2_Largest";
            this.button2_Largest.Size = new System.Drawing.Size(104, 23);
            this.button2_Largest.TabIndex = 5;
            this.button2_Largest.Text = "Square-Largest";
            this.button2_Largest.UseVisualStyleBackColor = true;
            this.button2_Largest.Click += new System.EventHandler(this.button2_Largest_Click);
            // 
            // button2_Right
            // 
            this.button2_Right.Location = new System.Drawing.Point(6, 60);
            this.button2_Right.Name = "button2_Right";
            this.button2_Right.Size = new System.Drawing.Size(104, 23);
            this.button2_Right.TabIndex = 4;
            this.button2_Right.Text = "Square-Right";
            this.button2_Right.UseVisualStyleBackColor = true;
            this.button2_Right.Click += new System.EventHandler(this.button2_Right_Click);
            // 
            // button2_Left
            // 
            this.button2_Left.Location = new System.Drawing.Point(131, 20);
            this.button2_Left.Name = "button2_Left";
            this.button2_Left.Size = new System.Drawing.Size(104, 23);
            this.button2_Left.TabIndex = 3;
            this.button2_Left.Text = "Square-Left";
            this.button2_Left.UseVisualStyleBackColor = true;
            this.button2_Left.Click += new System.EventHandler(this.button2_Left_Click);
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.button3_Largest);
            this.groupBox2.Controls.Add(this.button3_Right);
            this.groupBox2.Controls.Add(this.button3_Left);
            this.groupBox2.Controls.Add(this.button3_Besides);
            this.groupBox2.Location = new System.Drawing.Point(14, 175);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(248, 95);
            this.groupBox2.TabIndex = 5;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "紧密(Tight)";
            // 
            // button3_Largest
            // 
            this.button3_Largest.Location = new System.Drawing.Point(129, 61);
            this.button3_Largest.Name = "button3_Largest";
            this.button3_Largest.Size = new System.Drawing.Size(102, 23);
            this.button3_Largest.TabIndex = 6;
            this.button3_Largest.Text = "Tight_Largest";
            this.button3_Largest.UseVisualStyleBackColor = true;
            this.button3_Largest.Click += new System.EventHandler(this.button3_Largest_Click);
            // 
            // button3_Right
            // 
            this.button3_Right.Location = new System.Drawing.Point(6, 61);
            this.button3_Right.Name = "button3_Right";
            this.button3_Right.Size = new System.Drawing.Size(102, 23);
            this.button3_Right.TabIndex = 5;
            this.button3_Right.Text = "Tight_Right";
            this.button3_Right.UseVisualStyleBackColor = true;
            this.button3_Right.Click += new System.EventHandler(this.button3_Right_Click);
            // 
            // button3_Left
            // 
            this.button3_Left.Location = new System.Drawing.Point(129, 20);
            this.button3_Left.Name = "button3_Left";
            this.button3_Left.Size = new System.Drawing.Size(102, 23);
            this.button3_Left.TabIndex = 4;
            this.button3_Left.Text = "Tight_Left";
            this.button3_Left.UseVisualStyleBackColor = true;
            this.button3_Left.Click += new System.EventHandler(this.button3_Left_Click);
            // 
            // button3_Besides
            // 
            this.button3_Besides.Location = new System.Drawing.Point(6, 20);
            this.button3_Besides.Name = "button3_Besides";
            this.button3_Besides.Size = new System.Drawing.Size(102, 23);
            this.button3_Besides.TabIndex = 3;
            this.button3_Besides.Text = "Tight_Besides";
            this.button3_Besides.UseVisualStyleBackColor = true;
            this.button3_Besides.Click += new System.EventHandler(this.button3_Besides_Click);
            // 
            // groupBox3
            // 
            this.groupBox3.Controls.Add(this.button4_Largest);
            this.groupBox3.Controls.Add(this.button4_Right);
            this.groupBox3.Controls.Add(this.button4_Left);
            this.groupBox3.Controls.Add(this.button4_Besides);
            this.groupBox3.Location = new System.Drawing.Point(12, 276);
            this.groupBox3.Name = "groupBox3";
            this.groupBox3.Size = new System.Drawing.Size(249, 84);
            this.groupBox3.TabIndex = 6;
            this.groupBox3.TabStop = false;
            this.groupBox3.Text = "穿越(Through)";
            // 
            // button4_Largest
            // 
            this.button4_Largest.Location = new System.Drawing.Point(128, 55);
            this.button4_Largest.Name = "button4_Largest";
            this.button4_Largest.Size = new System.Drawing.Size(115, 23);
            this.button4_Largest.TabIndex = 7;
            this.button4_Largest.Text = "Through-Largest";
            this.button4_Largest.UseVisualStyleBackColor = true;
            this.button4_Largest.Click += new System.EventHandler(this.button4_Largest_Click);
            // 
            // button4_Right
            // 
            this.button4_Right.Location = new System.Drawing.Point(8, 55);
            this.button4_Right.Name = "button4_Right";
            this.button4_Right.Size = new System.Drawing.Size(115, 23);
            this.button4_Right.TabIndex = 6;
            this.button4_Right.Text = "Through-Right";
            this.button4_Right.UseVisualStyleBackColor = true;
            this.button4_Right.Click += new System.EventHandler(this.button4_Right_Click);
            // 
            // button4_Left
            // 
            this.button4_Left.Location = new System.Drawing.Point(129, 20);
            this.button4_Left.Name = "button4_Left";
            this.button4_Left.Size = new System.Drawing.Size(115, 23);
            this.button4_Left.TabIndex = 5;
            this.button4_Left.Text = "Through-Left";
            this.button4_Left.UseVisualStyleBackColor = true;
            this.button4_Left.Click += new System.EventHandler(this.button4_Left_Click);
            // 
            // button4_Besides
            // 
            this.button4_Besides.Location = new System.Drawing.Point(8, 20);
            this.button4_Besides.Name = "button4_Besides";
            this.button4_Besides.Size = new System.Drawing.Size(115, 23);
            this.button4_Besides.TabIndex = 4;
            this.button4_Besides.Text = "Through-Besides";
            this.button4_Besides.UseVisualStyleBackColor = true;
            this.button4_Besides.Click += new System.EventHandler(this.button4_Besides_Click);
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(145, 2);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(102, 32);
            this.button2.TabIndex = 7;
            this.button2.Text = "TopAndBottom";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // button3
            // 
            this.button3.Location = new System.Drawing.Point(20, 45);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(102, 27);
            this.button3.TabIndex = 8;
            this.button3.Text = "None(上方)";
            this.button3.UseVisualStyleBackColor = true;
            this.button3.Click += new System.EventHandler(this.button3_Click);
            // 
            // button4
            // 
            this.button4.Location = new System.Drawing.Point(145, 45);
            this.button4.Name = "button4";
            this.button4.Size = new System.Drawing.Size(102, 27);
            this.button4.TabIndex = 9;
            this.button4.Text = "None(下方)";
            this.button4.UseVisualStyleBackColor = true;
            this.button4.Click += new System.EventHandler(this.button4_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(277, 363);
            this.Controls.Add(this.button4);
            this.Controls.Add(this.button3);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.groupBox3);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.button1);
            this.Name = "Form1";
            this.Text = "Form1";
            this.groupBox1.ResumeLayout(false);
            this.groupBox2.ResumeLayout(false);
            this.groupBox3.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Button button2_Left;
        private System.Windows.Forms.Button button2_Largest;
        private System.Windows.Forms.Button button2_Right;
        private System.Windows.Forms.Button button2_Besides;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.Button button3_Besides;
        private System.Windows.Forms.Button button3_Largest;
        private System.Windows.Forms.Button button3_Right;
        private System.Windows.Forms.Button button3_Left;
        private System.Windows.Forms.GroupBox groupBox3;
        private System.Windows.Forms.Button button4_Besides;
        private System.Windows.Forms.Button button4_Largest;
        private System.Windows.Forms.Button button4_Right;
        private System.Windows.Forms.Button button4_Left;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.Button button3;
        private System.Windows.Forms.Button button4;
    }
}

