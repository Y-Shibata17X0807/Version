namespace 原価試算書作成ツール
{
    partial class Form1
    {
        /// <summary>
        /// 必要なデザイナー変数です。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 使用中のリソースをすべてクリーンアップします。
        /// </summary>
        /// <param name="disposing">マネージ リソースを破棄する場合は true を指定し、その他の場合は false を指定します。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows フォーム デザイナーで生成されたコード

        /// <summary>
        /// デザイナー サポートに必要なメソッドです。このメソッドの内容を
        /// コード エディターで変更しないでください。
        /// </summary>
        private void InitializeComponent()
        {
            this.Output = new System.Windows.Forms.TextBox();
            this.Dialog2 = new System.Windows.Forms.Button();
            this.label3 = new System.Windows.Forms.Label();
            this.CreateBtn = new System.Windows.Forms.Button();
            this.Dialog1 = new System.Windows.Forms.Button();
            this.Input = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // Output
            // 
            this.Output.Font = new System.Drawing.Font("メイリオ", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.Output.Location = new System.Drawing.Point(12, 125);
            this.Output.Name = "Output";
            this.Output.Size = new System.Drawing.Size(499, 27);
            this.Output.TabIndex = 2;
            // 
            // Dialog2
            // 
            this.Dialog2.Font = new System.Drawing.Font("メイリオ", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.Dialog2.Location = new System.Drawing.Point(517, 125);
            this.Dialog2.Name = "Dialog2";
            this.Dialog2.Size = new System.Drawing.Size(44, 27);
            this.Dialog2.TabIndex = 5;
            this.Dialog2.Text = "...";
            this.Dialog2.UseVisualStyleBackColor = true;
            this.Dialog2.Click += new System.EventHandler(this.Dialog2Btn_Click);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("メイリオ", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.label3.Location = new System.Drawing.Point(12, 102);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(289, 20);
            this.label3.TabIndex = 8;
            this.label3.Text = "作成したExcelシートの保存先を選択して下さい";
            // 
            // CreateBtn
            // 
            this.CreateBtn.Font = new System.Drawing.Font("メイリオ", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.CreateBtn.Location = new System.Drawing.Point(67, 174);
            this.CreateBtn.Name = "CreateBtn";
            this.CreateBtn.Size = new System.Drawing.Size(416, 39);
            this.CreateBtn.TabIndex = 9;
            this.CreateBtn.Text = "試算書作成開始";
            this.CreateBtn.UseVisualStyleBackColor = true;
            this.CreateBtn.Click += new System.EventHandler(this.CreateBtn_Click);
            // 
            // Dialog1
            // 
            this.Dialog1.Font = new System.Drawing.Font("メイリオ", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.Dialog1.Location = new System.Drawing.Point(517, 47);
            this.Dialog1.Name = "Dialog1";
            this.Dialog1.Size = new System.Drawing.Size(44, 27);
            this.Dialog1.TabIndex = 11;
            this.Dialog1.Text = "...";
            this.Dialog1.UseVisualStyleBackColor = true;
            this.Dialog1.Click += new System.EventHandler(this.Dialog1_Click);
            // 
            // Input
            // 
            this.Input.Font = new System.Drawing.Font("メイリオ", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.Input.Location = new System.Drawing.Point(12, 47);
            this.Input.Multiline = true;
            this.Input.Name = "Input";
            this.Input.Size = new System.Drawing.Size(499, 28);
            this.Input.TabIndex = 10;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("メイリオ", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.label1.Location = new System.Drawing.Point(12, 24);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(230, 20);
            this.label1.TabIndex = 12;
            this.label1.Text = "見積もり情報シートを選択して下さい";
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(587, 225);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.Dialog1);
            this.Controls.Add(this.Input);
            this.Controls.Add(this.CreateBtn);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.Dialog2);
            this.Controls.Add(this.Output);
            this.Name = "Form1";
            this.Text = "原価試算作成ツール V1.0.0.0 Pre-Alpha";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.TextBox Output;
        private System.Windows.Forms.Button Dialog2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Button CreateBtn;
        private System.Windows.Forms.Button Dialog1;
        private System.Windows.Forms.TextBox Input;
        private System.Windows.Forms.Label label1;
    }
}

