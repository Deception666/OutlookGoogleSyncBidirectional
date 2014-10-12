namespace OutlookGoogleSync
{
    partial class UserAccountForm
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
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
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
         System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(UserAccountForm));
         this.userAccountTextBox = new System.Windows.Forms.TextBox();
         this.OK = new System.Windows.Forms.Button();
         this.Cancel = new System.Windows.Forms.Button();
         this.SuspendLayout();
         // 
         // userAccountTextBox
         // 
         this.userAccountTextBox.Location = new System.Drawing.Point(12, 12);
         this.userAccountTextBox.Name = "userAccountTextBox";
         this.userAccountTextBox.Size = new System.Drawing.Size(350, 20);
         this.userAccountTextBox.TabIndex = 0;
         this.userAccountTextBox.Text = "user@gmail.com";
         this.userAccountTextBox.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.userAccountTextBox_KeyPress);
         // 
         // OK
         // 
         this.OK.DialogResult = System.Windows.Forms.DialogResult.OK;
         this.OK.Location = new System.Drawing.Point(368, 10);
         this.OK.Name = "OK";
         this.OK.Size = new System.Drawing.Size(75, 23);
         this.OK.TabIndex = 1;
         this.OK.Text = "OK";
         this.OK.UseVisualStyleBackColor = true;
         this.OK.Click += new System.EventHandler(this.OK_Click);
         // 
         // Cancel
         // 
         this.Cancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
         this.Cancel.Location = new System.Drawing.Point(449, 10);
         this.Cancel.Name = "Cancel";
         this.Cancel.Size = new System.Drawing.Size(75, 23);
         this.Cancel.TabIndex = 2;
         this.Cancel.Text = "Cancel";
         this.Cancel.UseVisualStyleBackColor = true;
         this.Cancel.Click += new System.EventHandler(this.Cancel_Click);
         // 
         // UserAccountForm
         // 
         this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
         this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
         this.ClientSize = new System.Drawing.Size(531, 41);
         this.Controls.Add(this.Cancel);
         this.Controls.Add(this.OK);
         this.Controls.Add(this.userAccountTextBox);
         this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
         this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
         this.MaximizeBox = false;
         this.MinimizeBox = false;
         this.Name = "UserAccountForm";
         this.ShowInTaskbar = false;
         this.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide;
         this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
         this.Text = "Enter Google User Account";
         this.ResumeLayout(false);
         this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox userAccountTextBox;
        private System.Windows.Forms.Button OK;
        private System.Windows.Forms.Button Cancel;
    }
}