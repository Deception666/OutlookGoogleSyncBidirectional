using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace OutlookGoogleSync
{
   public partial class UserAccountForm : Form
   {
      public UserAccountForm(string user_account = "")
      {
         InitializeComponent();

         if (user_account != "")
         {
            userAccountTextBox.Text = user_account;
         }
      }

      private void OK_Click(object sender, EventArgs e)
      {
         UserAccount = userAccountTextBox.Text;
      }

      private void Cancel_Click(object sender, EventArgs e)
      {
         UserAccount = null;
      }

      private void userAccountTextBox_KeyPress(object sender, KeyPressEventArgs e)
      {
         if (e.KeyChar == (char)Keys.Enter)
         {
            e.Handled = true;

            OK.PerformClick();
         }
      }

      public string UserAccount { get; set; }
   }
}
