using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Hotel1
{
    public partial class AuthorizationForm : Form
    {
        AllModel<User> users = new AllModel<User>("Users");
        public AuthorizationForm()
        {
            InitializeComponent();
        }
        public string GetHash(string input)
        {
            var md5 = MD5.Create();
            var hash = md5.ComputeHash(Encoding.UTF8.GetBytes(input));

            return Convert.ToBase64String(hash);
        }

        private async void btnVxod_Click(object sender, EventArgs e)
        {
            var pass = GetHash(passBox.Text);

            if (users.Objs.Any(user => user.Login == loginBox.Text && user.Password == pass))
            {
                var user = users.Objs.FirstOrDefault(user => user.Login == loginBox.Text && user.Password == pass);
                switch ((await new Role { Id = (int)user.RoleId }.Get()).Name)
                {
                    case "Администратор":
                        MainForm mainform = new MainForm();
                        this.Hide();
                        mainform.Show();
                        break;
                    case "Бухгалтер":
                        MainForm mainform2 = new MainForm();
                        this.Hide();
                        mainform2.Show();
                        mainform2.tabControl1.Visible = false;
                        mainform2.tabControl3.Visible = false;
                        mainform2.tabControl7.Visible = false;
                        mainform2.tabControl8.Visible = false;
                        mainform2.tabControl5.Visible = false;
                        mainform2.tabControl4.Visible = false;
                        break;
                    default:
                        return;
                }
            }
        }

        private void выходToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void checkBox1_Click(object sender, EventArgs e)
        {
            if (checkBox1.Checked)
            {
                passBox.PasswordChar = (char)0;
            }
            else
            {
                passBox.PasswordChar = '*';
            }
        }

        private void AuthorizationForm_FormClosed(object sender, FormClosedEventArgs e)
        {
            Application.Exit();
        }

        private void оПрограммеToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MessageBox.Show(
   "Программа для гостиниц",
   "Информация",
   MessageBoxButtons.OK,
   MessageBoxIcon.Information,
   MessageBoxDefaultButton.Button2,
   MessageBoxOptions.DefaultDesktopOnly);
        }
    }
}
