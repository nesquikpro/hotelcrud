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
            MainForm mainform = new MainForm();
            var pass = GetHash(passBox.Text);

            if (users.Objs.Any(user => user.Login == loginBox.Text && user.Password == pass))
            {
                var user = users.Objs.FirstOrDefault(user => user.Login == loginBox.Text && user.Password == pass);
                switch ((await new Role { Id = (int)user.RoleId }.Get()).Name)
                {
                    case "Администратор":
                        this.Hide();
                        mainform.tabControl2.TabPages.RemoveByKey("tabPage7");
                        mainform.Show();                    
                        break;
                    case "Бухгалтер":
                        this.Hide();
                        mainform.tabControl2.TabPages.RemoveByKey("tabPage6");
                        mainform.tabControl2.TabPages.RemoveByKey("tabPage3");
                        mainform.tabControl2.TabPages.RemoveByKey("tabPage4");
                        mainform.tabControl2.TabPages.RemoveByKey("tabPage8");
                        mainform.tabControl2.TabPages.RemoveByKey("tabPage18");
                        mainform.tabControl2.TabPages.RemoveByKey("tabPage5");
                        mainform.Show();
                        break;
                    case "Кадровик":
                        this.Hide();
                        mainform.tabControl2.TabPages.RemoveByKey("tabPage6");
                        mainform.tabControl2.TabPages.RemoveByKey("tabPage3");
                        mainform.tabControl2.TabPages.RemoveByKey("tabPage8");
                        mainform.tabControl2.TabPages.RemoveByKey("tabPage18");
                        mainform.tabControl2.TabPages.RemoveByKey("tabPage5");
                        mainform.Show();
                        break;
                    case "Шеф-повар":
                        this.Hide();
                        mainform.tabControl2.TabPages.RemoveByKey("tabPage6");
                        mainform.tabControl2.TabPages.RemoveByKey("tabPage3");
                        mainform.tabControl2.TabPages.RemoveByKey("tabPage4");
                        mainform.tabControl2.TabPages.RemoveByKey("tabPage8");
                        mainform.tabControl2.TabPages.RemoveByKey("tabPage18");
                        mainform.Show();
                        break;
                    case "Менеджер бронирования":
                        this.Hide();
                        mainform.tabControl2.TabPages.RemoveByKey("tabPage6");
                        mainform.tabControl2.TabPages.RemoveByKey("tabPage3");
                        mainform.tabControl2.TabPages.RemoveByKey("tabPage4");
                        mainform.tabControl2.TabPages.RemoveByKey("tabPage8");
                        mainform.tabControl2.TabPages.RemoveByKey("tabPage5");
                        mainform.Show();
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
