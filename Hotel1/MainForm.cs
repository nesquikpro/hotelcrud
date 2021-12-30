using iTextSharp.text;
using iTextSharp.text.pdf;
using Maroquio;
using System;
using System.IO;
using System.Security.Cryptography;
using System.Text;
using System.Windows.Forms;

namespace Hotel1
{
    public partial class MainForm : Form
    {
        AllModel<Role> roles = new AllModel<Role>("Roles");
        AllModel<User> users = new AllModel<User>("Users");
        AllModel<Food> foods = new AllModel<Food>("Foods");
        AllModel<OrderingFood> orderingfoods = new AllModel<OrderingFood>("OrderingFoods");
        AllModel<Client> clients = new AllModel<Client>("Clients");
        AllModel<Post> posts = new AllModel<Post>("Posts");
        AllModel<Departament> departaments = new AllModel<Departament>("Departaments");
        AllModel<Employee> employees = new AllModel<Employee>("Employees");
        AllModel<Accounting> accountings = new AllModel<Accounting>("Accountings");
        AllModel<TypeOfService> typeofservice = new AllModel<TypeOfService>("TypeOfServices");
        AllModel<BillForService> billofservice = new AllModel<BillForService>("BillForServices");
        AllModel<RoomNumber> roomnumbers = new AllModel<RoomNumber>("RoomNumbers");
        AllModel<OrderingRoom> orderingrooms = new AllModel<OrderingRoom>("OrderingRooms");
        public MainForm()
        {
            InitializeComponent();
        }

        private async void addRole_Click(object sender, EventArgs e)
        {
            string str = textRole.Text;

            if (String.IsNullOrWhiteSpace(str))
                MessageBox.Show(
        "Заполните поле",
        "Сообщение",
        MessageBoxButtons.OK,
        MessageBoxIcon.Error,
        MessageBoxDefaultButton.Button1,
        MessageBoxOptions.DefaultDesktopOnly);
            else
            {
                Role role = new Role();
                role.Name = textRole.Text;

                await role.Add();
                roleTableUpdate();
                textRole.Text = "";
            }
            comboRole.DataSource = roles.Objs;
            comboRole.DisplayMember = "Name";
            comboRole.ValueMember = "Id";
        }

        /// Загрузка формы
        private void RoleForm_Load(object sender, EventArgs e)
        {
            if (tabControl2.TabPages.ContainsKey("tabPage7"))
            {
                accountingTableUpdate();
            }
            else
                if(tabControl2.TabPages.ContainsKey("tabPage3") ||
                tabControl2.TabPages.ContainsKey("tabPage4") ||
                tabControl2.TabPages.ContainsKey("tabPage5") ||
                tabControl2.TabPages.ContainsKey("tabPage6") ||
                tabControl2.TabPages.ContainsKey("tabPage8") ||
                tabControl2.TabPages.ContainsKey("tabPage18"))
            {
                roleTableUpdate();
                userTableUpdate();
                foodTableUpdate();
                clientTableUpdate();
                orderingfoodTableUpdate();
                postTableUpdate();
                departamentTableUpdate();
                employeeTableUpdate();
                typeofserviceTableUpdate();
                chetauslugiTableUpdate();
                nomeraTableUpdate();
                broniravanieTableUpdate();
            }             

            dataUsers.DataSource = users.Objs;
            comboRole.DataSource = roles.Objs;
            comboRole.DisplayMember = "Name";
            comboRole.ValueMember = "Id";

            dataZakaz.DataSource = orderingfoods.Objs;
            comboMenu.DataSource = foods.Objs;
            comboMenu.DisplayMember = "Name";
            comboMenu.ValueMember = "Id";

            comboClient.DataSource = clients.Objs;
            comboClient.DisplayMember = "Surname";
            comboClient.ValueMember = "Id";

            comboMenu.SelectedItem = 0;
            comboClient.SelectedItem = 0;
            comboClientCheta.SelectedItem = 0;
            comboUsluga.SelectedItem = 0;
            comboBuhEmployee.SelectedItem = 0;
            floorNomer.SelectedIndex = 0;

            comboNumberBron.SelectedItem = 0;
            comboClientBron.SelectedItem = 0;
            comboEmployeeBron.SelectedItem = 0;
        }

        private async void deleteRole_Click(object sender, EventArgs e)
        {
            await (dataRole.SelectedRows[0].DataBoundItem as Role).Delete();
            roleTableUpdate();
            comboRole.DataSource = roles.Objs;
            comboRole.DisplayMember = "Name";
            comboRole.ValueMember = "Id";
        }

        private async void editRole_Click(object sender, EventArgs e)
        {
            string str = textRole.Text;

            if (String.IsNullOrWhiteSpace(str))
                MessageBox.Show(
        "Заполните поле",
        "Сообщение",
        MessageBoxButtons.OK,
        MessageBoxIcon.Error,
        MessageBoxDefaultButton.Button1,
        MessageBoxOptions.DefaultDesktopOnly);
            else
            {
                if (dataRole.SelectedRows[0].DataBoundItem is Role role)
                {
                    role.Name = textRole.Text;
                    await role.Update();
                }
                roleTableUpdate();
            }
            comboRole.DataSource = roles.Objs;
            comboRole.DisplayMember = "Name";
            comboRole.ValueMember = "Id";
        }

        private void roleTableUpdate()
        {
            dataRole.DataSource = roles.Objs;
            dataRole.Columns[0].HeaderText = "Номер роли";
            dataRole.Columns[1].HeaderText = "Название";
            dataRole.Columns[2].Visible = false;

            dataUsers.DataSource = users.Objs;
            comboRole.DataSource = roles.Objs;
            comboRole.DisplayMember = "Name";
            comboRole.ValueMember = "Id";
        }

        private void dataRole_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            textRole.Text = dataRole.SelectedRows[0].Cells[1].Value.ToString();
        }

        private async void deleteToolStripMenuItem_Click(object sender, EventArgs e)
        {
            await (dataRole.SelectedRows[0].DataBoundItem as Role).Delete();
            roleTableUpdate();
            comboRole.DataSource = roles.Objs;
            comboRole.DisplayMember = "Name";
            comboRole.ValueMember = "Id";
        }

        private void выходToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MenuForm menu = new MenuForm();
            this.Hide();
            menu.ShowDialog();
            this.Show();
        }

        private void userTableUpdate()
        {
            dataUsers.DataSource = users.Objs;
            dataUsers.Columns[0].HeaderText = "Номер пользователя";
            dataUsers.Columns[1].HeaderText = "Логин";
            dataUsers.Columns[2].HeaderText = "Пароль";
            dataUsers.Columns[3].HeaderText = "Номер роли";
            dataUsers.Columns[4].Visible = false;
        }
        public string GetHash(string input)
        {
            var md5 = MD5.Create();
            var hash = md5.ComputeHash(Encoding.UTF8.GetBytes(input));

            return Convert.ToBase64String(hash);
        }

        private async void addUser_Click(object sender, EventArgs e)
        {
            string pass = textPassword.Text;
            string str = textLogin.Text;

            if (String.IsNullOrWhiteSpace(str) || String.IsNullOrWhiteSpace(pass))
                MessageBox.Show(
        "Заполните поле",
        "Сообщение",
        MessageBoxButtons.OK,
        MessageBoxIcon.Error,
        MessageBoxDefaultButton.Button1,
        MessageBoxOptions.DefaultDesktopOnly);

            else
            {
                pass = GetHash(textPassword.Text);
                await new User
                {
                    Id = 0,
                    Login = textLogin.Text,
                    Password = pass,
                    RoleId = comboRole.SelectedValue as int?,
                }.Add();
                userTableUpdate();
            }

        }

        private async void editUser_Click(object sender, EventArgs e)
        {
            string pass = textPassword.Text;
            string str = textLogin.Text;

            if (String.IsNullOrWhiteSpace(str) || String.IsNullOrWhiteSpace(pass))
                MessageBox.Show(
        "Заполните поле",
        "Сообщение",
        MessageBoxButtons.OK,
        MessageBoxIcon.Error,
        MessageBoxDefaultButton.Button1,
        MessageBoxOptions.DefaultDesktopOnly);

            if (dataUsers.SelectedRows[0].DataBoundItem is User user)
            {
                pass = GetHash(textPassword.Text);
                user.Login = textLogin.Text;
                user.Password = pass;
                user.RoleId = int.Parse(comboRole.SelectedValue.ToString());
                await user.Update();
            }
            userTableUpdate();
        }

        private async void deleteUser_Click(object sender, EventArgs e)
        {
            await (dataUsers.SelectedRows[0].DataBoundItem as User).Delete();
            userTableUpdate();
        }

        private async void удалитьToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            await (dataUsers.SelectedRows[0].DataBoundItem as User).Delete();
            userTableUpdate();
        }

        private void btnsearchUsers_Click(object sender, EventArgs e)
        {
            //dataUsers.ClearSelection();

            //if (string.IsNullOrWhiteSpace(searchUsers.Text))
            //    return;
        }

        private void foodTableUpdate()
        {
            dataMenu.DataSource = foods.Objs;
            dataMenu.Columns[0].Visible = false;
            dataMenu.Columns[1].HeaderText = "Название";
            dataMenu.Columns[2].HeaderText = "Цена";

            dataZakaz.DataSource = orderingfoods.Objs;
            comboMenu.DataSource = foods.Objs;
            comboMenu.DisplayMember = "Name";
            comboMenu.ValueMember = "Id";


            comboClient.DataSource = clients.Objs;
            comboClient.DisplayMember = "Surname";
            comboClient.ValueMember = "Id";
        }

        private async void addMenu_Click_1(object sender, EventArgs e)
        {
            string str = nameMenu.Text;
            string str2 = priceMenu.Text;

            if (String.IsNullOrWhiteSpace(str) || String.IsNullOrWhiteSpace(str2))
                MessageBox.Show(
        "Заполните поле",
        "Сообщение",
        MessageBoxButtons.OK,
        MessageBoxIcon.Error,
        MessageBoxDefaultButton.Button1,
        MessageBoxOptions.DefaultDesktopOnly);
            else
            {
                await new Food
                {
                    Id = 0,
                    Name = nameMenu.Text,
                    Price = int.Parse(priceMenu.Text),
                }.Add();
                foodTableUpdate();
            }
            comboMenu.DataSource = foods.Objs;
            comboMenu.DisplayMember = "Name";
            comboMenu.ValueMember = "Id";
        }

        private async void editMenu_Click(object sender, EventArgs e)
        {
            string str = nameMenu.Text;
            string str2 = priceMenu.Text;

            if (String.IsNullOrWhiteSpace(str) || String.IsNullOrWhiteSpace(str2))
                MessageBox.Show(
        "Заполните поле",
        "Сообщение",
        MessageBoxButtons.OK,
        MessageBoxIcon.Error,
        MessageBoxDefaultButton.Button1,
        MessageBoxOptions.DefaultDesktopOnly);
            else
            {
                if (dataMenu.SelectedRows[0].DataBoundItem is Food food)
                {
                    food.Name = nameMenu.Text;
                    food.Price = int.Parse(priceMenu.Text);

                    await food.Update();
                }
                foodTableUpdate();
            }
            comboMenu.DataSource = foods.Objs;
            comboMenu.DisplayMember = "Name";
            comboMenu.ValueMember = "Id";
        }

        private async void deleteMenu_Click(object sender, EventArgs e)
        {
            await (dataMenu.SelectedRows[0].DataBoundItem as Food).Delete();
            foodTableUpdate();
            comboMenu.DataSource = foods.Objs;
            comboMenu.DisplayMember = "Name";
            comboMenu.ValueMember = "Id";
        }

        private async void удалитьToolStripMenuItem2_Click(object sender, EventArgs e)
        {
            await (dataMenu.SelectedRows[0].DataBoundItem as Food).Delete();
            foodTableUpdate();
            comboMenu.DataSource = foods.Objs;
            comboMenu.DisplayMember = "Name";
            comboMenu.ValueMember = "Id";
        }

        private void dataUsers_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            textLogin.Text = dataUsers.SelectedRows[0].Cells[1].Value.ToString();
            textPassword.Text = dataUsers.SelectedRows[0].Cells[2].Value.ToString();
        }

        private void dataMenu_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            nameMenu.Text = dataMenu.SelectedRows[0].Cells[1].Value.ToString();
            priceMenu.Text = dataMenu.SelectedRows[0].Cells[2].Value.ToString();
        }

        private void dataZakaz_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            nameMenu.Text = dataMenu.SelectedRows[0].Cells[1].Value.ToString();
            priceMenu.Text = dataMenu.SelectedRows[0].Cells[2].Value.ToString();
        }

        private void orderingfoodTableUpdate()
        {
            dataZakaz.DataSource = orderingfoods.Objs;
            dataZakaz.Columns[0].Visible = false;
            dataZakaz.Columns[1].HeaderText = "Название меню";
            dataZakaz.Columns[2].HeaderText = "Клиент";


        }

        private async void addZakaz_Click(object sender, EventArgs e)
        {
            await new OrderingFood
            {
                Id = 0,
                FoodId = int.Parse(comboMenu.SelectedValue.ToString()),
                ClientId = int.Parse(comboClient.SelectedValue.ToString()),
            }.Add();
            orderingfoodTableUpdate();
        }

        private async void editZakaz_Click(object sender, EventArgs e)
        {
            if (dataZakaz.SelectedRows[0].DataBoundItem is OrderingFood orderingFood)
            {
                orderingFood.FoodId = int.Parse(comboMenu.SelectedValue.ToString());
                orderingFood.ClientId = int.Parse(comboClient.SelectedValue.ToString());

                await orderingFood.Update();
            }
            orderingfoodTableUpdate();
        }

        private async void deleteZakaz_Click(object sender, EventArgs e)
        {
            await (dataZakaz.SelectedRows[0].DataBoundItem as OrderingFood).Delete();
            orderingfoodTableUpdate();
        }

        private async void удалитьToolStripMenuItem3_Click(object sender, EventArgs e)
        {
            await (dataZakaz.SelectedRows[0].DataBoundItem as OrderingFood).Delete();
            orderingfoodTableUpdate();
        }

        private void clientTableUpdate()
        {
            dataClients.DataSource = clients.Objs;

            dataClients.Columns[0].HeaderText = "№";
            dataClients.Columns[1].HeaderText = "Имя";
            dataClients.Columns[2].HeaderText = "Фамилия";
            dataClients.Columns[3].HeaderText = "Отчество";
            dataClients.Columns[4].HeaderText = "Серия паспорта";
            dataClients.Columns[5].HeaderText = "Номер паспорта";
            dataClients.Columns[6].HeaderText = "Телефон";
            dataClients.Columns[7].HeaderText = "Дата Рождения";
            dataClients.Columns[8].HeaderText = "Почта";

            dataClients.Columns[9].Visible = false;
        }


        private async void addClient_Click(object sender, EventArgs e)
        {
            string str = nameClient.Text;
            string str2 = familiaClient.Text;
            string str3 = otchestvoClient.Text;
            string str4 = seriaPassClient.Text;
            string str5 = numberPassClient.Text;
            string str7 = phoneClient.Text;
            string str9 = emailClient.Text;

            if (String.IsNullOrWhiteSpace(str) || String.IsNullOrWhiteSpace(str2)
                || String.IsNullOrWhiteSpace(str3)
                || String.IsNullOrWhiteSpace(str5)
                || String.IsNullOrWhiteSpace(str4)
                || String.IsNullOrWhiteSpace(str5)
                || String.IsNullOrWhiteSpace(str7)
                || String.IsNullOrWhiteSpace(str9))
                MessageBox.Show(
        "Заполните поле",
        "Сообщение",
        MessageBoxButtons.OK,
        MessageBoxIcon.Error,
        MessageBoxDefaultButton.Button1,
        MessageBoxOptions.DefaultDesktopOnly);
            else
            {
                await new Client
                {
                    Id = 0,
                    Name = nameClient.Text,
                    Surname = familiaClient.Text,
                    MiddleName = otchestvoClient.Text,
                    SeriaPass = seriaPassClient.Text,
                    NumberPass = numberPassClient.Text,
                    Phone = phoneClient.Text,
                    DateBirthClient = dateClient.Value,
                    Email = emailClient.Text,

                }.Add();
                clientTableUpdate();
            }
            comboClientCheta.DataSource = clients.Objs;
            comboClientCheta.DisplayMember = "Surname";
            comboClientCheta.ValueMember = "Id";

            comboClientBron.DataSource = clients.Objs;
            comboClientBron.DisplayMember = "Surname";
            comboClientBron.ValueMember = "Id";
        }

        private async void editClient_Click(object sender, EventArgs e)
        {
            string str = nameClient.Text;
            string str2 = familiaClient.Text;
            string str3 = otchestvoClient.Text;
            string str4 = seriaPassClient.Text;
            string str5 = numberPassClient.Text;
            string str7 = phoneClient.Text;
            string str9 = emailClient.Text;

            if (String.IsNullOrWhiteSpace(str) || String.IsNullOrWhiteSpace(str2)
                || String.IsNullOrWhiteSpace(str3)
                || String.IsNullOrWhiteSpace(str5)
                || String.IsNullOrWhiteSpace(str4)
                || String.IsNullOrWhiteSpace(str5)
                || String.IsNullOrWhiteSpace(str7)
                || String.IsNullOrWhiteSpace(str9))
                MessageBox.Show(
        "Заполните поле",
        "Сообщение",
        MessageBoxButtons.OK,
        MessageBoxIcon.Error,
        MessageBoxDefaultButton.Button1,
        MessageBoxOptions.DefaultDesktopOnly);
            else
            {
                if (dataClients.SelectedRows[0].DataBoundItem is Client client)
                {
                    client.Name = nameClient.Text;
                    client.Surname = familiaClient.Text;
                    client.MiddleName = otchestvoClient.Text;
                    client.SeriaPass = seriaPassClient.Text;
                    client.NumberPass = numberPassClient.Text;
                    client.Phone = phoneClient.Text;
                    client.DateBirthClient = dateClient.Value;
                    client.Email = emailClient.Text;

                    await client.Update();
                }
                clientTableUpdate();
            }
            comboClientCheta.DataSource = clients.Objs;
            comboClientCheta.DisplayMember = "Surname";
            comboClientCheta.ValueMember = "Id";

            comboClientBron.DataSource = clients.Objs;
            comboClientBron.DisplayMember = "Surname";
            comboClientBron.ValueMember = "Id";
        }

        private async void deleteClient_Click(object sender, EventArgs e)
        {
            await (dataClients.SelectedRows[0].DataBoundItem as Client).Delete();
            clientTableUpdate();
            comboClientCheta.DataSource = clients.Objs;
            comboClientCheta.DisplayMember = "Surname";
            comboClientCheta.ValueMember = "Id";

            comboClientBron.DataSource = clients.Objs;
            comboClientBron.DisplayMember = "Surname";
            comboClientBron.ValueMember = "Id";
        }

        private void dataClients_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            nameClient.Text = dataClients.SelectedRows[0].Cells[1].Value.ToString();
            familiaClient.Text = dataClients.SelectedRows[0].Cells[2].Value.ToString();
            otchestvoClient.Text = dataClients.SelectedRows[0].Cells[3].Value.ToString();
            seriaPassClient.Text = dataClients.SelectedRows[0].Cells[4].Value.ToString();
            numberPassClient.Text = dataClients.SelectedRows[0].Cells[5].Value.ToString();
            phoneClient.Text = dataClients.SelectedRows[0].Cells[6].Value.ToString();
            dateClient.Text = dataClients.SelectedRows[0].Cells[7].Value.ToString();
            emailClient.Text = dataClients.SelectedRows[0].Cells[8].Value.ToString();
        }

        private async void удалитьToolStripMenuItem5_Click(object sender, EventArgs e)
        {
            await (dataClients.SelectedRows[0].DataBoundItem as Client).Delete();
            clientTableUpdate();
            comboClientCheta.DataSource = clients.Objs;
            comboClientCheta.DisplayMember = "Surname";
            comboClientCheta.ValueMember = "Id";

            comboClientBron.DataSource = clients.Objs;
            comboClientBron.DisplayMember = "Surname";
            comboClientBron.ValueMember = "Id";
        }

        //.............POST - Должности....................................
        private void postTableUpdate()
        {
            dataPost.DataSource = posts.Objs;
            dataPost.Columns[0].HeaderText = "Номер должности";
            dataPost.Columns[1].HeaderText = "Название";
            dataPost.Columns[2].HeaderText = "Оклад";
            dataPost.Columns[3].Visible = false;
        }

        private async void addPost_Click(object sender, EventArgs e)
        {
            string str = namePost.Text;
            string str2 = okladPost.Text;

            if (String.IsNullOrWhiteSpace(str) || String.IsNullOrWhiteSpace(str2))
                MessageBox.Show(
        "Заполните поле",
        "Сообщение",
        MessageBoxButtons.OK,
        MessageBoxIcon.Error,
        MessageBoxDefaultButton.Button1,
        MessageBoxOptions.DefaultDesktopOnly);
            else
            {
                await new Post
                {
                    Id = 0,
                    Name = namePost.Text,
                    Salary = int.Parse(okladPost.Text),

                }.Add();
                postTableUpdate();
            }
            comboPost.DataSource = posts.Objs;
            comboPost.DisplayMember = "Name";
            comboPost.ValueMember = "Id";
        }

        private async void editPost_Click(object sender, EventArgs e)
        {
            string str = namePost.Text;
            string str2 = okladPost.Text;

            if (String.IsNullOrWhiteSpace(str) || String.IsNullOrWhiteSpace(str2))
                MessageBox.Show(
        "Заполните поле",
        "Сообщение",
        MessageBoxButtons.OK,
        MessageBoxIcon.Error,
        MessageBoxDefaultButton.Button1,
        MessageBoxOptions.DefaultDesktopOnly);
            else
            {
                if (dataPost.SelectedRows[0].DataBoundItem is Post post)
                {
                    post.Name = namePost.Text;
                    post.Salary = int.Parse(okladPost.Text);

                    await post.Update();
                }
                postTableUpdate();
            }
            comboPost.DataSource = posts.Objs;
            comboPost.DisplayMember = "Name";
            comboPost.ValueMember = "Id";
        }

        private async void deletePost_Click(object sender, EventArgs e)
        {
            await (dataPost.SelectedRows[0].DataBoundItem as Post).Delete();
            postTableUpdate();
            comboPost.DataSource = posts.Objs;
            comboPost.DisplayMember = "Name";
            comboPost.ValueMember = "Id";
        }

        private async void удалитьToolStripMenuItem4_Click(object sender, EventArgs e)
        {
            await (dataPost.SelectedRows[0].DataBoundItem as Post).Delete();
            postTableUpdate();
            comboPost.DataSource = posts.Objs;
            comboPost.DisplayMember = "Name";
            comboPost.ValueMember = "Id";
        }

        private void dataPost_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            namePost.Text = dataPost.SelectedRows[0].Cells[1].Value.ToString();
            okladPost.Text = dataPost.SelectedRows[0].Cells[2].Value.ToString();
        }

        //.............Отделы....................................
        private void departamentTableUpdate()
        {
            dataOtdel.DataSource = departaments.Objs;
            dataOtdel.Columns[0].HeaderText = "Номер отдела";
            dataOtdel.Columns[1].HeaderText = "Название";
            dataOtdel.Columns[2].HeaderText = "Номер телефона";
            dataOtdel.Columns[3].Visible = false;
        }
        private async void addOtdel_Click(object sender, EventArgs e)
        {
            string pass = nameOtdel.Text;
            string str = phoneOtdel.Text;

            if (String.IsNullOrWhiteSpace(str) || String.IsNullOrWhiteSpace(pass))
                MessageBox.Show(
        "Заполните поле",
        "Сообщение",
        MessageBoxButtons.OK,
        MessageBoxIcon.Error,
        MessageBoxDefaultButton.Button1,
        MessageBoxOptions.DefaultDesktopOnly);
            else
            {
                await new Departament
                {
                    Id = 0,
                    Name = nameOtdel.Text,
                    DepartamentPhone = phoneOtdel.Text,

                }.Add();
                departamentTableUpdate();
            }
            comboOtdel.DataSource = departaments.Objs;
            comboOtdel.DisplayMember = "Name";
            comboOtdel.ValueMember = "Id";
        }

        private async void editOtdel_Click(object sender, EventArgs e)
        {
            string pass = nameOtdel.Text;
            string str = phoneOtdel.Text;

            if (String.IsNullOrWhiteSpace(str) || String.IsNullOrWhiteSpace(pass))
                MessageBox.Show(
        "Заполните поле",
        "Сообщение",
        MessageBoxButtons.OK,
        MessageBoxIcon.Error,
        MessageBoxDefaultButton.Button1,
        MessageBoxOptions.DefaultDesktopOnly);
            else
            {
                if (dataOtdel.SelectedRows[0].DataBoundItem is Departament departament)
                {
                    departament.Name = nameOtdel.Text;
                    departament.DepartamentPhone = phoneOtdel.Text;

                    await departament.Update();
                }
                departamentTableUpdate();
            }
            comboOtdel.DataSource = departaments.Objs;
            comboOtdel.DisplayMember = "Name";
            comboOtdel.ValueMember = "Id";
        }

        private async void deleteOtdel_Click(object sender, EventArgs e)
        {
            await (dataOtdel.SelectedRows[0].DataBoundItem as Departament).Delete();
            departamentTableUpdate();
            comboOtdel.DataSource = departaments.Objs;
            comboOtdel.DisplayMember = "Name";
            comboOtdel.ValueMember = "Id";
        }

        private async void удалитьToolStripMenuItem6_Click(object sender, EventArgs e)
        {
            await (dataOtdel.SelectedRows[0].DataBoundItem as Departament).Delete();
            departamentTableUpdate();
            comboOtdel.DataSource = departaments.Objs;
            comboOtdel.DisplayMember = "Name";
            comboOtdel.ValueMember = "Id";
        }

        private void dataOtdel_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            nameOtdel.Text = dataOtdel.SelectedRows[0].Cells[1].Value.ToString();
            phoneOtdel.Text = dataOtdel.SelectedRows[0].Cells[3].Value.ToString();
        }


        //.............Сотружники....................................
        private void employeeTableUpdate()
        {
            dataEmployee.DataSource = employees.Objs;

            comboEmpl.DataSource = users.Objs;
            comboEmpl.DisplayMember = "Login";
            comboEmpl.ValueMember = "Id";

            comboPost.DataSource = posts.Objs;
            comboPost.DisplayMember = "Name";
            comboPost.ValueMember = "Id";

            comboOtdel.DataSource = departaments.Objs;
            comboOtdel.DisplayMember = "Name";
            comboOtdel.ValueMember = "Id";

            dataEmployee.Columns[0].HeaderText = "№";
            dataEmployee.Columns[1].HeaderText = "Имя";
            dataEmployee.Columns[2].HeaderText = "Фамилия";
            dataEmployee.Columns[3].HeaderText = "Отчество";
            dataEmployee.Columns[4].HeaderText = "Серия паспорта";
            dataEmployee.Columns[5].HeaderText = "Номер паспорта";
            dataEmployee.Columns[6].HeaderText = "Телефон";
            dataEmployee.Columns[7].HeaderText = "Дата Рождения";
            dataEmployee.Columns[8].HeaderText = "Почта";
            dataEmployee.Columns[9].HeaderText = "Должность";
            dataEmployee.Columns[10].HeaderText = "Отдел";
            dataEmployee.Columns[11].HeaderText = "Пользователь";
            dataEmployee.Columns[12].Visible = false;
        }

        private async void addEmployee_Click(object sender, EventArgs e)
        {
            string str = textName.Text;
            string str2 = familiaText.Text;
            string str3 = otchestvoText.Text;
            string str4 = seriaPass.Text;
            string str5 = nmberText.Text;
            string str7 = phoneText.Text;
            string str9 = emailText.Text;

            if (String.IsNullOrWhiteSpace(str) || String.IsNullOrWhiteSpace(str2)
                || String.IsNullOrWhiteSpace(str3)
                || String.IsNullOrWhiteSpace(str5)
                || String.IsNullOrWhiteSpace(str4)
                || String.IsNullOrWhiteSpace(str5)
                || String.IsNullOrWhiteSpace(str7)
                || String.IsNullOrWhiteSpace(str9))
                MessageBox.Show(
        "Заполните поле",
        "Сообщение",
        MessageBoxButtons.OK,
        MessageBoxIcon.Error,
        MessageBoxDefaultButton.Button1,
        MessageBoxOptions.DefaultDesktopOnly);

            else
            {
                await new Employee
                {
                    Id = 0,
                    Name = textName.Text,
                    Surname = familiaText.Text,
                    MiddleName = otchestvoText.Text,
                    SeriaPasport = seriaPass.Text,
                    NumberPasport = nmberText.Text,
                    EmployeePhone = phoneText.Text,
                    DateBirthEmployee = dateEmpl.Value,
                    EmployeeEmail = emailText.Text,
                    PostId = int.Parse(comboPost.SelectedValue.ToString()),
                    UserId = int.Parse(comboEmpl.SelectedValue.ToString()),
                    DepartamentId = int.Parse(comboOtdel.SelectedValue.ToString()),
                }.Add();
                employeeTableUpdate();
            }
            comboBuhEmployee.DataSource = employees.Objs;
            comboBuhEmployee.DisplayMember = "Surname";
            comboBuhEmployee.ValueMember = "Id";

            comboEmployeeBron.DataSource = employees.Objs;
            comboEmployeeBron.DisplayMember = "Surname";
            comboEmployeeBron.ValueMember = "Id";
        }

        private async void editEmployee_Click(object sender, EventArgs e)
        {
            string str = textName.Text;
            string str2 = familiaText.Text;
            string str3 = otchestvoText.Text;
            string str4 = seriaPass.Text;
            string str5 = nmberText.Text;
            string str7 = phoneText.Text;
            string str9 = emailText.Text;

            if (String.IsNullOrWhiteSpace(str) || String.IsNullOrWhiteSpace(str2)
                || String.IsNullOrWhiteSpace(str3)
                || String.IsNullOrWhiteSpace(str5)
                || String.IsNullOrWhiteSpace(str4)
                || String.IsNullOrWhiteSpace(str5)
                || String.IsNullOrWhiteSpace(str7)
                || String.IsNullOrWhiteSpace(str9))
                MessageBox.Show(
        "Заполните поле",
        "Сообщение",
        MessageBoxButtons.OK,
        MessageBoxIcon.Error,
        MessageBoxDefaultButton.Button1,
        MessageBoxOptions.DefaultDesktopOnly);
            else
            {
                if (dataEmployee.SelectedRows[0].DataBoundItem is Employee employee)
                {
                    employee.Name = textName.Text;
                    employee.Surname = familiaText.Text;
                    employee.MiddleName = otchestvoText.Text;
                    employee.SeriaPasport = seriaPass.Text;
                    employee.NumberPasport = nmberText.Text;
                    employee.EmployeePhone = phoneText.Text;
                    employee.DateBirthEmployee = dateEmpl.Value;
                    employee.EmployeeEmail = emailText.Text;
                    employee.PostId = int.Parse(comboPost.SelectedValue.ToString());
                    employee.UserId = int.Parse(comboEmpl.SelectedValue.ToString());
                    employee.DepartamentId = int.Parse(comboOtdel.SelectedValue.ToString());

                    await employee.Update();
                }
                employeeTableUpdate();
            }
            comboBuhEmployee.DataSource = employees.Objs;
            comboBuhEmployee.DisplayMember = "Surname";
            comboBuhEmployee.ValueMember = "Id";

            comboEmployeeBron.DataSource = employees.Objs;
            comboEmployeeBron.DisplayMember = "Surname";
            comboEmployeeBron.ValueMember = "Id";
        }

        private async void deleteEmployee_Click(object sender, EventArgs e)
        {
            await (dataEmployee.SelectedRows[0].DataBoundItem as Employee).Delete();
            employeeTableUpdate();
            comboBuhEmployee.DataSource = employees.Objs;
            comboBuhEmployee.DisplayMember = "Surname";
            comboBuhEmployee.ValueMember = "Id";

            comboEmployeeBron.DataSource = employees.Objs;
            comboEmployeeBron.DisplayMember = "Surname";
            comboEmployeeBron.ValueMember = "Id";
        }

        private async void удалитьToolStripMenuItem7_Click(object sender, EventArgs e)
        {
            await (dataEmployee.SelectedRows[0].DataBoundItem as Employee).Delete();
            employeeTableUpdate();
            comboBuhEmployee.DataSource = employees.Objs;
            comboBuhEmployee.DisplayMember = "Surname";
            comboBuhEmployee.ValueMember = "Id";

            comboEmployeeBron.DataSource = employees.Objs;
            comboEmployeeBron.DisplayMember = "Surname";
            comboEmployeeBron.ValueMember = "Id";
        }

        private void dataEmployee_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            textName.Text = dataEmployee.SelectedRows[0].Cells[1].Value.ToString();
            familiaText.Text = dataEmployee.SelectedRows[0].Cells[2].Value.ToString();
            otchestvoText.Text = dataEmployee.SelectedRows[0].Cells[3].Value.ToString();
            seriaPass.Text = dataEmployee.SelectedRows[0].Cells[4].Value.ToString();
            nmberText.Text = dataEmployee.SelectedRows[0].Cells[5].Value.ToString();
            phoneText.Text = dataEmployee.SelectedRows[0].Cells[6].Value.ToString();
            dateEmpl.Text = dataEmployee.SelectedRows[0].Cells[7].Value.ToString();
            emailText.Text = dataEmployee.SelectedRows[0].Cells[8].Value.ToString();
        }

        //.............Бухгалтерия....................................

        private void accountingTableUpdate()
        {
            dataBuh.DataSource = accountings.Objs;
            dataBuh.Columns[0].Visible = false;
            dataBuh.Columns[1].HeaderText = "Сумма";
            dataBuh.Columns[2].HeaderText = "Дата выдачи";
            dataBuh.Columns[3].HeaderText = "Сотрудник";
            dataBuh.Columns[4].Visible = false;

            comboBuhEmployee.DataSource = employees.Objs;
            comboBuhEmployee.DisplayMember = "Surname";
            comboBuhEmployee.ValueMember = "Id";
        }

        private async void addBuh_Click(object sender, EventArgs e)
        {
            string str = priceBuh.Text;

            if (String.IsNullOrWhiteSpace(str))
                MessageBox.Show(
        "Заполните поле",
        "Сообщение",
        MessageBoxButtons.OK,
        MessageBoxIcon.Error,
        MessageBoxDefaultButton.Button1,
        MessageBoxOptions.DefaultDesktopOnly);
            else
            {
                await new Accounting
                {
                    Id = 0,
                    Amount = int.Parse(priceBuh.Text),
                    DateIssue = dateBuh.Value,
                    EmployeeId = int.Parse(comboBuhEmployee.SelectedValue.ToString()),
                }.Add();
                accountingTableUpdate();
            }

        }

        private async void editBuh_Click(object sender, EventArgs e)
        {
            string str = priceBuh.Text;

            if (String.IsNullOrWhiteSpace(str))
                MessageBox.Show(
        "Заполните поле",
        "Сообщение",
        MessageBoxButtons.OK,
        MessageBoxIcon.Error,
        MessageBoxDefaultButton.Button1,
        MessageBoxOptions.DefaultDesktopOnly);
            else
            {
                if (dataBuh.SelectedRows[0].DataBoundItem is Accounting accounting)
                {
                    accounting.Amount = int.Parse(priceBuh.Text);
                    accounting.DateIssue = dateBuh.Value;
                    accounting.EmployeeId = int.Parse(comboBuhEmployee.SelectedValue.ToString());
                    await accounting.Update();
                }
                accountingTableUpdate();
            }
        }

        private async void deleteBuh_Click(object sender, EventArgs e)
        {
            await (dataBuh.SelectedRows[0].DataBoundItem as Accounting).Delete();
            accountingTableUpdate();
        }

        private async void удалитьToolStripMenuItem8_Click(object sender, EventArgs e)
        {
            await (dataBuh.SelectedRows[0].DataBoundItem as Accounting).Delete();
            accountingTableUpdate();
        }

        private void dataBuh_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            priceBuh.Text = dataBuh.SelectedRows[0].Cells[1].Value.ToString();
        }

        private void dataBuh_CellClick_1(object sender, DataGridViewCellEventArgs e)
        {
            priceBuh.Text = dataBuh.SelectedRows[0].Cells[1].Value.ToString();
        }

        //.............Услуги....................................

        private void typeofserviceTableUpdate()
        {
            dataUsluga.DataSource = typeofservice.Objs;
            dataUsluga.Columns[0].Visible = false;
            dataUsluga.Columns[1].HeaderText = "Название";
            dataUsluga.Columns[2].HeaderText = "Цена";
            dataUsluga.Columns[3].Visible = false;
        }

        private async void addUsluga_Click(object sender, EventArgs e)
        {
            string str = nameUsluga.Text;
            string str2 = priceUsluga.Text;

            if (String.IsNullOrWhiteSpace(str) || String.IsNullOrWhiteSpace(str2))
                MessageBox.Show(
        "Заполните поле",
        "Сообщение",
        MessageBoxButtons.OK,
        MessageBoxIcon.Error,
        MessageBoxDefaultButton.Button1,
        MessageBoxOptions.DefaultDesktopOnly);
            else
            {
                await new TypeOfService
                {
                    Id = 0,
                    Name = nameUsluga.Text,
                    Price = int.Parse(priceUsluga.Text),

                }.Add();
                typeofserviceTableUpdate();
            }
            comboUsluga.DataSource = typeofservice.Objs;
            comboUsluga.DisplayMember = "Name";
            comboUsluga.ValueMember = "Id";
        }

        private async void editUsluga_Click(object sender, EventArgs e)
        {
            string str = nameUsluga.Text;
            string str2 = priceUsluga.Text;

            if (String.IsNullOrWhiteSpace(str) || String.IsNullOrWhiteSpace(str2))
                MessageBox.Show(
        "Заполните поле",
        "Сообщение",
        MessageBoxButtons.OK,
        MessageBoxIcon.Error,
        MessageBoxDefaultButton.Button1,
        MessageBoxOptions.DefaultDesktopOnly);
            else
            {
                if (dataUsluga.SelectedRows[0].DataBoundItem is TypeOfService typeOfService)
                {
                    typeOfService.Name = nameUsluga.Text;
                    typeOfService.Price = int.Parse(priceUsluga.Text);

                    await typeOfService.Update();
                }
                typeofserviceTableUpdate();
            }
            comboUsluga.DataSource = typeofservice.Objs;
            comboUsluga.DisplayMember = "Name";
            comboUsluga.ValueMember = "Id";
        }

        private async void deleteUsluga_Click(object sender, EventArgs e)
        {
            await (dataUsluga.SelectedRows[0].DataBoundItem as TypeOfService).Delete();
            typeofserviceTableUpdate();
            comboUsluga.DataSource = typeofservice.Objs;
            comboUsluga.DisplayMember = "Name";
            comboUsluga.ValueMember = "Id";
        }

        private async void удалитьToolStripMenuItem9_Click(object sender, EventArgs e)
        {
            await (dataUsluga.SelectedRows[0].DataBoundItem as TypeOfService).Delete();
            typeofserviceTableUpdate();
            comboUsluga.DataSource = typeofservice.Objs;
            comboUsluga.DisplayMember = "Name";
            comboUsluga.ValueMember = "Id";
        }

        private void dataUsluga_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            nameUsluga.Text = dataUsluga.SelectedRows[0].Cells[1].Value.ToString();
            priceUsluga.Text = dataUsluga.SelectedRows[0].Cells[2].Value.ToString();
        }

        /////////Счета за услуги//////////////////////////////////////////////////////////////
        private void chetauslugiTableUpdate()
        {
            dataChetaUsluga.DataSource = billofservice.Objs;

            comboClientCheta.DataSource = clients.Objs;
            comboClientCheta.DisplayMember = "Surname";
            comboClientCheta.ValueMember = "Id";

            comboUsluga.DataSource = typeofservice.Objs;
            comboUsluga.DisplayMember = "Name";
            comboUsluga.ValueMember = "Id";
            dataChetaUsluga.Columns[0].Visible = false;
            dataChetaUsluga.Columns[1].HeaderText = "Дата";
            dataChetaUsluga.Columns[2].HeaderText = "Клиент";
            dataChetaUsluga.Columns[3].Visible = false;
        }
        private async void addChetaUsluga_Click(object sender, EventArgs e)
        {
            await new BillForService
            {
                Id = 0,
                InvoiceDate = textDataUsluga.Value,
                ClientId = int.Parse(comboClientCheta.SelectedValue.ToString()),
                TypeOfServicesId = int.Parse(comboUsluga.SelectedValue.ToString()),

            }.Add();
            chetauslugiTableUpdate();
        }

        private async void удалитьToolStripMenuItem10_Click(object sender, EventArgs e)
        {
            await (dataChetaUsluga.SelectedRows[0].DataBoundItem as BillForService).Delete();
            chetauslugiTableUpdate();
        }

        private async void deleteChetaUsluga_Click(object sender, EventArgs e)
        {
            await (dataChetaUsluga.SelectedRows[0].DataBoundItem as BillForService).Delete();
            chetauslugiTableUpdate();
        }

        private async void editChetaUsluga_Click(object sender, EventArgs e)
        {
            if (dataChetaUsluga.SelectedRows[0].DataBoundItem is BillForService billForService)
            {
                billForService.InvoiceDate = textDataUsluga.Value;
                billForService.ClientId = int.Parse(comboClientCheta.SelectedValue.ToString());
                billForService.TypeOfServicesId = int.Parse(comboUsluga.SelectedValue.ToString());

                await billForService.Update();
            }
            chetauslugiTableUpdate();
        }

        //Номера//////////////////////////////////////////////////////////////////////
        private void nomeraTableUpdate()
        {
            dataNomer.DataSource = roomnumbers.Objs;

            dataNomer.Columns[0].Visible = false;
            dataNomer.Columns[1].HeaderText = "Номер комнаты";
            dataNomer.Columns[2].HeaderText = "Этаж";
            dataNomer.Columns[3].HeaderText = "Цена";
            dataNomer.Columns[4].Visible = false;
        }

        private async void addNomer_Click(object sender, EventArgs e)
        {
            string str = textNomer.Text;
            string str2 = priceNomer.Text;

            if (String.IsNullOrWhiteSpace(str)
                || String.IsNullOrWhiteSpace(str2))
                MessageBox.Show(
        "Заполните поле",
        "Сообщение",
        MessageBoxButtons.OK,
        MessageBoxIcon.Error,
        MessageBoxDefaultButton.Button1,
        MessageBoxOptions.DefaultDesktopOnly);
            else
            {
                await new RoomNumber
                {
                    Id = 0,
                    RoomNumber1 = int.Parse(textNomer.Text),
                    RoomFloor = int.Parse(floorNomer.Text),
                    Price = int.Parse(priceNomer.Text),

                }.Add();
                nomeraTableUpdate();
            }

            comboNumberBron.DataSource = roomnumbers.Objs;
            comboNumberBron.DisplayMember = "RoomNumber1";
            comboNumberBron.ValueMember = "Id";
        }

        private async void deleteNomer_Click(object sender, EventArgs e)
        {
            await (dataNomer.SelectedRows[0].DataBoundItem as RoomNumber).Delete();
            nomeraTableUpdate();

            comboNumberBron.DataSource = roomnumbers.Objs;
            comboNumberBron.DisplayMember = "RoomNumber1";
            comboNumberBron.ValueMember = "Id";
        }

        private async void editNomer_Click(object sender, EventArgs e)
        {
            string str = textNomer.Text;
            string str2 = priceNomer.Text;

            if (String.IsNullOrWhiteSpace(str)
                || String.IsNullOrWhiteSpace(str2))
                MessageBox.Show(
        "Заполните поле",
        "Сообщение",
        MessageBoxButtons.OK,
        MessageBoxIcon.Error,
        MessageBoxDefaultButton.Button1,
        MessageBoxOptions.DefaultDesktopOnly);
            else
            {
                if (dataNomer.SelectedRows[0].DataBoundItem is RoomNumber roomNumber)
                {
                    roomNumber.RoomNumber1 = int.Parse(textNomer.Text);
                    roomNumber.RoomFloor = int.Parse(floorNomer.Text);
                    roomNumber.Price = int.Parse(priceNomer.Text);

                    await roomNumber.Update();
                }
                nomeraTableUpdate();
            }

            comboNumberBron.DataSource = roomnumbers.Objs;
            comboNumberBron.DisplayMember = "RoomNumber1";
            comboNumberBron.ValueMember = "Id";
        }

        private async void удалитьToolStripMenuItem12_Click(object sender, EventArgs e)
        {
            await (dataNomer.SelectedRows[0].DataBoundItem as RoomNumber).Delete();
            nomeraTableUpdate();

            comboNumberBron.DataSource = roomnumbers.Objs;
            comboNumberBron.DisplayMember = "RoomNumber1";
            comboNumberBron.ValueMember = "Id";
        }

        private async void dataNomer_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            textNomer.Text = dataNomer.SelectedRows[0].Cells[1].Value.ToString();
        }

        //Бронирование//////////////////////////////////////////////////////
        private void broniravanieTableUpdate()
        {
            dataBron.DataSource = orderingrooms.Objs;

            comboNumberBron.DataSource = roomnumbers.Objs;
            comboNumberBron.DisplayMember = "RoomNumber1";
            comboNumberBron.ValueMember = "Id";

            comboEmployeeBron.DataSource = employees.Objs;
            comboEmployeeBron.DisplayMember = "Surname";
            comboEmployeeBron.ValueMember = "Id";

            comboClientBron.DataSource = clients.Objs;
            comboClientBron.DisplayMember = "Surname";
            comboClientBron.ValueMember = "Id";

            dataBron.Columns[0].Visible = false;
            dataBron.Columns[1].HeaderText = "Дата заезда";
            dataBron.Columns[2].HeaderText = "Дата выезда";
            dataBron.Columns[3].HeaderText = "Клиент";
            dataBron.Columns[4].HeaderText = "Cотрудник";
            dataBron.Columns[5].HeaderText = "Забронированный номер";
            dataBron.Columns[6].Visible = false;

        }

        private async void addBron_Click(object sender, EventArgs e)
        {
            await new OrderingRoom
            {
                Id = 0,
                ArrivalDate = dataZaezda.Value,
                DepartureDate = dataVezda.Value,
                ClientId = int.Parse(comboClientBron.SelectedValue.ToString()),
                EmployeeId = int.Parse(comboEmployeeBron.SelectedValue.ToString()),
                NumberId = int.Parse(comboNumberBron.SelectedValue.ToString()),

            }.Add();
            broniravanieTableUpdate();
        }

        private async void editBron_Click(object sender, EventArgs e)
        {
            if (dataBron.SelectedRows[0].DataBoundItem is OrderingRoom orderingRoom)
            {
                orderingRoom.ArrivalDate = dataZaezda.Value;
                orderingRoom.DepartureDate = dataVezda.Value;
                orderingRoom.ClientId = int.Parse(comboClientBron.SelectedValue.ToString());
                orderingRoom.EmployeeId = int.Parse(comboEmployeeBron.SelectedValue.ToString());
                orderingRoom.NumberId = int.Parse(comboNumberBron.SelectedValue.ToString());

                await orderingRoom.Update();
            }
            broniravanieTableUpdate();
        }

        private async void deleteBron_Click(object sender, EventArgs e)
        {
            await (dataBron.SelectedRows[0].DataBoundItem as OrderingRoom).Delete();
            broniravanieTableUpdate();
        }

        private async void удалитьToolStripMenuItem13_Click(object sender, EventArgs e)
        {
            await (dataBron.SelectedRows[0].DataBoundItem as OrderingRoom).Delete();
            broniravanieTableUpdate();
        }

        private void MainForm_FormClosed(object sender, FormClosedEventArgs e)
        {
            Application.Exit();
        }

        //////////////////////////ЭКСПОРТ///////////////////////////////////////

        private void eXCELToolStripMenuItem_Click_1(object sender, EventArgs e)
        {
            Microsoft.Office.Interop.Excel.Application ExcelApp = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook ExcelWorkBook;
            Microsoft.Office.Interop.Excel.Worksheet ExcelWorkSheet;
            //Книга.
            ExcelWorkBook = ExcelApp.Workbooks.Add(System.Reflection.Missing.Value);
            //Таблица.
            ExcelWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)ExcelWorkBook.Worksheets.get_Item(1);

            for (int i = 0; i < dataBuh.Rows.Count; i++)
            {
                for (int j = 0; j < dataBuh.ColumnCount; j++)
                {
                    ExcelApp.Cells[i + 1, j + 1] = dataBuh.Rows[i].Cells[j].Value; //Тут ошибку выдает
                }
            }
            //Вызываем нашу созданную эксельку.
            ExcelApp.Visible = true;
            ExcelApp.UserControl = true;
        }

        private void pDFToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (dataBuh.Rows.Count > 0)
            {
                SaveFileDialog sfd = new SaveFileDialog();
                sfd.Filter = "PDF (*.pdf)|*.pdf";
                sfd.FileName = "PFDFile.pdf";
                bool fileError = false;
                if (sfd.ShowDialog() == DialogResult.OK)
                {
                    if (File.Exists(sfd.FileName))
                    {
                        try
                        {
                            File.Delete(sfd.FileName);
                        }
                        catch (IOException ex)
                        {
                            fileError = true;
                            MessageBox.Show("Записать данные на диск не удалось." + ex.Message);
                        }
                    }
                    if (!fileError)
                    {
                        try
                        {
                            PdfPTable pdfTable = new PdfPTable(dataBuh.Columns.Count);
                            pdfTable.DefaultCell.Padding = 3;
                            pdfTable.WidthPercentage = 100;
                            pdfTable.HorizontalAlignment = Element.ALIGN_LEFT;

                            foreach (DataGridViewColumn column in dataBuh.Columns)
                            {
                                PdfPCell cell = new PdfPCell(new Phrase(column.HeaderText));
                                pdfTable.AddCell(cell);
                            }

                            foreach (DataGridViewRow row in dataBuh.Rows)
                            {
                                foreach (DataGridViewCell cell in row.Cells)
                                {
                                    pdfTable.AddCell(cell.Value.ToString());
                                }
                            }

                            using (FileStream stream = new FileStream(sfd.FileName, FileMode.Create))
                            {
                                Document pdfDoc = new Document(PageSize.A4, 10f, 20f, 20f, 10f);
                                PdfWriter.GetInstance(pdfDoc, stream);
                                pdfDoc.Open();
                                pdfDoc.Add(pdfTable);
                                pdfDoc.Close();
                                stream.Close();
                            }

                            MessageBox.Show("Данные Успешно Экспортированы!!!", "Info");
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show("Ошибка :" + ex.Message);
                        }
                    }
                }
            }
        }

        private void wORDToolStripMenuItem_Click(object sender, EventArgs e)
        {
            SaveFileDialog sfd = new SaveFileDialog();

            sfd.Filter = "Word Documents (*.docx)|*.docx";

            sfd.FileName = "export.docx";

            if (sfd.ShowDialog() == DialogResult.OK)
            {
                Export_Data_To_Word(dataBuh, sfd.FileName);
            }
        }

        public void Export_Data_To_Word(DataGridView DGV, string filename)
        {
            if (DGV.Rows.Count != 0)
            {
                int RowCount = DGV.Rows.Count;
                int ColumnCount = DGV.Columns.Count;
                Object[,] DataArray = new object[RowCount + 1, ColumnCount + 1];

                //add rows
                int r = 0;
                for (int c = 0; c <= ColumnCount - 1; c++)
                {
                    for (r = 0; r <= RowCount - 1; r++)
                    {
                        DataArray[r, c] = DGV.Rows[r].Cells[c].Value;
                    } //end row loop
                } //end column loop

                Microsoft.Office.Interop.Word.Document oDoc = new Microsoft.Office.Interop.Word.Document();
                oDoc.Application.Visible = true;

                //page orintation
                oDoc.PageSetup.Orientation = Microsoft.Office.Interop.Word.WdOrientation.wdOrientLandscape;


                dynamic oRange = oDoc.Content.Application.Selection.Range;
                string oTemp = "";
                for (r = 0; r <= RowCount - 1; r++)
                {
                    for (int c = 0; c <= ColumnCount - 1; c++)
                    {
                        oTemp = oTemp + DataArray[r, c] + "\t";

                    }
                }

                //table format
                oRange.Text = oTemp;

                object Separator = Microsoft.Office.Interop.Word.WdTableFieldSeparator.wdSeparateByTabs;
                object ApplyBorders = true;
                object AutoFit = true;
                object AutoFitBehavior = Microsoft.Office.Interop.Word.WdAutoFitBehavior.wdAutoFitContent;

                oRange.ConvertToTable(ref Separator, ref RowCount, ref ColumnCount,
                                      Type.Missing, Type.Missing, ref ApplyBorders,
                                      Type.Missing, Type.Missing, Type.Missing,
                                      Type.Missing, Type.Missing, Type.Missing,
                                      Type.Missing, ref AutoFit, ref AutoFitBehavior, Type.Missing);

                oRange.Select();

                oDoc.Application.Selection.Tables[1].Select();
                oDoc.Application.Selection.Tables[1].Rows.AllowBreakAcrossPages = 0;
                oDoc.Application.Selection.Tables[1].Rows.Alignment = 0;
                oDoc.Application.Selection.Tables[1].Rows[1].Select();
                oDoc.Application.Selection.InsertRowsAbove(1);
                oDoc.Application.Selection.Tables[1].Rows[1].Select();

                //header row style
                oDoc.Application.Selection.Tables[1].Rows[1].Range.Bold = 1;
                oDoc.Application.Selection.Tables[1].Rows[1].Range.Font.Name = "Tahoma";
                oDoc.Application.Selection.Tables[1].Rows[1].Range.Font.Size = 14;

                //add header row manually
                for (int c = 0; c <= ColumnCount - 1; c++)
                {
                    oDoc.Application.Selection.Tables[1].Cell(1, c + 1).Range.Text = DGV.Columns[c].HeaderText;
                }

                //table style 
                //oDoc.Application.Selection.Tables[1].set_Style("Grid Table 4 - Accent 5");
                oDoc.Application.Selection.Tables[1].Rows[1].Select();
                oDoc.Application.Selection.Cells.VerticalAlignment = Microsoft.Office.Interop.Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;

                //header text
                foreach (Microsoft.Office.Interop.Word.Section section in oDoc.Application.ActiveDocument.Sections)
                {
                    Microsoft.Office.Interop.Word.Range headerRange = section.Headers[Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                    headerRange.Fields.Add(headerRange, Microsoft.Office.Interop.Word.WdFieldType.wdFieldPage);
                    headerRange.Text = "";
                    headerRange.Font.Size = 16;
                    headerRange.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter;
                }

                //save the file
                oDoc.SaveAs2(filename);

                //NASSIM LOUCHANI
            }
        }

        private void cVSToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (dataBuh.Rows.Count > 0)
            {
                SaveFileDialog sfd = new SaveFileDialog();
                sfd.Filter = "CSV (*.csv)|*.csv";
                sfd.FileName = "Output.csv";
                bool fileError = false;
                if (sfd.ShowDialog() == DialogResult.OK)
                {
                    if (File.Exists(sfd.FileName))
                    {
                        try
                        {
                            File.Delete(sfd.FileName);
                        }
                        catch (IOException ex)
                        {
                            fileError = true;
                            MessageBox.Show("Записать данные на диск не удалось." + ex.Message);
                        }
                    }
                    if (!fileError)
                    {
                        try
                        {
                            int columnCount = dataBuh.Columns.Count;
                            string columnNames = "";
                            string[] outputCsv = new string[dataBuh.Rows.Count + 1];
                            for (int i = 0; i < columnCount; i++)
                            {
                                columnNames += dataBuh.Columns[i].HeaderText.ToString() + ",";
                            }
                            outputCsv[0] += columnNames;

                            for (int i = 1; (i - 1) < dataBuh.Rows.Count; i++)
                            {
                                for (int j = 0; j < columnCount; j++)
                                {
                                    outputCsv[i] += dataBuh.Rows[i - 1].Cells[j].Value.ToString() + ",";
                                }
                            }

                            File.WriteAllLines(sfd.FileName, outputCsv, Encoding.UTF8);
                            MessageBox.Show("Данные Успешно Экспортированы!!!", "Info");
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show("Ошибка :" + ex.Message);
                        }
                    }
                }
            }
            else
            {
                MessageBox.Show("Нет Записи Для Экспорта!!!", "Info");
            }
        }

        private void btnsearchUsers_Click_1(object sender, EventArgs e)
        {
            string searchValue = searchUsers.Text;
            int rowIndex = 1;

            dataUsers.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            try
            {
                bool valueResulet = true;
                foreach (DataGridViewRow row in dataUsers.Rows)
                {
                    if (row.Cells[rowIndex].Value.ToString().Equals(searchValue))
                    {
                        rowIndex = row.Index;
                        dataUsers.Rows[rowIndex].Selected = true;
                        rowIndex++;
                        valueResulet = false;
                    }
                }
                if (valueResulet != false)
                {
                    MessageBox.Show("Нет такой записи " + searchUsers.Text, "Не найдено");
                    return;
                }
            }
            catch (Exception exc)
            {
                MessageBox.Show(exc.Message);
            }
        }

        private void searchMenu_Click(object sender, EventArgs e)
        {
            string searchValue = textMenuSearch.Text;
            int rowIndex = 1;

            dataMenu.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            try
            {
                bool valueResulet = true;
                foreach (DataGridViewRow row in dataMenu.Rows)
                {
                    if (row.Cells[rowIndex].Value.ToString().Equals(searchValue))
                    {
                        rowIndex = row.Index;
                        dataMenu.Rows[rowIndex].Selected = true;
                        rowIndex++;
                        valueResulet = false;
                    }
                }
                if (valueResulet != false)
                {
                    MessageBox.Show("Нет такой записи " + textMenuSearch.Text, "Не найдено");
                    return;
                }
            }
            catch (Exception exc)
            {
                MessageBox.Show(exc.Message);
            }
        }

        private void btnSearchEmployee_Click(object sender, EventArgs e)
        {
            string searchValue = searchEmployee.Text;
            int rowIndex = 1;

            dataEmployee.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            try
            {
                bool valueResulet = true;
                foreach (DataGridViewRow row in dataEmployee.Rows)
                {
                    if (row.Cells[rowIndex].Value.ToString().Equals(searchValue))
                    {
                        rowIndex = row.Index;
                        dataEmployee.Rows[rowIndex].Selected = true;
                        rowIndex++;
                        valueResulet = false;
                    }
                }
                if (valueResulet != false)
                {
                    MessageBox.Show("Нет такой записи " + searchEmployee.Text, "Не найдено");
                    return;
                }
            }
            catch (Exception exc)
            {
                MessageBox.Show(exc.Message);
            }
        }

        private void btnSearchPost_Click(object sender, EventArgs e)
        {
            string searchValue = searchPost.Text;
            int rowIndex = 1;

            dataPost.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            try
            {
                bool valueResulet = true;
                foreach (DataGridViewRow row in dataPost.Rows)
                {
                    if (row.Cells[rowIndex].Value.ToString().Equals(searchValue))
                    {
                        rowIndex = row.Index;
                        dataPost.Rows[rowIndex].Selected = true;
                        rowIndex++;
                        valueResulet = false;
                    }
                }
                if (valueResulet != false)
                {
                    MessageBox.Show("Нет такой записи " + searchPost.Text, "Не найдено");
                    return;
                }
            }
            catch (Exception exc)
            {
                MessageBox.Show(exc.Message);
            }
        }

        private void btnSearchOtdel_Click(object sender, EventArgs e)
        {
            string searchValue = searchOtdel.Text;
            int rowIndex = 1;

            dataOtdel.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            try
            {
                bool valueResulet = true;
                foreach (DataGridViewRow row in dataOtdel.Rows)
                {
                    if (row.Cells[rowIndex].Value.ToString().Equals(searchValue))
                    {
                        rowIndex = row.Index;
                        dataOtdel.Rows[rowIndex].Selected = true;
                        rowIndex++;
                        valueResulet = false;
                    }
                }
                if (valueResulet != false)
                {
                    MessageBox.Show("Нет такой записи " + searchOtdel.Text, "Не найдено");
                    return;
                }
            }
            catch (Exception exc)
            {
                MessageBox.Show(exc.Message);
            }
        }

        private void btnSearchUsluga_Click(object sender, EventArgs e)
        {
            string searchValue = searchUsluga.Text;
            int rowIndex = 1;

            dataUsluga.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            try
            {
                bool valueResulet = true;
                foreach (DataGridViewRow row in dataUsluga.Rows)
                {
                    if (row.Cells[rowIndex].Value.ToString().Equals(searchValue))
                    {
                        rowIndex = row.Index;
                        dataUsluga.Rows[rowIndex].Selected = true;
                        rowIndex++;
                        valueResulet = false;
                    }
                }
                if (valueResulet != false)
                {
                    MessageBox.Show("Нет такой записи " + searchUsluga.Text, "Не найдено");
                    return;
                }
            }
            catch (Exception exc)
            {
                MessageBox.Show(exc.Message);
            }
        }

        private void btnsearchBuh_Click(object sender, EventArgs e)
        {
            string searchValue = searchBuh.Text;
            int rowIndex = 1;

            dataBuh.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            try
            {
                bool valueResulet = true;
                foreach (DataGridViewRow row in dataBuh.Rows)
                {
                    if (row.Cells[rowIndex].Value.ToString().Equals(searchValue))
                    {
                        rowIndex = row.Index;
                        dataBuh.Rows[rowIndex].Selected = true;
                        rowIndex++;
                        valueResulet = false;
                    }
                }
                if (valueResulet != false)
                {
                    MessageBox.Show("Нет такой записи " + searchBuh.Text, "Не найдено");
                    return;
                }
            }
            catch (Exception exc)
            {
                MessageBox.Show(exc.Message);
            }
        }

        private void btnsearchClient_Click(object sender, EventArgs e)
        {
            string searchValue = searchClient.Text;
            int rowIndex = 2;

            dataClients.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            try
            {
                bool valueResulet = true;
                foreach (DataGridViewRow row in dataClients.Rows)
                {
                    if (row.Cells[rowIndex].Value.ToString().Equals(searchValue))
                    {
                        rowIndex = row.Index;
                        dataClients.Rows[rowIndex].Selected = true;
                        rowIndex++;
                        valueResulet = false;
                    }
                }
                if (valueResulet != false)
                {
                    MessageBox.Show("Нет такой записи " + searchClient.Text, "Не найдено");
                    return;
                }
            }
            catch (Exception exc)
            {
                MessageBox.Show(exc.Message);
            }
        }

        private void btnsearchNomer_Click(object sender, EventArgs e)
        {
            string searchValue = searchNomer.Text;
            int rowIndex = 1;

            dataNomer.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            try
            {
                bool valueResulet = true;
                foreach (DataGridViewRow row in dataNomer.Rows)
                {
                    if (row.Cells[rowIndex].Value.ToString().Equals(searchValue))
                    {
                        rowIndex = row.Index;
                        dataNomer.Rows[rowIndex].Selected = true;
                        rowIndex++;
                        valueResulet = false;
                    }
                }
                if (valueResulet != false)
                {
                    MessageBox.Show("Нет такой записи " + searchNomer.Text, "Не найдено");
                    return;
                }
            }
            catch (Exception exc)
            {
                MessageBox.Show(exc.Message);
            }
        }

        public void sortiroka()
        {
            dataBuh.DataSource = new SortableBindingList<Accounting>(accountings.Objs);
            dataBuh.Sort(dataBuh.Columns["Сумма"], System.ComponentModel.ListSortDirection.Ascending);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            dataBuh.DataSource = new SortableBindingList<Accounting>(accountings.Objs);

            switch (comboBox3.SelectedIndex)
            {
                case 0:
                    {
                        dataBuh.Sort(dataBuh.Columns["Amount"], System.ComponentModel.ListSortDirection.Ascending);
                    }
                    break;

                case 1:
                    dataBuh.Sort(dataBuh.Columns["Amount"], System.ComponentModel.ListSortDirection.Descending);
                    break;
            }
        }



        private void priceMenu_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (Char.IsNumber(e.KeyChar) ||
               (!string.IsNullOrEmpty(priceMenu.Text) && e.KeyChar == ','))
            {
                return;
            }

            e.Handled = true;
        }

        private void seriaPass_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (Char.IsNumber(e.KeyChar) ||
             (!string.IsNullOrEmpty(seriaPass.Text) && e.KeyChar == ','))
            {
                return;
            }

            e.Handled = true;
        }

        private void nmberText_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (Char.IsNumber(e.KeyChar) ||
             (!string.IsNullOrEmpty(nmberText.Text) && e.KeyChar == ','))
            {
                return;
            }

            e.Handled = true;
        }

        private void phoneText_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (Char.IsNumber(e.KeyChar) ||
             (!string.IsNullOrEmpty(phoneText.Text) && e.KeyChar == ','))
            {
                return;
            }

            e.Handled = true;
        }

        private void okladPost_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (Char.IsNumber(e.KeyChar) ||
           (!string.IsNullOrEmpty(okladPost.Text) && e.KeyChar == ','))
            {
                return;
            }

            e.Handled = true;
        }

        private void phoneOtdel_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (Char.IsNumber(e.KeyChar) ||
          (!string.IsNullOrEmpty(phoneOtdel.Text) && e.KeyChar == ','))
            {
                return;
            }

            e.Handled = true;
        }

        private void priceUsluga_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (Char.IsNumber(e.KeyChar) ||
         (!string.IsNullOrEmpty(priceUsluga.Text) && e.KeyChar == ','))
            {
                return;
            }

            e.Handled = true;
        }

        private void priceBuh_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (Char.IsNumber(e.KeyChar) ||
        (!string.IsNullOrEmpty(priceBuh.Text) && e.KeyChar == ','))
            {
                return;
            }

            e.Handled = true;
        }

        private void numberPassClient_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (Char.IsNumber(e.KeyChar) ||
    (!string.IsNullOrEmpty(numberPassClient.Text) && e.KeyChar == ','))
            {
                return;
            }

            e.Handled = true;
        }

        private void seriaPassClient_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (Char.IsNumber(e.KeyChar) ||
   (!string.IsNullOrEmpty(seriaPassClient.Text) && e.KeyChar == ','))
            {
                return;
            }

            e.Handled = true;
        }

        private void phoneClient_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (Char.IsNumber(e.KeyChar) ||
 (!string.IsNullOrEmpty(phoneClient.Text) && e.KeyChar == ','))
            {
                return;
            }

            e.Handled = true;
        }

        private void textNomer_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (Char.IsNumber(e.KeyChar) ||
(!string.IsNullOrEmpty(textNomer.Text) && e.KeyChar == ','))
            {
                return;
            }

            e.Handled = true;
        }

        private void priceNomer_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (Char.IsNumber(e.KeyChar) ||
(!string.IsNullOrEmpty(priceNomer.Text) && e.KeyChar == ','))
            {
                return;
            }

            e.Handled = true;
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

