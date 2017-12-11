using System;
using System.Drawing;
using System.Windows.Forms;
using System.Data.OleDb;
using MetroFramework;
using System.Net.Mail;

namespace BackLogProgramNew
{
    public partial class Form1 : MetroFramework.Forms.MetroForm
    {
        string pathdb = "";
        int conOtvet = 0;
        int connOtvet = 0;

        int countFio = 0;
        int countEmail = 0;

        bool useBD = false;
        bool ifDontEmail = false;
        bool statusReg = false;

        public Form1()
        {
            InitializeComponent();

            panelReg.Visible = false;
            panelReg.Enabled = false;
        }

        private  bool IsValidEmail(string email)
        {
            int z = 0;

            for (int i = 0; i < email.Length; i++)
            {
                if (email[i] == '.')
                {
                    z++;
                }
            }
            try
            {
                var mail = new MailAddress(email);

                if (z >= 1)
                {
                    return true;
                }
                else
                    return false;
            }
            catch
            {
                return false;
            }
        }


        private void button1_Click(object sender, EventArgs e)
        {
            if (textBox2.Text != "" && textBox1.Text != ""
                                    && textBox3.Text != ""
                                    && comboBox2.Text != ""
                                    && textBox5.Text != ""
                                    && textBox6.Text != ""
                                    && textBox7.Text != ""
                                    && comboBox1.Text != "")
            {

                using (var connection = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=BackLogDb.accdb;Persist Security Info=False;"))
                {
                    connection.Open();
                    using (var command = connection.CreateCommand())
                    {
                        command.CommandText = @"INSERT INTO student(Фамилия, Имя, Отчество, Email, Логин, Пароль, Список_дисциплин, Институт, statusReg)
                                VALUES (@THE_FAM, @THE_NAME, @THE_OTCH, @THE_MAIL, @THE_LOGIN, @THE_PASS, @THE_SPISOK, @THE_INST, @THE_STATUSREG);";
                        command.Parameters.Add("@THE_FAM", OleDbType.VarChar).Value = textBox1.Text;
                        command.Parameters.Add("@THE_NAME", OleDbType.VarChar).Value = textBox2.Text;
                        command.Parameters.Add("@THE_OTCH", OleDbType.VarChar).Value = textBox3.Text;
                        command.Parameters.Add("@THE_MAIL", OleDbType.VarChar).Value = comboBox2.Text;
                        command.Parameters.Add("@THE_LOGIN", OleDbType.VarChar).Value = textBox5.Text;
                        command.Parameters.Add("@THE_PASS", OleDbType.VarChar).Value = textBox6.Text;
                        command.Parameters.Add("@THE_SPISOK", OleDbType.VarChar).Value = textBox7.Text;
                        command.Parameters.Add("@THE_INST", OleDbType.VarChar).Value = comboBox1.Text;
                        //статус регистрации на сайте hi-edu.ru
                        if (statusReg)
                            command.Parameters.Add("@THE_STATUSREG", OleDbType.VarChar).Value = 1;
                        else
                            command.Parameters.Add("@THE_STATUSREG", OleDbType.VarChar).Value = 0;

                        var i = command.ExecuteNonQuery();

                        if (i <= 0)
                        {
                            throw new Exception("Невозможно создать запись.");
                        }
                    }
                    connection.Close();
                    MetroMessageBox.Show(Owner, "Регистрация прошла успешно", "   Поздравляю!", MessageBoxButtons.OK, MessageBoxIcon.Question);
                }

                textBox1.Clear();
                textBox2.Clear();
                textBox3.Clear();
                textBox5.Clear();
                textBox6.Clear();
                textBox7.Clear();
                label9.Text = "";
                sovpadFIO.Text = "";
                comboBox2.Enabled = false;
                textBox5.Enabled = false;
                textBox6.Enabled = false;
                textBox7.Enabled = false;
                comboBox1.Enabled = false;
                // comboBox по умолчанию выбранным, первый элемент списка
                // comboBox.SelectedIndex = 0;
                comboBox1.Text = ""; 
                comboBox2.Enabled = false;
                comboBox2.Text = "";
                timer.Enabled = false;
                login_panel.Visible = false;
                pass_panel.Visible = false;



            }
            else
                MetroMessageBox.Show(Owner, "Для регистрации нужно заполнить все поля", "   Внимание!", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }

        //Выбрать(подключить) БД
        private void button2_Click(object sender, EventArgs e)
        {
            
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            openFileDialog1.FileName = "";
            openFileDialog1.Filter = "Access Database Files (*.mdb;*.accdb)|*.mdb;*.accdb";
            openFileDialog1.Title = "Select an Access Database File";
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                pathdb = openFileDialog1.FileName;
                MetroMessageBox.Show(Owner, "Выбрана " + openFileDialog1.SafeFileName + " База Данных !", "   Внимание!", MessageBoxButtons.OK, MessageBoxIcon.Question);
                useBD = true;
            }


        }

        // Проверка e-mail на правильность ввода и на бесповторность
        private void nextreg_Click(object sender, EventArgs e)
        {
            IsValidEmail(comboBox2.Text);

            if (textBox1.Text != "" & textBox2.Text != "" & textBox3.Text != "" & comboBox2.Text != "" & IsValidEmail(comboBox2.Text) == true)
            {
                
                using (OleDbConnection con = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + pathdb + " ;Persist Security Info=False;"))
                {

                    using (OleDbCommand Command = new OleDbCommand(("select count(SecondName and FirstName and PatronymicName and EMail) from people where SecondName='" + textBox1.Text + "' and FirstName='" + textBox2.Text + "' and PatronymicName='" + textBox3.Text + "' and EMail='" + comboBox2.Text + "'"), con))
                    {
                        con.Open();
                        OleDbDataReader DB_Reader = Command.ExecuteReader();
                        if (DB_Reader.HasRows)
                        {
                            DB_Reader.Read();
                            int id = DB_Reader.GetInt32(0);
                            conOtvet = id;
                        }
                        con.Close();
                    }


                }//заканчивается con

                using (var conn = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=BackLogDb.accdb;Persist Security Info=False;"))
                {
                    using (OleDbCommand Command2 = new OleDbCommand(("select count(Фамилия and Имя and Отчество and Email) from student where Фамилия='" + textBox1.Text + "' and Имя='" + textBox2.Text + "' and Отчество='" + textBox3.Text + "' and Email='" + comboBox2.Text + "'"), conn))
                    {
                        conn.Open();
                        OleDbDataReader DB_Reader = Command2.ExecuteReader();
                        if (DB_Reader.HasRows)
                        {
                            DB_Reader.Read();
                            int id = DB_Reader.GetInt32(0);
                            connOtvet = id;
                        }
                    }


                }//заканчивается conn

               
                if ((conOtvet | connOtvet) == 0)
                {
                    label9.Text = "Совпадений не найдено. Продолжайте регистрацию";
                    textBox5.Clear();
                    textBox6.Clear();
                    textBox5.Enabled = true;
                    textBox6.Enabled = true;
                    textBox7.Enabled = true;
                    comboBox1.Enabled = true;
                    timer.Enabled = true;

                }
                else
                {
                    label9.Text = "Вы уже есть в системе. Заполните оставшиеся поля ";
                    textBox5.Text = "Существует";
                    textBox6.Text = "Существует";
                    textBox5.Enabled = false;
                    textBox6.Enabled = false;
                    textBox7.Enabled = true;
                    comboBox1.Enabled = true;
                    timer.Enabled = false;
                    login_panel.Visible = false;
                    pass_panel.Visible = false;
                }


                if (statusReg == false & ifDontEmail == true)
                {
                   // label9.Text = "Совпадений не найдено. Продолжайте регистрацию";
                    label9.Text = "Заполните оставшиеся поля";
                    textBox5.Clear();
                    textBox6.Clear();
                    textBox5.Enabled = true;
                    textBox6.Enabled = true;
                    textBox7.Enabled = true;
                    comboBox1.Enabled = true;
                    timer.Enabled = true;

                }
                else
                {
                    //label9.Text = "Вы уже есть в системе. Заполните оставшиеся поля ";
                    label9.Text = "Заполните оставшиеся поля";
                    textBox5.Text = "Существует";
                    textBox6.Text = "Существует";
                    textBox5.Enabled = false;
                    textBox6.Enabled = false;
                    textBox7.Enabled = true;
                    comboBox1.Enabled = true;
                    timer.Enabled = false;
                    login_panel.Visible = false;
                    pass_panel.Visible = false;
                }
            }
            else
            {
                MetroMessageBox.Show(Owner, "Заполните поля ФИО и проверьте правильность E-mail(формат: email@text.com)", "   Внимание!", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                textBox5.Enabled = false;
                textBox6.Enabled = false;
                textBox7.Enabled = false;
                comboBox1.Enabled = false;

            }

        }

        //проверка ФИО на совпадения и на e-mail(существует ли)
        private void findEmail_Click(object sender, EventArgs e)
        {

            sovpadFIO.Text = "";
            label9.Text = "";
            timer.Enabled = false;
            if (login_panel.Visible == true)
                login_panel.Visible = false;
            if (pass_panel.Visible == true)
                pass_panel.Visible = false;
            textBox5.Text = "";
            textBox6.Text = "";
            textBox7.Text = "";
            textBox5.Enabled = false;
            textBox6.Enabled = false;
            textBox7.Enabled = false;
            comboBox1.Enabled = false;
            comboBox2.Items.Clear();
            comboBox2.Text = "";
            comboBox1.Text = "";
            

            if (textBox1.Text == "" | textBox2.Text == "" | textBox3.Text == "" | useBD == false)
            {
                MetroMessageBox.Show(Owner, "Заполните поля ФИО и подключите Базу Данных", "   Внимание!", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                comboBox2.Enabled = false;
            }
            else
            {
                comboBox2.Enabled = true;
            }

            //Проверка count(ФИО)
            if (textBox1.Text != "" & textBox2.Text != "" & textBox3.Text != "" & useBD == true)
            {

                using (OleDbConnection conFio = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + pathdb + " ;Persist Security Info=False;"))
                {

                    using (OleDbCommand CommandFio = new OleDbCommand(("select count(SecondName and FirstName and PatronymicName) from people where SecondName='" + textBox1.Text + "' and FirstName='" + textBox2.Text + "' and PatronymicName='" + textBox3.Text + "'"), conFio))
                    {
                        conFio.Open();
                        OleDbDataReader DB_Reader = CommandFio.ExecuteReader();
                        if (DB_Reader.HasRows)
                        {
                            DB_Reader.Read();
                            int id = DB_Reader.GetInt32(0);
                            countFio = id;
                        }
                        conFio.Close();
                    }


                }//заканчивается conFio




                if (textBox1.Text != "" & textBox2.Text != "" & textBox3.Text != "" & useBD == true & countFio >= 1)
                {
                    //проверка count(e-mail) по ФИО у данного пользователя в БД KCT.accdb
                    if (textBox1.Text != "" & textBox2.Text != "" & textBox3.Text != "" & useBD == true)
                    {

                        using (OleDbConnection conEmail = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + pathdb + " ;Persist Security Info=False;"))
                        {

                            using (OleDbCommand CommandEmail = new OleDbCommand(("select count(EMail) from people where SecondName='" + textBox1.Text + "' and FirstName='" + textBox2.Text + "' and PatronymicName='" + textBox3.Text + "'"), conEmail))
                            {
                                conEmail.Open();
                                OleDbDataReader DB_Reader = CommandEmail.ExecuteReader();
                                if (DB_Reader.HasRows)
                                {
                                    DB_Reader.Read();
                                    int id = DB_Reader.GetInt32(0);
                                    countEmail = id;
                                }
                                conEmail.Close();
                            }
                        }
                    }
                    //закончена проверка count(email)

                    // если есть e-mail у данного ФИО
                    if (textBox1.Text != "" & textBox2.Text != "" & textBox3.Text != "" & useBD == true & countEmail >= 1)
                    {
                        OleDbConnection connectionFind = new OleDbConnection();

                        connectionFind.ConnectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + pathdb + " ;Persist Security Info=False;";

                        OleDbCommand commandFind = new OleDbCommand();
                        commandFind.CommandText = "Select EMail from people where SecondName='" + textBox1.Text + "' and FirstName='" + textBox2.Text + "' and PatronymicName='" + textBox3.Text + "'";

                        commandFind.Connection = connectionFind;
                        try
                        {
                            connectionFind.Open();

                            OleDbDataReader dr = commandFind.ExecuteReader();

                            while (dr.Read())
                            {
                                //Добавляем новую строку в элементы управления,
                                // где в качестве источника значения
                                // указывается имя столбца.
                                if (!comboBox2.Items.Contains(dr["EMail"]))
                                {
                                    comboBox2.Items.Add(dr["EMail"]);
                                }
                            }
                            comboBox2.Enabled = true;
                            sovpadFIO.Text = "Найдено совпадение !\nЕсли вы регистрировались ранее, то выберите\nсвой e-mail из предложенного ниже списка.";

                        }
                        catch (Exception ex)
                        {
                            //Сообщение об ошибке

                            MetroMessageBox.Show(Owner, "Ошибка получения данных: " + Environment.NewLine + ex.ToString(), "   Внимание!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                        finally
                        {
                            //Закрываем соединение с базой данных.
                            connectionFind.Close();
                        }

                        //Закончилось пополнение comboBox элементами e-mail 
                    }
                    else // если нет e-mail у данного ФИО
                    {

                        panelReg.Visible = true;
                        panelReg.Enabled = true;

                        textBox1.Enabled = false;
                        textBox2.Enabled = false;
                        textBox3.Enabled = false;
                        textBox5.Enabled = false;
                        textBox6.Enabled = false;
                        textBox7.Enabled = false;
                        comboBox1.Enabled = false;
                        comboBox2.Enabled = false;
                        findEmail.Enabled = false;
                        nextreg.Enabled = false;
                        button1.Enabled = false;
                        button2.Enabled = false;
                        timer.Enabled = false;
                        login_panel.Visible = false;
                        pass_panel.Visible = false;
                    }


                }
                else
                {
                    sovpadFIO.Text = "Совпадений не найдено. \nПроверьте правильность ввода ФИО";

                    textBox5.Enabled = false;
                    textBox6.Enabled = false;
                    textBox7.Enabled = false;
                    comboBox1.Enabled = false;
                }

                

            }
            else
            {
                sovpadFIO.Text = "Совпадений не найдено. \nПроверьте правильность ввода ФИО";
            }

            //если подключения к базе нет, то в label совпадения ФИО ничего не писать
            if (useBD == false)
            { sovpadFIO.Text = ""; }
        }

        // таймер на полосы (красная и зеленая)
        private void timer_Tick(object sender, EventArgs e)
        {
            if (textBox5.Text == "" & textBox5.Text != "Существует")
            {
                login_panel.Visible = true;
                login_panel.BackColor = Color.FromArgb(255, 192, 192);
            }
            else
            {
                login_panel.Visible = true;
                login_panel.BackColor = Color.FromArgb(144, 238, 144);
            }
            if (textBox6.Text == "" & textBox5.Text != "Существует")
            {
                pass_panel.Visible = true;
                pass_panel.BackColor = Color.FromArgb(255, 192, 192);
            }
            else
            {
                pass_panel.Visible = true;
                pass_panel.BackColor = Color.FromArgb(144, 238, 144);
            }

        }

        private void btYes_Click(object sender, EventArgs e)
        {
            statusReg = true;
            panelReg.Visible = false;
            panelReg.Enabled = false;
            sovpadFIO.Text = "Продолжайте регистрацию";
            ifDontEmail = true;

            textBox1.Enabled = true;
            textBox2.Enabled = true;
            textBox3.Enabled = true;
            textBox5.Enabled = true;
            textBox6.Enabled = true;
            textBox7.Enabled = true;
            comboBox1.Enabled = true;
            comboBox2.Enabled = true;
            findEmail.Enabled = true;
            nextreg.Enabled = true;
            button1.Enabled = true;
            button2.Enabled = true;
            timer.Enabled = false;
            login_panel.Visible = false;
            pass_panel.Visible = false;
        }

        private void btNo_Click(object sender, EventArgs e)
        {
            statusReg = false;
            panelReg.Visible = false;
            panelReg.Enabled = false;
            sovpadFIO.Text = "Продолжайте регистрацию";
            ifDontEmail = true;

            textBox1.Enabled = true;
            textBox2.Enabled = true;
            textBox3.Enabled = true;
            textBox5.Enabled = true;
            textBox6.Enabled = true;
            textBox7.Enabled = true;
            comboBox1.Enabled = true;
            comboBox2.Enabled = true;
            findEmail.Enabled = true;
            nextreg.Enabled = true;
            button1.Enabled = true;
            button2.Enabled = true;
            timer.Enabled = false;
            login_panel.Visible = false;
            pass_panel.Visible = false;
        }

       
    }
}
