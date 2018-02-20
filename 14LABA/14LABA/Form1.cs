using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.OleDb;

namespace _14LABA
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            dataGridView1.Columns.Add("Code", "Код");
            dataGridView1.Columns.Add("Surname", "Фамилия");
            dataGridView1.Columns.Add("Name", "Имя");
            dataGridView1.Columns.Add("Secondname", "Отчество");
            dataGridView1.Columns.Add("Tel", "Телефон");

            string conn_param = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + @"БД.mdb";
            OleDbConnection connection = new OleDbConnection(conn_param);
            OleDbCommand command = connection.CreateCommand();
            command.CommandText = "select Код, Фамилия, Имя, Отчество, Телефон from Владельцы";
            connection.Open();
            OleDbDataReader reader = command.ExecuteReader();

            int i = 0;

            try
            {
                while (reader.Read())
                {
                    dataGridView1.Rows.Add();

                    for (int i2 = 0; i2 < 5; i2++) dataGridView1[i2, i].Value = reader.GetValue(i2);

                    ++i;
                }
            }

            finally
            {
                reader.Close();
                connection.Close();
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            int id = dataGridView1.SelectedCells[0].RowIndex;
            int code = Convert.ToInt32(dataGridView1[0, id].Value);

            string conn_param = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + @"БД.mdb";
            OleDbConnection connection = new OleDbConnection(conn_param);
            OleDbCommand command = connection.CreateCommand();
            connection.Open();

            command.CommandText = "DELETE FROM Владельцы WHERE Код = @code";

            command.Parameters.Add("@code", OleDbType.Integer);
            command.Parameters["@code"].Value = code;

            command.ExecuteNonQuery();
            connection.Close();

            dataGridView1.Rows.RemoveAt(id);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            string conn_param = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + @"БД.mdb";
            OleDbConnection connection = new OleDbConnection(conn_param);
            OleDbCommand command = connection.CreateCommand();
            OleDbCommand command2 = connection.CreateCommand();
            connection.Open();

            for (int id = 0; id < dataGridView1.RowCount; id++)
            {
                int code = Convert.ToInt32(dataGridView1[0, id].Value);
                string surname = Convert.ToString(dataGridView1[1, id].Value);
                string name = Convert.ToString(dataGridView1[2, id].Value);
                string secondname = Convert.ToString(dataGridView1[3, id].Value);
                string tel = Convert.ToString(dataGridView1[4, id].Value);

                command.CommandText = "UPDATE Владельцы SET Код = @code, Фамилия = @surname, Имя = @name, Отчество = @secondname, Телефон = @tel WHERE Код = @code";

                command.Parameters.Add("@code", OleDbType.Integer);
                command.Parameters["@code"].Value = code;

                command.Parameters.Add("@surname", OleDbType.VarChar);
                command.Parameters["@surname"].Value = surname;

                command.Parameters.Add("@name", OleDbType.VarChar);
                command.Parameters["@name"].Value = name;

                command.Parameters.Add("@secondname", OleDbType.VarChar);
                command.Parameters["@secondname"].Value = secondname;

                command.Parameters.Add("@tel", OleDbType.VarChar);
                command.Parameters["@tel"].Value = tel;
                command.ExecuteNonQuery();
            }

            connection.Close();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            string conn_param = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + @"БД.mdb";
            OleDbConnection connection = new OleDbConnection(conn_param);
            OleDbCommand command = connection.CreateCommand();
            connection.Open();

            int id = dataGridView1.SelectedCells[0].RowIndex;
            string code = Convert.ToString(dataGridView1[0, id].Value);
            string surname = Convert.ToString(dataGridView1[1, id].Value);
            string name = Convert.ToString(dataGridView1[2, id].Value);
            string secondname = Convert.ToString(dataGridView1[3, id].Value);
            string tel = Convert.ToString(dataGridView1[4, id].Value);
            command.CommandText = "INSERT INTO Владельцы (Код, Фамилия, Имя, Отчество, Телефон) VALUES (@code, @surname, @name, @secondname, @tel)";

            command.Parameters.Add("@code", OleDbType.VarChar);
            command.Parameters["@code"].Value = code;

            command.Parameters.Add("@surname", OleDbType.VarChar);
            command.Parameters["@surname"].Value = surname;

            command.Parameters.Add("@name", OleDbType.VarChar);
            command.Parameters["@name"].Value = name;

            command.Parameters.Add("@secondname", OleDbType.VarChar);
            command.Parameters["@secondname"].Value = secondname;

            command.Parameters.Add("@tel", OleDbType.VarChar);
            command.Parameters["@tel"].Value = tel;
            command.ExecuteNonQuery();

            connection.Close();
        }
    }
}
