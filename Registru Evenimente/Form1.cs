using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Runtime.Remoting.Messaging;
using System.Xml;
using Microsoft.Office.Interop.Excel;
using _Excel = Microsoft.Office.Interop.Excel;
using System.Data.SqlClient;
using System.Configuration;

namespace Registru_Evenimente
{  
    public partial class Form1 : Form
    {
        
        public Form1()
        {
            InitializeComponent();
            textBox3.Select();
           
            using (SqlConnection conn = new SqlConnection(Engine.setting.ConnectionString))
            {
                SqlCommand commandER = new SqlCommand("SELECT * FROM dbo.Eveniment", conn);
                conn.Open();
                using (SqlDataReader reader = commandER.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        string x = (string)reader[1];
                        comboBox1.Items.Add(x.Trim());
                    }
                }
                try
                {
                    string[] f = Directory.GetFiles(Engine.excelPath, "*.xlsx");
                    Engine.Excel(f[0], 1);
                    Engine.CloseExcel();
                }
                catch (Exception)
                {
                    textBox2.Text = "Lipsa fisier excel";
                    textBox2.BackColor = Color.Red;
                }
            }
        }
      
        public void CreateEvent()
        {
            using(SqlConnection conn=new SqlConnection(Engine.setting.ConnectionString))
            {
                bool ok = true;
                SqlCommand commandER = new SqlCommand("SELECT * FROM dbo.Eveniment", conn);
                SqlCommand commandEI = new SqlCommand("INSERT INTO dbo.Eveniment (nume_event,data_event) VALUES (@nume,@zi)", conn);
                conn.Open();
                using(SqlDataReader reader=commandER.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        string x = (string)reader[1];
                        if (x.Trim() == "") 
                        {
                            reader.Close();
                            ok = false;
                            textBox2.Text = "Nume invalid";
                            textBox2.BackColor = Color.Red;
                            timer1.Start();
                            textBox3.Clear();
                            conn.Close();
                            break;
                        }
                        else
                        if (x.Trim() == textBox3.Text)
                        {
                            reader.Close();
                            ok = false;
                            textBox2.Text = "Evenimentul a fost creat deja";
                            textBox2.BackColor = Color.Red;
                            timer1.Start();
                            textBox3.Clear();
                            conn.Close();
                            break;
                        }
                    }
                    if(ok)
                    {
                        reader.Close();
                        Engine.date = DateTime.Now;
                        label3.Text = textBox3.Text;
                        commandEI.Parameters.AddWithValue("@nume", textBox3.Text);
                        commandEI.Parameters.AddWithValue("@zi", Engine.date.Date);
                        commandEI.ExecuteNonQuery();
                        comboBox1.Items.Add(textBox3.Text);
                        textBox3.Clear();
                        textBox1.Select();
                        textBox2.BackColor = Color.GreenYellow;
                        textBox2.Text = "Evenimentul a fost creat";
                        timer1.Start();
                        listBox1.Items.Clear();
                        conn.Close();
                    }
                }
            }
        }
        public void CreateExcel()
        {
            
            Engine.wb = Engine.excel.Workbooks.Add();
            Engine.ws = Engine.wb.Worksheets[1];
            Engine.ws.Name = "Participanti";
            string m="";
            
            using (SqlConnection conn=new SqlConnection(Engine.setting.ConnectionString))
            {  
                int n = 1;
                SqlCommand commandER = new SqlCommand("SELECT * FROM dbo.Eveniment", conn);
                SqlCommand commandPR = new SqlCommand("SELECT * FROM dbo.Persoana", conn);
                conn.Open();
                using(SqlDataReader reader=commandER.ExecuteReader())
                {
                    while(reader.Read())
                    {
                        m = (string)reader[1];
                        if(m.Trim()==comboBox1.Text)
                        {
                            Engine.id = (int)reader[0];
                            Engine.date = (DateTime)reader[2];
                            reader.Close();
                            break;
                        }
                    }
                }
                using(SqlDataReader PR=commandPR.ExecuteReader())
                {
                    while(PR.Read())
                    {
                        if(Engine.id==(int)PR[4])
                        {
                            Engine.ws.Cells[n, 1] = (string)PR[1];
                            Engine.ws.Cells[n, 2] = (string)PR[2];
                            Engine.ws.Cells[n, 3] = (string)PR[3];
                            n++;
                        }
                    }
                    PR.Close();
                }
                Range r = Engine.ws.Range[Engine.ws.Cells[1, 1], Engine.ws.Cells[n, 3]];
                r.EntireColumn.AutoFit();
                FileInfo f = new FileInfo("C:\\Registru\\Rapoarte\\"+ m.ToString() + " " + Engine.date.Day+"."+ Engine.date.Month+"."+ Engine.date.Year+ ".xlsx");
                Engine.wb.SaveAs(f);
                Engine.CloseExcel();
            }
        }

        public void Load()
        { 
            try
            {
                using (SqlConnection conn = new SqlConnection(Engine.setting.ConnectionString))
                {
                   
                    SqlCommand commandER = new SqlCommand("SELECT * FROM dbo.Eveniment", conn);
                    SqlCommand commandPR = new SqlCommand("SELECT * FROM dbo.Persoana", conn);
                    conn.Open();
                    textBox3.Select();
                    using (SqlDataReader ER = commandER.ExecuteReader())
                    {
                        while (ER.Read())
                        {
                            string x = (string)ER[1];
                            if (x.Trim() == comboBox1.Text)
                            {
                                label3.Text = comboBox1.Text;
                                 Engine.id = (int)ER[0];
                                ER.Close();
                                break;
                            }
                        }
                    }
                    using (SqlDataReader PR = commandPR.ExecuteReader())
                    {
                        while (PR.Read())
                        {
                            if (Engine.id == (int)PR[4])
                            {
                                string a = (string)PR[1];
                                string b = (string)PR[2];
                                string c = (string)PR[3];
                                listBox1.Items.Add(a.Trim() + " " + b.Trim() + " " + c.Trim());
                            }
                        }
                    }
                }
            }
            catch (Exception)
            {
                textBox2.Text = "Evenimentul nu a fost gasit";
                textBox2.BackColor = Color.Red;
                timer1.Start();
                label3.Text = "";
            }
        }

        public void Add()
        {
            bool ok=true;
            string cod = textBox1.Text;
            if(cod=="" || cod.Count()>5) { textBox2.Text = "Cod invalid"; textBox2.BackColor = Color.Red; ok = false;timer1.Start(); }
            if (cod.Count()<5)
            {
                int n = 5 - cod.Count();
                StringBuilder s = new StringBuilder();
                while(n!=0)
                {
                    s.Append("0");
                    n--;
                }
                s.Append(cod);
                cod = s.ToString();
            }
            cod = "ro" + cod;
            using (SqlConnection conn= new SqlConnection(Engine.setting.ConnectionString))
            {
                SqlCommand commandER = new SqlCommand("SELECT * FROM dbo.Eveniment", conn);
                SqlCommand commandPR = new SqlCommand("SELECT * FROM dbo.Persoana", conn);
                SqlCommand commandPI = new SqlCommand("INSERT INTO dbo.Persoana (nume,prenume,cod_angajat,persoana_id) VALUES (@nume,@prenume,@cod,@idp)", conn);
                conn.Open();
                using (SqlDataReader ER = commandER.ExecuteReader())
                    while (ER.Read())
                        if(label3.Text==(string)ER[1])
                        {
                            Engine.id = (int)ER[0];
                            break;
                        }
                        
                using(SqlDataReader PR= commandPR.ExecuteReader())
                {
                    while(PR.Read())
                    {
                        string x = (string)PR[3];
                        if(Engine.id==(int)PR[4] && cod==x.Trim())
                        {
                            PR.Close();
                            conn.Close();
                            textBox2.BackColor = Color.Red;
                            textBox2.Text = "Utilizatorul a fost scanat deja";
                            timer1.Start();
                            ok = false;
                            break;
                        }
                    }
                    if(ok)
                    {
                        string[] f = Directory.GetFiles(Engine.excelPath, "*.xlsx");
                        Engine.Excel(f[0], 1);
                        textBox3.Select();
                        if (Engine.VerifyExcel(cod,1))
                        {
                            int n = 1;
                            while (Engine.ws.Cells[n, 3].Value != null)
                            {
                                if (cod == (string)(Engine.ws.Cells[n, 3]).Value)
                                {
                                    PR.Close();
                                    commandPI.Parameters.AddWithValue("@nume", (string)Engine.ws.Cells[n, 1].Value);
                                    commandPI.Parameters.AddWithValue("@prenume", (string)Engine.ws.Cells[n, 2].Value);
                                    commandPI.Parameters.AddWithValue("@cod", (string)Engine.ws.Cells[n, 3].Value);
                                    commandPI.Parameters.AddWithValue("@idp", Engine.id);
                                    commandPI.ExecuteNonQuery();
                                    listBox1.Items.Add((string)Engine.ws.Cells[n, 1].Value + " " + (string)Engine.ws.Cells[n, 2].Value + " " + (string)Engine.ws.Cells[n, 3].Value);
                                    conn.Close();
                                    Engine.CloseExcel();
                                    textBox2.BackColor = Color.GreenYellow;
                                    textBox2.Text = "Utilizatorul a fost scanat cu succes";
                                    timer1.Start();
                                    break;
                                }
                                n++;
                            }
                        }
                        else
                        {
                            PR.Close();
                            commandPI.Parameters.AddWithValue("@nume", " ");
                            commandPI.Parameters.AddWithValue("@prenume", " ");
                            commandPI.Parameters.AddWithValue("@cod", cod);
                            commandPI.Parameters.AddWithValue("@idp", Engine.id);
                            commandPI.ExecuteNonQuery();
                            listBox1.Items.Add(cod);
                            textBox2.BackColor = Color.GreenYellow;
                            textBox2.Text = "Utilizatorul a fost scanat cu succes";
                            timer1.Start();
                            conn.Close();
                        }
                        
                        

                    }
                }
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            textBox1.Text = textBox1.Text + "1";
        }

        private void button3_Click(object sender, EventArgs e)
        {
            textBox1.Text = textBox1.Text + "2";
        }

        private void button4_Click(object sender, EventArgs e)
        {
            textBox1.Text = textBox1.Text + "3";
        }

        private void button5_Click(object sender, EventArgs e)
        {
            textBox1.Text = textBox1.Text + "4";
        }

        private void button6_Click(object sender, EventArgs e)
        {
            textBox1.Text = textBox1.Text + "5";
        }

        private void button7_Click(object sender, EventArgs e)
        {
            textBox1.Text = textBox1.Text + "6";
        }

        private void button8_Click(object sender, EventArgs e)
        {
            textBox1.Text = textBox1.Text + "7";
        }

        private void button9_Click(object sender, EventArgs e)
        {
            textBox1.Text = textBox1.Text + "8";
        }

        private void button10_Click(object sender, EventArgs e)
        {
            textBox1.Text = textBox1.Text + "9";
        }

        private void button11_Click(object sender, EventArgs e)
        {
            textBox1.Text = textBox1.Text + "0";
        }

        private void button13_Click(object sender, EventArgs e)
        {
            CreateEvent();
        }

        private void button14_Click(object sender, EventArgs e)
        {
            listBox1.Items.Clear();
            Load();
            comboBox1.Text = "";
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Add();
            textBox1.Clear();
            textBox1.Select();
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            textBox2.Clear();
            textBox2.BackColor = Color.White;
        }

        private void button15_Click(object sender, EventArgs e)
        {
            CreateExcel();
        }

        private void button12_Click(object sender, EventArgs e)
        {
            wfStatistic myStatistic = new wfStatistic();
            myStatistic.ShowDialog();
        }
    }
}
