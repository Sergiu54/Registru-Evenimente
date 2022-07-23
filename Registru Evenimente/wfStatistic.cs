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
    public partial class wfStatistic : Form
    {
        public wfStatistic()
        {
            InitializeComponent();
        }

        private void Statistic_Load(object sender, EventArgs e)
        {
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
                    this.Close();
                }
            }
        }
        private void Id_Statistic()
        {
            int i = 1;
            List<int> id = new List<int>();
            string cod = textBox1.Text;
            if (cod.Count() < 5)
            {
                int n = 5 - cod.Count();
                StringBuilder s = new StringBuilder();
                while (n != 0)
                {
                    s.Append("0");
                    n--;
                }
                s.Append(cod);
                cod = s.ToString();
            }
            cod = "ro" + cod;
            string[] f = Directory.GetFiles(Engine.excelPath, "*.xlsx");
            Engine.Excel(f[0], 1);
            while (Engine.ws.Cells[i, 3].Value != null)
            {
                if(cod==(string)Engine.ws.Cells[i, 3].Value)
                {
                    label3.Text = ((string)Engine.ws.Cells[i, 1].Value + " " + (string)Engine.ws.Cells[i, 2].Value);
                    break;
                }
                i++;
            }
            using (SqlConnection conn = new SqlConnection(Engine.setting.ConnectionString))
            {
                SqlCommand commandER = new SqlCommand("SELECT * FROM dbo.Eveniment", conn);
                SqlCommand commandPR = new SqlCommand("SELECT * FROM dbo.Persoana", conn);
                conn.Open();
                using (SqlDataReader PR = commandPR.ExecuteReader())
                {
                    while(PR.Read())
                    {
                        string x = (string)PR[3];
                        if(cod==x.Trim())
                        {
                            id.Add((int)PR[4]);
                            if ((string)PR[1] == " ") label3.Text = cod + ":" + id.Count;
                            else label3.Text = label3.Text + ":" + id.Count;

                        }
                        
                    }
                    if (id.Count == 0) label3.Text = label3.Text + ":" + id.Count;
                    PR.Close();
                    Engine.CloseExcel();
                }
                using(SqlDataReader ER=commandER.ExecuteReader())
                {
                    while(ER.Read())
                    {
                        foreach(int x in id)
                            if(x==(int)ER[0])
                            {
                                listBox1.Items.Add((string)ER[1]);
                                break;
                            }
                    }
                }
                conn.Close();
            }

        }
        private void Event_Statistic()
        {
            int i = 1;
            float n = 0;
            float m = 0;
            float p = 0;
            string[] f = Directory.GetFiles(Engine.excelPath, "*.xlsx");
            Engine.Excel(f[0], 1);
            while (Engine.ws.Cells[i, 3].Value != null)
            {
                n++;
                i++;
            }

            using (SqlConnection conn = new SqlConnection(Engine.setting.ConnectionString)) 
            {
                SqlCommand commandER = new SqlCommand("SELECT * FROM dbo.Eveniment", conn);
                SqlCommand commandPR = new SqlCommand("SELECT * FROM dbo.Persoana", conn);
                conn.Open();
                using (SqlDataReader ER = commandER.ExecuteReader())
                {
                    while (ER.Read())
                    {
                        string x = (string)ER[1];
                        if (x.Trim() == comboBox1.Text)
                        {
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
                            listBox2.Items.Add(a.Trim() + " " + b.Trim() + " " + c.Trim());
                            m++;
                            if ((string)PR[1] == " ") n++;
                        }
                    }
                }
                conn.Close();
                Engine.CloseExcel();
            }
            p = m / n * 100;
            label4.Text = comboBox1.Text + ":" +p.ToString("0.00")+"%";   
        }
        private void button3_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            listBox1.Items.Clear();
            Id_Statistic();
            textBox1.Clear();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            listBox2.Items.Clear();
            Event_Statistic();
            comboBox1.Text = "";
        }
    }
}
