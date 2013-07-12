using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.Sql;
using System.Data.SqlClient;
using System.IO;
using System.Data.OleDb;
using System.Data.Common;

namespace SqlTester
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            if (ofd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                textBox1.Text = ofd.FileName;
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            string constr = @"Data Source=.\SQLEXPRESS;AttachDbFilename=""c:\users\sagnik\documents\visual studio 2010\Projects\SqlTester\SqlTester\Database1.mdf"";Integrated Security=True;User Instance=True";
            SqlConnection con = new SqlConnection(constr);
            try
            {
                con.Open();
            }
            catch (Exception)
            {
                MessageBox.Show("Did Not Connect");
            }
            con.Close();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (textBox1.Text == "")
            {
                return;
            }
            else if (Path.GetExtension(textBox1.Text) != ".xls" && Path.GetExtension(textBox1.Text) != ".xlsx")
            {
                MessageBox.Show("Please Select An Excel File Only");
            }
            else
            {
                System.Data.DataTable dt = null;
                string path = textBox1.Text;
                string SourceConstr = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" + path + "';Extended Properties= 'Excel 8.0;HDR=No;IMEX=1'";
                OleDbConnection con = new OleDbConnection(SourceConstr);
                con.Open();
                dt = con.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                if (dt == null)
                {
                    MessageBox.Show("dt was null");
                    return;
                }
                String[] excelSheets = new String[dt.Rows.Count];
                int sh = 0;
                foreach (DataRow row in dt.Rows)
                {
                    if (row["TABLE_NAME"].ToString().Contains("$"))
                    {
                        excelSheets[sh] = row["TABLE_NAME"].ToString();
                        sh++;
                    }
                }
                int count = 0;
                for (int z = 0; z < sh; z++)
                {
                    if (!excelSheets[z].ToString().Contains("$Print_Area"))
                    {
                        count = 0;
                        string query = "SELECT top 1 * FROM [" + excelSheets[z] + "];";
                        OleDbCommand command = new OleDbCommand(query, con);
                        OleDbDataReader odr = command.ExecuteReader();
                        string[] names = new string[50];
                        while (odr.Read())
                        {
                            count = odr.FieldCount;
                            int i = 0;
                            while (i < count)
                            {
                                names[i] = cleaner(odr[i].ToString()).Replace(" ", "");
                                i += 1;
                            }
                            break;
                        }
                        int j = 0;
                        string name = Path.GetFileNameWithoutExtension(textBox1.Text);
                        query = "CREATE TABLE " + cleaner(name) + "_" + cleaner(excelSheets[z]).Replace(" ", "").Replace("$", "") + " ( ";
                        for (int i = 0; i < count; i++)
                        {
                            if (names[i] != "")
                            {
                                query += (names[i] + " VARCHAR(100), ");
                                j += 1;
                            }
                        }
                        if(query.Contains("VARCHAR(100)"))
                        {
                            count = j;
                            query = query.Substring(0, query.Length - 2);
                            query = query + ");";
                            if (insertinsql(query) == 1)
                            {
                                MessageBox.Show("Starting copy of Sheet:" + excelSheets[z].Replace("$", ""));
                            }
                            else
                            {
                                return;
                            }
                            odr.Close();
                            string src = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" + path + "';Extended Properties= 'Excel 8.0;HDR=Yes;IMEX=1'";
                            OleDbConnection conn = new OleDbConnection(src);
                            conn.Open();
                            string q = "SELECT * FROM [" + excelSheets[z] + "];";
                            OleDbCommand cmd = new OleDbCommand(q, conn);
                            odr = cmd.ExecuteReader();
                            int check = 1;
                            while (odr.Read())
                            {
                                query = "INSERT INTO " + cleaner(name) + "_" + cleaner(excelSheets[z]).Replace(" ", "").Replace("$", "") + " (";
                                for (int i = 0; i < count; i++)
                                {
                                    query += (names[i] + ", ");
                                }
                                query = query.Substring(0, query.Length - 2);
                                query += ") VALUES (";
                                for (int i = 0; i < count; i++)
                                {
                                    query += ("'" + cleaner(odr[i].ToString()) + "', ");
                                }
                                query = query.Substring(0, query.Length - 2);
                                query += ");";
                                if (check == 1)
                                {
                                    insertinsql(query);
                                }
                                else
                                {
                                    return;
                                }
                            }
                            MessageBox.Show("Sheet: " + excelSheets[z].Replace("$", "") + " Copied Successfully");
                            conn.Close();
                        }
                    }
                }
                con.Close();
            }
        }

        private int insertinsql(string query)
        {
            string constr = @"Data Source=.\SQLEXPRESS;AttachDbFilename=""c:\users\sagnik\documents\visual studio 2010\Projects\SqlTester\SqlTester\Database1.mdf"";Integrated Security=True;User Instance=True";
            SqlConnection con = new SqlConnection(constr);
            try
            {
                con.Open();
            }
            catch (Exception e)
            {
                MessageBox.Show("Did Not Connect" + e.ToString());
                return 0;
            }
            SqlCommand cmd = new SqlCommand(query, con);
            cmd.Connection = con;
            try
            {
                cmd.ExecuteNonQuery();
            }
            catch (Exception e)
            {
                MessageBox.Show(e.ToString() + "\n\n" + query);
                return 0;
            }
            con.Close();
            return 1;
        }

        private string cleaner(string query)
        {
            query = query.Replace("'", "");
            query = query.Replace("\\", "");
            query = query.Replace("/", "");
            query = query.Replace("\"", "");
            query = query.Replace(".", "");
            query = query.Replace(";", ",");
            return query;
        }
    }
}
