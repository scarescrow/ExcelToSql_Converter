using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Services;
using System.Data;
using System.Data.OleDb;
using System.Data.Common;
using System.Data.Sql;
using System.Data.SqlClient;
using System.IO;

namespace FileUploader
{
    /// <summary>
    /// Summary description for Service1
    /// </summary>
    [WebService(Namespace = "http://tempuri.org/")]
    [WebServiceBinding(ConformsTo = WsiProfiles.BasicProfile1_1)]
    [System.ComponentModel.ToolboxItem(false)]
    // To allow this Web Service to be called from script, using ASP.NET AJAX, uncomment the following line. 
    // [System.Web.Script.Services.ScriptService]
    public class Service1 : System.Web.Services.WebService
    {
        [WebMethod]
        public string simple()
        {
            return "Hello World";
        }
        [WebMethod]
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
        [WebMethod]
        private string insertinsql(string query)
        {
            string constr = @"Data Source=.\SQLEXPRESS;AttachDbFilename=""C:\Users\Sagnik\documents\visual studio 2010\Projects\FileUploader\FileUploader\App_Data\Database1.mdf"";Integrated Security=True;User Instance=True";
            SqlConnection con = new SqlConnection(constr);
            try
            {
                con.Open();
            }
            catch (Exception e)
            {
                return e.ToString();
            }
            SqlCommand cmd = new SqlCommand(query, con);
            cmd.Connection = con;
            try
            {
                cmd.ExecuteNonQuery();
            }
            catch (Exception e)
            {
                return e.ToString();
            }
            con.Close();
            return "1";
        }

        [WebMethod]
        public string upload(string file)
        {
            //string path = Path.GetFullPath(file);
            System.Data.DataTable dt = null;
            string SourceConstr = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" + file + "';Extended Properties= 'Excel 8.0;HDR=No;IMEX=1'";
            OleDbConnection con = new OleDbConnection(SourceConstr);
            try
            {
                con.Open();
            }
            catch(Exception e)
            {
                return e.ToString();
            }
            dt = con.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
            if (dt == null)
            {
                return "dt was null";
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
                if (!excelSheets[z].ToString().Contains("Print_Area"))
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
                    string name = Path.GetFileNameWithoutExtension(file);
                    query = "CREATE TABLE " + cleaner(name) + "_" + cleaner(excelSheets[z]).Replace(" ", "").Replace("$", "") + " ( ";
                    for (int i = 0; i < count; i++)
                    {
                        if (names[i] != "")
                        {
                            query += (names[i] + " VARCHAR(100), ");
                            j += 1;
                        }
                    }
                    if (query.Contains("VARCHAR(100)"))
                    {
                        count = j;
                        query = query.Substring(0, query.Length - 2);
                        query = query + ");";
                        string checker = insertinsql(query);
                        if (checker != "1")
                        {
                            return checker;
                        }
                        odr.Close();
                        string src = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" + file + "';Extended Properties= 'Excel 8.0;HDR=Yes;IMEX=1'";
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
                                return "Error";
                            }
                        }
                        conn.Close();
                    }
                }
            }
            con.Close();
            return "Copy Successful...Tables Saved As " + Path.GetFileNameWithoutExtension(file);
        }
    }
}