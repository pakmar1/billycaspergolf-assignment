using System.Collections.Generic;
using System.Data;
using MySql.Data.MySqlClient;
using System.Data.SqlClient;
using System.Data.OleDb;
using Microsoft.AspNetCore.Http;
using NPOI.SS.UserModel;
using System;
using System.IO;
using NPOI.XSSF.UserModel;

namespace BillyCasperGAssignment.Models
{
    public class BillyCasperContext
    {
        public IFormFile file { get; set; }
        public  long size { get; set; }
        public string extension { get; set; }


        public string ConnectionString { get; set; }

        public BillyCasperContext(string connectionString)
        {
            this.ConnectionString = connectionString;
        }

        private MySqlConnection GetConnection()
        {
            return new MySqlConnection(ConnectionString);
        }

        public MySqlConnection createConnect()
        {
            return GetConnection();
        }

        public List<Costumer> GetAllCostumers()
        {
            List<Costumer> list = new List<Costumer>();

            using (MySqlConnection conn = GetConnection())
            {
                conn.Open();
                MySqlCommand cmd = new MySqlCommand("SELECT * FROM Costumer order by ID", conn);
                using (MySqlDataReader reader = cmd.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        list.Add(new Costumer()
                        {
                            ID = reader.GetInt32("ID"),
                            CreatedOn = reader.GetDateTime("CreatedOn"),
                            ModifiedOn = reader.GetDateTime("ModifiedOn"),
                            Costumer_LastName = reader.GetString("Costumer_LastName"),
                            Costumer_FirstName = reader.GetString("Costumer_FirstName"),
                            AddressLine1 = reader.GetString("Costumer_AddressLine1"),
                            Costumer_City = reader.GetString("Costumer_City"),
                            Costumer_State = reader.GetString("Costumer_State"),
                            Costumer_zip = reader.GetString("Costumer_Zip"),
                            Costumer_Homephone = reader.GetString("Costumer_HomePhone"),
                            Costumer_InternetEmail = reader.GetString("Costumer_InternetEmail")
                        });
                    }
                }
            }
            return list;
        }

        public void deduplicate()
        {
            using (MySqlConnection conn = GetConnection())
            {
                conn.Open();
                MySqlCommand cmd;
                using(MySqlTransaction trans = conn.BeginTransaction())
                {
                    cmd = new MySqlCommand("DELETE t1 FROM Costumer t1 INNER JOIN Costumer t2 WHERE t1.modifiedon < t2.modifiedon " +
                    	"AND t1.costumer_lastname = t2.costumer_lastname" +
                    	"AND t1.costumer_firstname = t2.costumer_firstname" +
                        "AND (t1.costumer_internetemail = t2.costumer_internetemail " +
                        "OR t1.costumer_addressline1 = t2.costumer_addressline1 " +
                        "OR t1.costumer_homephone = t2.costumer_homephone)", conn,trans) ;

                    cmd.ExecuteNonQuery();
                    trans.Commit();
                }
            }
        }



        public void AddCost()
        {
            int Id = 99999;
            string createdon = "1/1/17 12:00 AM", modifiedon = "1/1/17 12:00 AM", firstname = "first", lastname = "last", address = "adress", city = "ct", state = "md", zip = "21211", phone = "23423", email = "sdf@fdsfg.com";

            using (MySqlConnection conn = GetConnection())
            {
                conn.Open();
                MySqlCommand cmd;
                using (MySqlTransaction trans = conn.BeginTransaction())
                {

                    cmd = new MySqlCommand("INSERT INTO Costumer" +
                    "(ID," +
                    "createdon," +
                    "modifiedon," +
                    "costumer_lastname," +
                    "costumer_firstname," +
                    "costumer_addressline1," +
                    "costumer_city," +
                    "costumer_state," +
                    "costumer_zip," +
                    "costumer_homephone," +
                    "costumer_internetemail) " +
                    "values" +
                    "('" + Id + "',STR_TO_DATE('" + createdon + "','%m/%d/%y %h:%i %p'),STR_TO_DATE('" + modifiedon + "','%m/%d/%y %h:%i %p'),'" + lastname + "','"
                        + firstname + "','" + address + "','" + city + "','" + state + "','" + zip + "','"
                        + phone + "','" + email + "');", conn, trans);

                    cmd.Parameters.Clear();
                    cmd.ExecuteNonQuery();
                    trans.Commit();
                }
            }
        }

        public IWorkbook ExportData()
        {
            IWorkbook workbook;
            workbook = new XSSFWorkbook();
            ISheet excelSheet = workbook.CreateSheet("Demo");
            List<Costumer> list = new List<Costumer>();
            IRow row; int num = 1;
            using (MySqlConnection conn = createConnect())
            {
                conn.Open();
                MySqlCommand cmd = new MySqlCommand("SELECT * FROM Costumer order by modifiedon", conn);
                using (MySqlDataReader reader = cmd.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        list.Add(new Costumer()
                        {
                            ID = reader.GetInt32("ID"),
                            CreatedOn = reader.GetDateTime("CreatedOn"),
                            ModifiedOn = reader.GetDateTime("ModifiedOn"),
                            Costumer_LastName = reader.GetString("Costumer_LastName"),
                            Costumer_FirstName = reader.GetString("Costumer_FirstName"),
                            AddressLine1 = reader.GetString("Costumer_AddressLine1"),
                            Costumer_City = reader.GetString("Costumer_City"),
                            Costumer_State = reader.GetString("Costumer_State"),
                            Costumer_zip = reader.GetString("Costumer_Zip"),
                            Costumer_Homephone = reader.GetString("Costumer_HomePhone"),
                            Costumer_InternetEmail = reader.GetString("Costumer_InternetEmail")
                        });
                    }
                }

                row = excelSheet.CreateRow(num);
                row.CreateCell(0).SetCellValue("ID");
                row.CreateCell(1).SetCellValue("CreatedOn");
                row.CreateCell(2).SetCellValue("ModifiedOn");
                row.CreateCell(3).SetCellValue("Costumer_LastName");
                row.CreateCell(4).SetCellValue("Costumer_FirstName");
                row.CreateCell(5).SetCellValue("AddressLine1");
                row.CreateCell(6).SetCellValue("Costumer_City");
                row.CreateCell(7).SetCellValue("Costumer_State");
                row.CreateCell(8).SetCellValue("Costumer_zip");
                row.CreateCell(9).SetCellValue("Costumer_Homephone");
                row.CreateCell(10).SetCellValue("Costumer_InternetEmail");

                foreach (var item in list)
                {
                    row = excelSheet.CreateRow(num);
                    row.CreateCell(0).SetCellValue(item.ID);
                    row.CreateCell(1).SetCellValue(item.CreatedOn);
                    row.CreateCell(2).SetCellValue(item.ModifiedOn);
                    row.CreateCell(3).SetCellValue(item.Costumer_LastName);
                    row.CreateCell(4).SetCellValue(item.Costumer_FirstName);
                    row.CreateCell(5).SetCellValue(item.AddressLine1);
                    row.CreateCell(6).SetCellValue(item.Costumer_City);
                    row.CreateCell(7).SetCellValue(item.Costumer_State);
                    row.CreateCell(8).SetCellValue(item.Costumer_zip);
                    row.CreateCell(9).SetCellValue(item.Costumer_Homephone);
                    row.CreateCell(10).SetCellValue(item.Costumer_InternetEmail);
                    num++;
                }
            }

            return workbook;
        }

        public void AddCostumers(ISheet sheet)
        {

            int Id = 0;
            string createdon = "", modifiedon = "", firstname = "", lastname = "", address = "", city = "", state = "", zip = "", phone = "", email = "";


            using (MySqlConnection conn = GetConnection())
            {
                conn.Open();
                IRow row;

                MySqlCommand cmd;
                using (MySqlTransaction trans = conn.BeginTransaction())
                {
                    for (int i = 1; i < sheet.LastRowNum; i++)
                    {
                        row = sheet.GetRow(i);
                        if (Convert.ToInt32(row.GetCell(0).ToString()) == 0) { Id = 0; continue; }
                        else { Id = Convert.ToInt32(row.GetCell(0).ToString()); }

                        if (Convert.ToDateTime(row.GetCell(1).ToString()) == null) { createdon = row.GetCell(1).ToString(); }
                        else { createdon = "1/11/10 12:10 AM"; }
                        //Convert.ToDateTime(row.GetCell(1)).ToString("M/d/yy h:mm tt"); }

                        if (Convert.ToDateTime(row.GetCell(2).ToString()) == null) { modifiedon = "1/11/10 12:10 AM"; }
                        else { modifiedon = "1/11/10 12:10 AM"; }
                        //Convert.ToDateTime(row.GetCell(2)).ToString("M/d/yy h:mm tt"); }

                        if (row.GetCell(3) != null) { lastname = row.GetCell(3).ToString(); }
                        else { lastname = ""; }

                        if (row.GetCell(4) != null) { firstname = row.GetCell(4).ToString(); }
                        else { firstname = ""; }

                        if (row.GetCell(5) != null) { address = row.GetCell(5).ToString(); }
                        else { address = ""; }

                        if (row.GetCell(6) != null) { city = row.GetCell(6).ToString(); }
                        else { city = ""; }

                        if (row.GetCell(7) != null) { state = row.GetCell(7).ToString(); }
                        else { state = ""; }

                        if (row.GetCell(8) != null) { zip = row.GetCell(8).ToString(); }
                        else { zip = ""; }

                        if (row.GetCell(9) != null) { phone = row.GetCell(9).ToString(); }
                        else { phone = ""; }

                        if (row.GetCell(10) != null) { email = row.GetCell(10).ToString(); }
                        else { email = ""; }



                        cmd = new MySqlCommand("INSERT INTO Costumer" +
                        "(ID," +
                        "createdon," +
                        "modifiedon," +
                        "costumer_lastname," +
                        "costumer_firstname," +
                        "costumer_addressline1," +
                        "costumer_city," +
                        "costumer_state," +
                        "costumer_zip," +
                        "costumer_homephone," +
                        "costumer_internetemail) values" +
                        "(@ID,STR_TO_DATE(@createdon,'%m/%d/%y %h:%i %p'),STR_TO_DATE(@modifiedon,'%m/%d/%y %h:%i %p'),@costumer_lastname,@costumer_firstname,@costumer_addressline1,@costumer_city,@costumer_state,@costumer_zip,@costumer_homephone,@costumer_internetemail)", conn,trans);
                        cmd.Parameters.Clear();

                        cmd.Parameters.AddWithValue("ID", Id);
                        cmd.Parameters.AddWithValue("createdon", createdon);
                        cmd.Parameters.AddWithValue("modifiedon", modifiedon);
                        cmd.Parameters.AddWithValue("costumer_lastname", lastname);
                        cmd.Parameters.AddWithValue("costumer_firstname", firstname);
                        cmd.Parameters.AddWithValue("costumer_addressline1", address);
                        cmd.Parameters.AddWithValue("costumer_city", city);
                        cmd.Parameters.AddWithValue("costumer_state", state);
                        cmd.Parameters.AddWithValue("costumer_zip", zip);
                        cmd.Parameters.AddWithValue("costumer_homephone", phone);
                        cmd.Parameters.AddWithValue("costumer_internetemail", email);

                        cmd.ExecuteNonQuery();
                        cmd.Dispose(); 

                    }
                    trans.Commit();
                }
            }
        }
    }
}

