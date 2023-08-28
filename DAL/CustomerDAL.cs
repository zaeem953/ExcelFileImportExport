//using adodotnet.Models;
using adodotnet.Utility;
using ExcelFileImportExport.Models;
using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Data;

namespace adodotnet.DAL
{
    public class CustomerDAL
    {
        private readonly string connectionString;

        public CustomerDAL()
        {
            connectionString = ConnectionString._ConnectionString;
        }

        public IEnumerable<ExcelCustomer> GetAllExcelCustomers()
        {
            List<ExcelCustomer> ExcelCustomers = new List<ExcelCustomer>();

            using (MySqlConnection con = new MySqlConnection(connectionString))
            {
                MySqlCommand cmd = new MySqlCommand("SELECT * FROM exceldata", con);
                con.Open();
                MySqlDataReader reader = cmd.ExecuteReader();

                while (reader.Read())
                {
                    ExcelCustomer customer = new ExcelCustomer();
                    customer.Id = Convert.ToInt32(reader["Id"]);
                    customer.FirstName = reader["FirstName"].ToString();
                    customer.LastName = reader["LastName"].ToString();
                    customer.Gender = reader["Gender"].ToString();
                    customer.Country = reader["Country"].ToString();
                    customer.Age = Convert.ToInt32(reader["Age"]);

                    ExcelCustomers.Add(customer);
                }

                con.Close();
            }

            return ExcelCustomers;
        }

        public void AddExcelCustomer(ExcelCustomer customer)
        {
            using (MySqlConnection con = new MySqlConnection(connectionString))
            {
                MySqlCommand cmd = new MySqlCommand("INSERT INTO exceldata (FirstName, LastName, Gender, Country, Age) VALUES (@FirstName, @LastName, @Gender, @Country, @Age);", con);
                cmd.CommandType = CommandType.StoredProcedure;

                cmd.Parameters.AddWithValue("p_FirstName", customer.FirstName);
                cmd.Parameters.AddWithValue("p_LastName", customer.LastName);
                cmd.Parameters.AddWithValue("p_Gender", customer.Gender);
                cmd.Parameters.AddWithValue("p_Country", customer.Country);
                cmd.Parameters.AddWithValue("p_Age", customer.Age);

                con.Open();
                cmd.ExecuteNonQuery();
                con.Close();
            }
        }

        public void UpdateExcelCustomer(ExcelCustomer customer)
        {
            using (MySqlConnection con = new MySqlConnection(connectionString))
            {
                MySqlCommand cmd = new MySqlCommand("UpdateRecord", con);
                cmd.CommandType = CommandType.StoredProcedure;

                cmd.Parameters.AddWithValue("p_FirstName", customer.FirstName);
                cmd.Parameters.AddWithValue("p_LastName", customer.LastName);
                cmd.Parameters.AddWithValue("p_Gender", customer.Gender);
                cmd.Parameters.AddWithValue("p_Country", customer.Country);
                cmd.Parameters.AddWithValue("p_Age", customer.Age);
                con.Open();
                cmd.ExecuteNonQuery();
                con.Close();
            }
        }

        public void DeleteExcelCustomer(int? id)
        {
            using (MySqlConnection con = new MySqlConnection(connectionString))
            {
                MySqlCommand cmd = new MySqlCommand("DeleteRecord", con);
                cmd.CommandType = CommandType.StoredProcedure;

                cmd.Parameters.AddWithValue("p_Id", id);

                con.Open();
                cmd.ExecuteNonQuery();
                con.Close();
            }
        }

        public ExcelCustomer GetExcelCustomersData(int? id)
        {
            ExcelCustomer customer = new ExcelCustomer();

            using (MySqlConnection con = new MySqlConnection(connectionString))
            {
                MySqlCommand cmd = new MySqlCommand("SELECT * FROM ExcelCustomer WHERE Id = @Id", con);
                cmd.Parameters.AddWithValue("@Id", id);
                con.Open();
                MySqlDataReader reader = cmd.ExecuteReader();

                while (reader.Read())
                {
                    customer.Id = Convert.ToInt32(reader["Id"]);
                    customer.FirstName = reader["FirstName"].ToString();
                    customer.LastName = reader["LastName"].ToString();
                    customer.Gender = reader["Gender"].ToString();
                    customer.Country = reader["Country"].ToString();
                    customer.Age = Convert.ToInt32(reader["Age"]);

                }

                con.Close();
            }

            return customer;
        }

    }
}
