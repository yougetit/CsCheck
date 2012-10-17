using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Windows.Forms;
using System.IO;

namespace CsCheck
{
    class MyClass
    {
        static string ConnString = ConfigurationManager.ConnectionStrings["GHHPConnectionString"].ToString();
        public static DataSet GetDataSet(string sqlstring)
        {
            SqlConnection conn = new SqlConnection(ConnString);
            SqlDataAdapter sda = new SqlDataAdapter(sqlstring, conn);
            DataSet ds = new DataSet();
            sda.Fill(ds);
            return ds;
        }

        public static int ExecuteNonQuery(string sqlstring)
        {
            using (SqlConnection conn = new SqlConnection(ConnString))
            {
                int n;
                conn.Open();
                SqlCommand cmd = new SqlCommand(sqlstring, conn);
                n = cmd.ExecuteNonQuery();
                cmd.Dispose();
                return n;
            }

        }

        public static string ExecuteScalar(string sqlstring)
        {
            using (SqlConnection conn = new SqlConnection(ConnString))
            {
                conn.Open();
                SqlCommand cmd = new SqlCommand(sqlstring, conn);
                object obj = cmd.ExecuteScalar();
                if (obj == null)
                {
                    return "0";
                }
                else
                {
                    string value = obj.ToString();
                    return value;
                }
            }
        }

        public static void ToggleConfigEncryption(string exeConfigName)
        {
            // Takes the executable file name without the
            // .config extension.
            try
            {
                // Open the configuration file and retrieve 
                // the connectionStrings section.
                Configuration config = ConfigurationManager.OpenExeConfiguration(exeConfigName);

                ConnectionStringsSection section = config.GetSection("connectionStrings") as ConnectionStringsSection;

                if (section.SectionInformation.IsProtected == false)
                {
                    section.SectionInformation.ProtectSection(
                        "DataProtectionConfigurationProvider");
                    config.Save();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }
    }
}
