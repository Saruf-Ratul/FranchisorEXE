using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FranchisorEXE
{
    class Program
    {
        static void Main(string[] args)
        {
            //Franchise
            //FinancialRatioManager fr = new FinancialRatioManager();
            //fr.Get_FinancialRatio_Data();
            //Console.WriteLine("Financial Ratio Data Loaded");

            //TBK
            //Saruf Ratul//-- Updated by Saruf Ratul on 2024-09-10

            //FinancialRatioManagerTBK frTBK = new FinancialRatioManagerTBK();
            //frTBK.Get_FinancialRatio_DataTBK();

            //Console.WriteLine("Financial Ratio Data TBK Loaded");

            //AM
            //Saruf Ratul//-- Updated by Saruf Ratul on 2024-10-12

            FinancialRatioManagerAM frAM = new FinancialRatioManagerAM();
            frAM.Get_FinancialRatio_DataAM();

            Console.WriteLine("Financial Ratio Data AM Loaded");


            //Demo
            //Saruf Ratul//-- Updated by Saruf Ratul on 2024-10-20

            //FinancialRatioManagerDemo frDemo = new FinancialRatioManagerDemo();
            //frDemo.Get_FinancialRatio_DataDemo();

            //Console.WriteLine("Financial Ratio Data Demo Loaded");


            //FLC
            //Saruf Ratul//-- Updated by Saruf Ratul on 2024-11-19

            //FinancialRatioManagerFLC fr = new FinancialRatioManagerFLC();
            //fr.Get_FinancialRatio_DataFLC();

            //Console.WriteLine("Financial Ratio Data FLC Loaded");
        }
    }
}


//using System;
//using System.Data;
//using System.Data.SqlClient;

//namespace FranchisorEXE
//{
//    class Program
//    {
//        static void Main(string[] args)
//        {
//            // Server to Dev QBO Sync
//            // Connection strings
//            string prodConnectionString = "Data Source=172.168.90.16;Initial Catalog=Franchisor;User ID=Mobilizedba;Password=Mobilizedba;";
//            string devConnectionString = "Data Source=18.222.27.177;Initial Catalog=Franchisor;User ID=Mobilizedba;Password=Mobilizedba;";

//            try
//            {
//                using (SqlConnection prodConnection = new SqlConnection(prodConnectionString))
//                using (SqlConnection devConnection = new SqlConnection(devConnectionString))
//                {
//                    // Query to get records from production tbl_Location
//                    string selectQuery = @"SELECT LocationID, FranchisorID, AccessToken, RefreshToken, BQOFileID FROM tbl_Location";

//                    // Open production connection and fetch data
//                    SqlDataAdapter dataAdapter = new SqlDataAdapter(selectQuery, prodConnection);
//                    DataTable prodData = new DataTable();
//                    dataAdapter.Fill(prodData);

//                    // Open dev connection for inserting/updating data
//                    devConnection.Open();
//                    foreach (DataRow row in prodData.Rows)
//                    {
//                        string locationId = row["LocationID"].ToString();
//                        string franchisorId = row["FranchisorID"].ToString();
//                        string accessToken = row["AccessToken"].ToString();
//                        string refreshToken = row["RefreshToken"].ToString();
//                        string bqoFileId = row["BQOFileID"].ToString();

//                        // Upsert query for dev tbl_Location
//                        string upsertQuery = @"
//                            IF EXISTS (SELECT 1 FROM tbl_Location WHERE LocationID = @LocationID AND FranchisorID = @FranchisorID)
//                            BEGIN
//                                UPDATE tbl_Location
//                                SET AccessToken = @AccessToken, RefreshToken = @RefreshToken, BQOFileID = @BQOFileID
//                                WHERE LocationID = @LocationID AND FranchisorID = @FranchisorID;
//                            END
//                            ELSE
//                            BEGIN
//                                INSERT INTO tbl_Location (LocationID, FranchisorID, AccessToken, RefreshToken, BQOFileID)
//                                VALUES (@LocationID, @FranchisorID, @AccessToken, @RefreshToken, @BQOFileID);
//                            END";

//                        using (SqlCommand cmd = new SqlCommand(upsertQuery, devConnection))
//                        {
//                            cmd.Parameters.AddWithValue("@LocationID", locationId);
//                            cmd.Parameters.AddWithValue("@FranchisorID", franchisorId);
//                            cmd.Parameters.AddWithValue("@AccessToken", accessToken);
//                            cmd.Parameters.AddWithValue("@RefreshToken", refreshToken);
//                            cmd.Parameters.AddWithValue("@BQOFileID", bqoFileId);

//                            cmd.ExecuteNonQuery();
//                        }
//                    }

//                    Console.WriteLine("Data synchronization completed successfully.");
//                }

//                //Franchise
//                //FinancialRatioManager fr = new FinancialRatioManager();
//                //fr.Get_FinancialRatio_Data();
//                //Console.WriteLine("End DataLoaded");

//                // Load financial ratio data

//                //TBK
//                //Saruf Ratul//
//                FinancialRatioManagerTBK frTBK = new FinancialRatioManagerTBK();
//                frTBK.Get_FinancialRatio_DataTBK();
//                Console.WriteLine("Financial Ratio Data TBK Loaded");

//                //AM
//                //Saruf Ratul//
//                FinancialRatioManagerAM frAM = new FinancialRatioManagerAM();
//                frAM.Get_FinancialRatio_DataAM();
//                Console.WriteLine("Financial Ratio Data AM Loaded");
//            }
//            catch (Exception ex)
//            {
//                Console.WriteLine("An error occurred: " + ex.Message);
//            }
//        }
//    }
//}
