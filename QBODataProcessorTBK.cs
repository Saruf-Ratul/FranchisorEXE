using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Intuit.Ipp.Core;
using Intuit.Ipp.Data;
using Intuit.Ipp.DataService;
using Intuit.Ipp.QueryFilter;
using Intuit.Ipp.Security;
using Intuit.Ipp.ReportService;
using Intuit.Ipp.OAuth2PlatformClient;
using System.Web;
using System.IO;
using System.Xml.Linq;
using System.Globalization;
using System.Runtime.Remoting.Contexts;


namespace FranchisorEXE
{
    //Saruf Ratul//
    public class QBODataProcessorTBK
    {

        public static DataTable GetDisplayItemsFromDBTBK(string FranchisorID, string LocatinoID)
        {
            Database db = new Database();
            db.Open();

            DataTable dt;

            string strSQL = "SELECT ki.ItemID,ki.ItemName, " +
                            " ISNULL(ks.DisplayInDB,0)DisplayInDB,ISNULL(ks.DisplayInKPI,0)DisplayInKPI,ISNULL(ks.DisplayInCD,0)DisplayInCD " +
                            " FROM tbl_KPIDisplayItem ki " +
                            " left join Franchisor.dbo.tbl_KPIDisplayLocationWise ks " +
                            " on ki.ItemID = ks.ItemID " +
                            " Where FranchisorID='" + FranchisorID + "'";

            db.Execute(strSQL, out dt);
            if (dt.Rows.Count <= 0)
            {
                string InsertDefault = @"INSERT INTO Franchisor.dbo.tbl_KPIDisplayLocationWise(LocationID,FranchisorID, ItemID, DisplayInDB, DisplayInKPI,DisplayInCD)  VALUES('LocationReplaceid','FrancisorReplacedid', '1', 'False', 'False' ,'True');
                INSERT INTO Franchisor.dbo.tbl_KPIDisplayLocationWise( LocationID,FranchisorID,ItemID, DisplayInDB, DisplayInKPI,DisplayInCD)  VALUES('LocationReplaceid','FrancisorReplacedid', '2', 'False', 'False' ,'True');
                INSERT INTO Franchisor.dbo.tbl_KPIDisplayLocationWise( LocationID,FranchisorID,ItemID, DisplayInDB, DisplayInKPI,DisplayInCD)  VALUES('LocationReplaceid','FrancisorReplacedid', '3', 'False', 'False' ,'True');";
                InsertDefault = InsertDefault.Replace("FrancisorReplacedid", FranchisorID);
                InsertDefault = InsertDefault.Replace("LocationReplacedid", LocatinoID);


                db.Execute(InsertDefault);

                db.Execute(strSQL, out dt);
            }


            db.Close();

            return dt;
        }

        public static DataSet LoadGoalDataTBK(string FranchisorID, string LocationID)
        {
            DataSet dtAppt = new DataSet();
            try
            {

                Database db = new Database();

                db.Init("Sp_GetKpiGoals");
                db.AddParameter("@FranchisorID", FranchisorID, SqlDbType.NVarChar);
                db.AddParameter("@LocationID", LocationID, SqlDbType.NVarChar);

                db.Execute(out dtAppt);

                //Session["KpiGoals"] = dtAppt;

            }
            catch { }

            return dtAppt;
        }

        public FinancialData GetProfitAndLossInfoTBKs(string dateStart,string dateEnd, QBOSettings qbs)
        {
            FinancialData pnl = new FinancialData();
            if (qbs == null)
            {
                return pnl;
            }

            TimeZoneInfo ZoneInfo = TimeZoneInfo.FindSystemTimeZoneById("Pacific Standard Time");
            DateTime dateTime = Convert.ToDateTime(qbs.LastSyncDateTime);
            dateTime = TimeZoneInfo.ConvertTime(dateTime, ZoneInfo);

            string LastUpdatedTime = dateTime.ToString("yyyy-MM-ddTHH:mm:ss");

            ServiceContext context = GetServiceContextTBK(qbs.AccessToken, qbs.FileID, qbs.Sandbox);

            if (qbs.Sandbox)
            {
                context.IppConfiguration.BaseUrl.Qbo = "https://sandbox-quickbooks.api.intuit.com/";
            }

            try
            {
                PopulateProfitAndLossReport(dateStart, dateEnd, context, pnl);

                Common.WriteLog("GetProfitAndLossInfoTBKs from QBO: " + DateTime.Now.ToString());
            }
            catch (Exception ex)
            {
                Common.WriteLog("GetProfitAndLossInfoTBKs from QBO: " + DateTime.Now.ToString()+ "-"+ ex.Message);

            }


            return pnl;
        }

        public   (string FromDate, string ToDate) GetDateRange(string TimePeriod)
        {
            string dateStart = DateTime.Today.ToString("yyyy-MM-dd");
            //string dateEnd = DateTime.Today.AddDays(1).ToString("yyyy-MM-dd");
            string dateEnd = DateTime.Today.ToString("yyyy-MM-dd");

            switch (TimePeriod)
            {
                case "Today":
                    
                    break;
                case "Week":
                    dateStart = DateTime.Now.FirstDayOfWeek().ToString("yyyy-MM-dd");
                    
                    break;
                case "Month":
                    dateStart = DateTime.Now.FirstDayOfMonth().ToString("yyyy-MM-dd");
                   
                    break;
                case "Year":
                    dateStart = DateTime.Now.FirstDayOfYear().ToString("yyyy-MM-dd");
                    
                    break;
                case "Year to last month":
                    dateStart = DateTime.Now.FirstDayOfYear().ToString("yyyy-MM-dd");
                    int month = DateTime.Now.Month - 1;
                    int year = DateTime.Now.Year;

                    if (month == 0)
                    {
                        month = 12;
                        year -= 1;
                    }

                    dateEnd = new DateTime(year, month, DateTime.DaysInMonth(year, month)).ToString("yyyy-MM-dd");
                    break;

                case "Previous year":
                    dateStart = new DateTime(DateTime.Now.Year - 1, 1, 1).ToString("yyyy-MM-dd"); // 1st January of the previous year
                    dateEnd = new DateTime(DateTime.Now.Year - 1, 12, 31).ToString("yyyy-MM-dd");// 31st December of the previous year
                    
                    break;
                case "Previous month":
                    dateStart = new DateTime(DateTime.Now.Year, DateTime.Now.Month, 1).AddMonths(-1).ToString("yyyy-MM-dd"); ; // 1st day of the previous month
                    dateEnd = new DateTime(DateTime.Now.Year, DateTime.Now.Month, 1).AddDays(-1).ToString("yyyy-MM-dd"); ; // Last day of the previous month
                    
                    break;
                case "This year quarter":
                    int currentQuarter = (DateTime.Now.Month - 1) / 3 + 1;
                    dateStart = GetStartOfQuarter(DateTime.Now.Year, currentQuarter).ToString("yyyy-MM-dd"); ;
                    dateEnd = GetEndOfQuarter(DateTime.Now.Year, currentQuarter).ToString("yyyy-MM-dd"); ;
                    
                    break;
                case "Previous year quarter":
                    DateTime lastYear = DateTime.Now.AddYears(-1);
                    int lastYearQuarter = (DateTime.Now.Month - 1) / 3 + 1;
                    dateStart = GetStartOfQuarter(lastYear.Year, lastYearQuarter).ToString("yyyy-MM-dd"); ;
                    dateEnd = GetEndOfQuarter(lastYear.Year, lastYearQuarter).ToString("yyyy-MM-dd"); ;
                     
                    break;
                case "Last year to date":
                    dateStart = new DateTime(DateTime.Now.Year - 1, 1, 1).ToString("yyyy-MM-dd"); ;
                    int previousYear = DateTime.Now.Year - 1;
                    int currentMonth = DateTime.Now.Month;
                    int currentDay = DateTime.Now.Day;

                    // for invalid dates (February 29th in a non-leap year)
                    try
                    {
                        dateEnd = new DateTime(previousYear, currentMonth, currentDay).ToString("yyyy-MM-dd");   // Same day and month of the previous year
                    }
                    catch (ArgumentOutOfRangeException)
                    {
                        // set to last valid date of the month
                        dateEnd = new DateTime(previousYear, currentMonth, DateTime.DaysInMonth(previousYear, currentMonth)).ToString("yyyy-MM-dd");
                    }
                     
                    break;
                 case "Last year to last month":
                    dateStart = new DateTime(DateTime.Now.Year - 1, 1, 1).ToString("yyyy-MM-dd"); ;
                      previousYear = DateTime.Now.Year - 1;
                      currentMonth = DateTime.Now.Month-1;
                      currentDay = DateTime.DaysInMonth(previousYear, currentMonth);

                    // for invalid dates (February 29th in a non-leap year)
                    try
                    {
                        dateEnd = new DateTime(previousYear, currentMonth, currentDay).ToString("yyyy-MM-dd");   // Same day and month of the previous year
                    }
                    catch (ArgumentOutOfRangeException)
                    {
                        // set to last valid date of the month
                        dateEnd = new DateTime(previousYear, currentMonth, DateTime.DaysInMonth(previousYear, currentMonth)).ToString("yyyy-MM-dd");
                    }
                     
                    break;


            }
             return (dateStart,dateEnd);

        }
        //public FinancialData GetProfitAndLossInfoTBKs(string TimePeriod, QBOSettings qbs)
        //{
        //    FinancialData pnl = new FinancialData();
        //    if (qbs == null)
        //    {
        //        return pnl;
        //    }

        //    TimeZoneInfo ZoneInfo = TimeZoneInfo.FindSystemTimeZoneById("Pacific Standard Time");
        //    DateTime dateTime = Convert.ToDateTime(qbs.LastSyncDateTime);
        //    dateTime = TimeZoneInfo.ConvertTime(dateTime, ZoneInfo);

        //    string LastUpdatedTime = dateTime.ToString("yyyy-MM-ddTHH:mm:ss");

        //    ServiceContext context = GetServiceContextTBK(qbs.AccessToken, qbs.FileID, qbs.Sandbox);

        //    if (qbs.Sandbox)
        //    {
        //        context.IppConfiguration.BaseUrl.Qbo = "https://sandbox-quickbooks.api.intuit.com/";
        //    }

        //    try
        //    {
        //        string dateStart = DateTime.Today.ToString("yyyy-MM-dd");
        //        //string dateEnd = DateTime.Today.AddDays(1).ToString("yyyy-MM-dd");
        //        string dateEnd = DateTime.Today.ToString("yyyy-MM-dd");

        //        switch (TimePeriod)
        //        {
        //            case "Today":
        //                PopulateProfitAndLossReport( dateStart, dateEnd, context, pnl );
        //                break;
        //            case "Week":
        //                dateStart = DateTime.Now.FirstDayOfWeek().ToString("yyyy-MM-dd");
        //                PopulateProfitAndLossReport( dateStart, dateEnd, context, pnl);
        //                break;
        //            case "Month":
        //                dateStart = DateTime.Now.FirstDayOfMonth().ToString("yyyy-MM-dd");
        //                PopulateProfitAndLossReport( dateStart, dateEnd, context, pnl);
        //                break;
        //            case "Year":
        //                dateStart = DateTime.Now.FirstDayOfYear().ToString("yyyy-MM-dd");
        //                PopulateProfitAndLossReport( dateStart, dateEnd, context, pnl);
        //                break; 
        //            case "Previous year":
        //                dateStart = new DateTime(DateTime.Now.Year - 1, 1, 1).ToString("yyyy-MM-dd"); // 1st January of the previous year
        //                dateEnd = new DateTime(DateTime.Now.Year - 1, 12, 31).ToString("yyyy-MM-dd");// 31st December of the previous year
        //                PopulateProfitAndLossReport( dateStart, dateEnd, context, pnl);
        //                break;
        //            case "Previous month":
        //                dateStart = new DateTime(DateTime.Now.Year, DateTime.Now.Month, 1).AddMonths(-1).ToString("yyyy-MM-dd"); ; // 1st day of the previous month
        //                dateEnd = new DateTime(DateTime.Now.Year, DateTime.Now.Month, 1).AddDays(-1).ToString("yyyy-MM-dd"); ; // Last day of the previous month
        //                PopulateProfitAndLossReport( dateStart, dateEnd, context, pnl);
        //                break;
        //            case "This year quarter":
        //                int currentQuarter = (DateTime.Now.Month - 1) / 3 + 1;
        //                dateStart = GetStartOfQuarter(DateTime.Now.Year, currentQuarter).ToString("yyyy-MM-dd"); ;
        //                dateEnd = GetEndOfQuarter(DateTime.Now.Year, currentQuarter).ToString("yyyy-MM-dd"); ;
        //                PopulateProfitAndLossReport( dateStart, dateEnd, context, pnl);
        //                break;
        //            case "Previous year quarter":
        //                DateTime lastYear = DateTime.Now.AddYears(-1);
        //                int lastYearQuarter = (DateTime.Now.Month - 1) / 3 + 1;
        //                dateStart = GetStartOfQuarter(lastYear.Year, lastYearQuarter).ToString("yyyy-MM-dd"); ;
        //                dateEnd = GetEndOfQuarter(lastYear.Year, lastYearQuarter).ToString("yyyy-MM-dd"); ;
        //                PopulateProfitAndLossReport(dateStart, dateEnd, context, pnl);  
        //                break;
        //            case "Last year to date":
        //                dateStart = new DateTime(DateTime.Now.Year - 1, 1, 1).ToString("yyyy-MM-dd"); ;
        //                int previousYear = DateTime.Now.Year - 1;
        //                int currentMonth = DateTime.Now.Month;
        //                int currentDay = DateTime.Now.Day;

        //                // for invalid dates (February 29th in a non-leap year)
        //                try
        //                {
        //                    dateEnd = new DateTime(previousYear, currentMonth, currentDay).ToString("yyyy-MM-dd");   // Same day and month of the previous year
        //                }
        //                catch (ArgumentOutOfRangeException)
        //                {
        //                    // set to last valid date of the month
        //                    dateEnd = new DateTime(previousYear, currentMonth, DateTime.DaysInMonth(previousYear, currentMonth)).ToString("yyyy-MM-dd"); 
        //                }
        //                PopulateProfitAndLossReport(dateStart, dateEnd, context, pnl);
        //                break;

        //        }

        //    }
        //    catch (Exception ex)
        //    {

        //    }


        //    return pnl;
        //}
        public static DateTime GetStartOfQuarter(int year, int quarter)
        {
            switch (quarter)
            {
                case 1:
                    return new DateTime(year, 1, 1);
                case 2:
                    return new DateTime(year, 4, 1);
                case 3:
                    return new DateTime(year, 7, 1);
                case 4:
                    return new DateTime(year, 10, 1);
                default:
                    throw new ArgumentException("Invalid quarter");
            }
        }

        public static DateTime GetEndOfQuarter(int year, int quarter)
        {
            switch (quarter)
            {
                case 1:
                    return new DateTime(year, 3, 31);
                case 2:
                    return new DateTime(year, 6, 30);
                case 3:
                    return new DateTime(year, 9, 30);
                case 4:
                    return new DateTime(year, 12, 31);
                default:
                    throw new ArgumentException("Invalid quarter");
            }
        }
    
        public static QBOSettings GetQBOSettingsLocationWiseTBK(string FranchisorID, string LocationID)
        {
            QBOSettings qbs = null;

            try
            {
                Database db = new Database();
                db.Open();

                DataTable dt = new DataTable();

                string sql = "SELECT FranchisorID,LocationID,AccessToken,RefreshToken,BQOFileID " +
                      " FROM Franchisor.dbo.tbl_Location " +
                      " Where FranchisorID='" + FranchisorID + "' and LocationID='" + LocationID + "'";

                db.Execute(sql, out dt);

                if (dt.Rows.Count > 0)
                {
                    DataRow dr = dt.Rows[0];

                    qbs = new QBOSettings();

                    qbs.FranchisorID = dr["FranchisorID"].ToString();
                    qbs.LocationID = dr["LocationID"].ToString();
                    qbs.AccessToken = dr["AccessToken"].ToString();
                    qbs.RefreshToken = dr["RefreshToken"].ToString();
                    qbs.FileID = dr["BQOFileID"].ToString();
                    bool sandbox = Convert.ToBoolean(ConfigurationManager.AppSettings["QBOSandBox"].ToString());
                    qbs.Sandbox = sandbox;
                    // qbs.LastSyncDateTime = Convert.ToDateTime(dr["LastSyncDateTime"]);

                }

                db.Close();

            }
            catch { }

            return qbs;
        }

        public ServiceContext GetServiceContextTBK(string accessToken, string realmId, bool QBOSandBox = false)
        {
            //Get the QBO ServiceContext
            try
            {
                OAuth2RequestValidator oauthValidator = new OAuth2RequestValidator(accessToken);
                ServiceContext serviceContext = new ServiceContext(realmId, IntuitServicesType.QBO, oauthValidator);

                if (QBOSandBox)
                {
                    serviceContext.IppConfiguration.BaseUrl.Qbo = "https://sandbox-quickbooks.api.intuit.com/";
                }
                else
                {
                    serviceContext.IppConfiguration.BaseUrl.Qbo = "https://quickbooks.api.intuit.com/";//prod
                }


                serviceContext.IppConfiguration.MinorVersion.Qbo = "29";
                return serviceContext;
            }
            catch (Exception ex)
            {

                Common.WriteLog("Get Service Context from QBO: " + DateTime.Now.ToString() + "-" + ex.Message);

                return null;

            }



        }

        private T GetRowPropertyTBK<T>(Row row, ItemsChoiceType1 itemsChoiceType)
        {
            int choiceElementIndex = GetChoiceElementIndexTBK(row, itemsChoiceType);
            if (choiceElementIndex == -1) { return default(T); } else { return (T)row.AnyIntuitObjects[choiceElementIndex]; }
        }

        private int GetChoiceElementIndexTBK(Row row, ItemsChoiceType1 itemsChoiceType)
        {
            if (row.ItemsElementName != null)
            {
                for (int itemsChoiceTypeIndex = 0; itemsChoiceTypeIndex < row.ItemsElementName.Count(); itemsChoiceTypeIndex++)
                {
                    if (row.ItemsElementName[itemsChoiceTypeIndex] == itemsChoiceType) { return itemsChoiceTypeIndex; }
                }
            }
            return -1;
        }

        public bool RefreshAccessTokenTBK(QBOSettings qboSettings)
        {
            bool retVal = false;

            try
            {
                retVal = IsTokenValidTBK(qboSettings);

                if (!retVal)
                {
                    string prevAccessToken = qboSettings.AccessToken;
                    string prevRefreshToken = qboSettings.RefreshToken;

                    //Refresh AccessToken

                    string appEnvironment = ConfigurationManager.AppSettings["appEnvironment"];
                    string redirectURI = ConfigurationManager.AppSettings["redirectURI"];
                    string clientID = ConfigurationManager.AppSettings["clientID"];
                    string clientSecret = ConfigurationManager.AppSettings["clientSecret"];

                    if (qboSettings.Sandbox)
                    {
                        appEnvironment = "Sandbox";
                        redirectURI = ConfigurationManager.AppSettings["SBredirectURI"];
                        clientID = ConfigurationManager.AppSettings["SBclientId"];
                        clientSecret = ConfigurationManager.AppSettings["SBclientSecret"];
                    }

                    OAuth2Client client = new OAuth2Client(clientID, clientSecret, redirectURI, appEnvironment);

                    TokenResponse tokenResp = client.RefreshTokenAsync(prevRefreshToken).GetAwaiter().GetResult();

                    if (string.IsNullOrWhiteSpace(tokenResp.AccessToken))
                    {
                        qboSettings.ErrorMessage = "QBO access Error.";
                        string sError = "Error: ";

                        if (!string.IsNullOrWhiteSpace(tokenResp.Error))
                        {
                            qboSettings.ErrorMessage = qboSettings.ErrorMessage + " " + tokenResp.Error;
                            sError = "Error: " + tokenResp.Error;
                        }

                        string retRefreshToken = "";

                        if (!string.IsNullOrWhiteSpace(tokenResp.RefreshToken))
                        {
                            retRefreshToken = tokenResp.RefreshToken;
                        }

                    }
                    else
                    {
                        qboSettings.ErrorMessage = "";
                        qboSettings.AccessToken = tokenResp.AccessToken;
                        qboSettings.RefreshToken = tokenResp.RefreshToken;

                        UpdateTokenTBK(qboSettings);

                        retVal = true;

                    }
                }

            }
            catch (Exception ex)
            {
                qboSettings.ErrorMessage = "Error: " + ex.Message;
            }

            return retVal;

        }

        private bool IsTokenValidTBK(QBOSettings qboSettings)
        {

            bool retVal = false;

            try
            {

                ServiceContext context = GetServiceContextTBK(qboSettings.AccessToken, qboSettings.FileID, qboSettings.Sandbox);

                string qboQuery = "Select * From CompanyInfo";

                QueryService<Intuit.Ipp.Data.CompanyInfo> CompanyQueryService = new QueryService<Intuit.Ipp.Data.CompanyInfo>(context);

                Intuit.Ipp.Data.CompanyInfo QBOcompany = CompanyQueryService.ExecuteIdsQuery(qboQuery).FirstOrDefault<Intuit.Ipp.Data.CompanyInfo>(); ;

                if (QBOcompany != null)
                {
                    retVal = true;
                }

            }
            catch (Intuit.Ipp.Exception.IdsException ex)
            {
                retVal = false;
            }

            return retVal;
        }

        private void UpdateTokenTBK(QBOSettings qboSettings)
        {
            Database db = new Database();
            db.Open();

            string sql = "Update [Franchisor].[dbo].[tbl_Location] set " +
                       " AccessToken = '" + qboSettings.AccessToken + "'," +
                       " RefreshToken = '" + qboSettings.RefreshToken + "'" +
                       " Where FranchisorID = '" + qboSettings.FranchisorID + "' and LocationID='" + qboSettings.LocationID + "'";

            try
            {
                db.Execute(sql);
                db.Close();
            }

            catch (Exception ex)
            {
                qboSettings.ErrorMessage += ex.Message;
            }
        }

        //public FinancialData PopulateProfitAndLossInfoTBK(QBOSettings qbs, string startDate = "", string endDate = "",string period="")
        //{
        //    FinancialData pnl = new FinancialData();
        //    if (qbs == null)
        //    {
        //        return pnl;
        //    }

        //    TimeZoneInfo ZoneInfo = TimeZoneInfo.FindSystemTimeZoneById("Pacific Standard Time");
        //    DateTime dateTime = Convert.ToDateTime(qbs.LastSyncDateTime);
        //    dateTime = TimeZoneInfo.ConvertTime(dateTime, ZoneInfo);

        //    string LastUpdatedTime = dateTime.ToString("yyyy-MM-ddTHH:mm:ss");

        //    ServiceContext context = GetServiceContextTBK(qbs.AccessToken, qbs.FileID, qbs.Sandbox);

        //    if (qbs.Sandbox)
        //    {
        //        context.IppConfiguration.BaseUrl.Qbo = "https://sandbox-quickbooks.api.intuit.com/";
        //    }

        //    try
        //    {
        //        string dateStart = Convert.ToDateTime(startDate).ToString("yyyy-MM-dd");
        //        string dateEnd = Convert.ToDateTime(endDate).ToString("yyyy-MM-dd");
        //        PopulateProfitAndLossReport(dateStart, dateEnd, context, pnl, period);
        //    }
        //    catch (Exception ex)
        //    {

        //    }

           
        //    return pnl;
        //}

        private static T GetRowProperty<T>(Row row, ItemsChoiceType1 itemsChoiceType)
        {
            int choiceElementIndex = GetChoiceElementIndex(row, itemsChoiceType);
            if (choiceElementIndex == -1) { return default(T); } else { return (T)row.AnyIntuitObjects[choiceElementIndex]; }
        }
        private static int GetChoiceElementIndex(Row row, ItemsChoiceType1 itemsChoiceType)
        {
            if (row.ItemsElementName != null)
            {
                for (int itemsChoiceTypeIndex = 0; itemsChoiceTypeIndex < row.ItemsElementName.Count(); itemsChoiceTypeIndex++)
                {
                    if (row.ItemsElementName[itemsChoiceTypeIndex] == itemsChoiceType) { return itemsChoiceTypeIndex; }
                }
            }
            return -1;
        }
        public static void PopulateProfitAndLossReport(string startDate, string endDate, ServiceContext context, FinancialData financialData)
        {
            ReportService reportsService = new ReportService(context);
            reportsService.start_date = startDate;
            reportsService.end_date = endDate;

            // Execute Report API call
            Report report = reportsService.ExecuteReport("ProfitAndLoss");

            Row[] rows = report.Rows;

            try
            {
                foreach (Row row in rows)
                {
                    ProcessRowForProfitAndLossData(row, financialData);

                    // Process child rows recursively
                    ProcessChildRows(row, financialData);
                }
            }
            catch (Exception ex)
            {
                // Handle exceptions
            }
        }

        public static void ProcessRowForProfitAndLossData(Row row, FinancialData financialData)
        {
            // Extract summary data
            Summary rowSummary = GetRowProperty<Summary>(row, ItemsChoiceType1.Summary);
            if (rowSummary != null && rowSummary.ColData != null)
            {
                FetchValuesFromColData(rowSummary.ColData, financialData);
            }

            // Extract data from ColData
            ColData[] colData = GetRowProperty<ColData[]>(row, ItemsChoiceType1.ColData);
            if (colData != null)
            {
                FetchValuesFromColData(colData, financialData);
            }
        }

        public static void ProcessChildRows(Row row, FinancialData financialData)
        {
            Rows childRows = GetRowProperty<Rows>(row, ItemsChoiceType1.Rows);
            if (childRows != null)
            {
                foreach (Row childRow in childRows.Row)
                {
                    ProcessRowForProfitAndLossData(childRow, financialData);

                    // Process grandchild rows if any
                    ProcessChildRows(childRow, financialData);
                }
            }
        }

        public static void FetchValuesFromColData(ColData[] colData, FinancialData financialData)
        {
            foreach (var col in colData)
            {
                string key = col.value.ToLower();
                string value = colData[1].value;

                if (key.Contains("cost of goods sold"))
                {
                    financialData.CostOfGoods = value;
                }
                else if (key.Contains("gross profit"))
                {
                    financialData.GrossProfit = value;
                }
                else if (key.Contains("labor"))
                {
                    financialData.CostOfLabor = value;
                }
                else if (key.Contains("net income"))
                {
                    financialData.NetIncome = value;
                }
                else if (key.Contains("total income"))
                {
                    financialData.TotalIncome = value;
                }

                else if (key.Contains("total expenses"))
                {
                    financialData.TotalExpenses = value;
                }
                else if (key.Contains("total other expenses"))
                {
                    financialData.TotalOtherExpenses = value;
                }
                else if (key.Contains("net event sales"))
                {
                    financialData.NetEventSales = value;
                }
                else if (key.Contains("event supplies"))
                {
                    financialData.EventSupplies = value;
                }
                else if (key.Contains("food costs"))
                {
                    financialData.FoodCosts = value;
                }
                else if (key.Contains("total payroll"))
                {
                    financialData.TotalPayroll = value;
                }
                else if (key.Contains("4100 pass-through gratuities"))
                {
                    financialData.PassThroughGratuities = value;
                }
                else if (key.Contains("credit card processing"))
                {
                    financialData.MerchantCreditCardProcessing = value;
                }
                else if (key.Contains("advertising and marketing"))
                {
                    financialData.TotalAdvertisingAndMarketing = value;
                }
                else if (key.Contains("occupancy expenses"))
                {
                    financialData.TotalOccupancyExpenses = value;
                }
                else if (key.Contains("repairs & maintenance"))
                {
                    financialData.RepairsAndMaintenance = value;
                }
                else if (key.Contains("utilities"))
                {
                    financialData.Utilities = value;
                }
                else if (key.Contains("royalty"))
                {
                    financialData.Royalty = value;
                }
                else if (key.Contains("technology fee"))
                {
                    financialData.TechnologyFee = value;
                }
                else if (key.Contains("brand fund contribution"))
                {
                    financialData.BrandFundContribution = value;
                }
                else if (key.Contains("business insurance"))
                {
                    financialData.BusinessInsurance = value;
                }
                else if (key.Contains("office supplies & subscriptions"))
                {
                    financialData.OfficeSuppliesAndSubscriptions = value;
                }
                else if (key.Contains("other expenses"))
                {
                    financialData.OtherExpenses = value;
                }
                else if (key.Contains("sales tax"))
                {
                    financialData.SalesTax = value;
                }
                else if (key.Contains("health insurance premiums"))
                {
                    financialData.HealthInsurancePremiumsExpenses = value;
                }
                else if (key.Contains("retirement contributions"))
                {
                    financialData.RetirementContributionsExpenses = value;
                }
                else if (key.Contains("net operating income"))
                {
                    financialData.NetOperatingIncome = value;
                }

            }
        }


        





        //  YoY Growth function
        public   string YoYGrowth(string currentValue, string previousValue)
        {
            double current = SafeConvertToDouble(currentValue);
            double previous = SafeConvertToDouble(previousValue);

            if (previous == 0 || double.IsInfinity((current - previous) / previous) || double.IsNaN((current - previous) / previous))
            {
                return "0.00%"; // Return zero growth if calculation is invalid
            }

            double growthRate = ((current - previous) / previous) * 100;
            return growthRate.ToString("N2", CultureInfo.GetCultureInfo("en-US")) + "%";
        }

        public   double SafeConvertToDouble(string value)
        {
            if (double.TryParse(value, NumberStyles.Any, CultureInfo.InvariantCulture, out double result))
            {
                return result;
            }
            return 0.0; // Default to zero if conversion fails
        }
        public static ServiceContext GetServiceContext(string accessToken, string realmId, bool QBOSandBox = false)
        {
            //Get the QBO ServiceContext
            try
            {
                OAuth2RequestValidator oauthValidator = new OAuth2RequestValidator(accessToken);
                ServiceContext serviceContext = new ServiceContext(realmId, IntuitServicesType.QBO, oauthValidator);

                if (serviceContext != null)
                {


                }
                else
                {

                }

                if (QBOSandBox)
                {
                    serviceContext.IppConfiguration.BaseUrl.Qbo = "https://sandbox-quickbooks.api.intuit.com/";
                }
                else
                {
                    serviceContext.IppConfiguration.BaseUrl.Qbo = "https://quickbooks.api.intuit.com/";//prod
                }


                serviceContext.IppConfiguration.MinorVersion.Qbo = "29";
                return serviceContext;
            }
            catch (Exception ex)
            {

                return null;

            }



        }
        public   FinancialData GetFinancialReport(QBOSettings qbs, string startDate = "", string endDate = "")
        {
            FinancialData financialData = new FinancialData();
            if (qbs == null)
            {
                return financialData;
            }

            TimeZoneInfo ZoneInfo = TimeZoneInfo.FindSystemTimeZoneById("Pacific Standard Time");
            DateTime dateTime = Convert.ToDateTime(qbs.LastSyncDateTime);
            dateTime = TimeZoneInfo.ConvertTime(dateTime, ZoneInfo);

            string LastUpdatedTime = dateTime.ToString("yyyy-MM-ddTHH:mm:ss");

            ServiceContext context = GetServiceContext(qbs.AccessToken, qbs.FileID, qbs.Sandbox);

            if (qbs.Sandbox)
            {
                context.IppConfiguration.BaseUrl.Qbo = "https://sandbox-quickbooks.api.intuit.com/";
            }

            try
            {
                string dateStart = Convert.ToDateTime(startDate).ToString("yyyy-MM-dd");
                string dateEnd = Convert.ToDateTime(endDate).ToString("yyyy-MM-dd");

                //Getting report data from Profit and loss
                PopulateProfitAndLossReport(dateStart, dateEnd, context, financialData);
                //financialData.CalculateDerivedMetrics(); // Calculate derived fields
            }
            catch (Exception ex)
            {
                // Handle exceptions
            }




            return financialData;
        }

    }

}
