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



namespace FranchisorEXE
{
    class QBODataProcessor
    {
        public static DataTable GetDisplayItemsFromDB(string FranchisorID, string LocatinoID)
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
        public static DataSet LoadGoalData(string FranchisorID, string LocationID)
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
        public static QBOSettings GetQBOSettingsLocationWise(string FranchisorID, string LocationID)
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


                }

                db.Close();

            }
            catch { }

            return qbs;
        }
        public ServiceContext GetServiceContext(string accessToken, string realmId, bool QBOSandBox = false)
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

                //serviceContext.IppConfiguration.MinorVersion.Qbo = "29";
                return serviceContext;
            }
            catch (Exception ex)
            {
                //LogError(ex.Message);
                return null;

            }

        }
        public bool RefreshAccessToken(QBOSettings qboSettings)
        {
            bool retVal = false;

            try
            {
                //retVal = IsTokenValid(qboSettings);

                //if (!retVal)
                //{
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

                    UpdateToken(qboSettings);

                    retVal = true;

                }
                //}

            }
            catch (Exception ex)
            {
                qboSettings.ErrorMessage = "Error: " + ex.Message;
            }

            return retVal;

        }
        private bool IsTokenValid(QBOSettings qboSettings)
        {

            bool retVal = false;

            try
            {

                ServiceContext context = GetServiceContext(qboSettings.AccessToken, qboSettings.FileID, qboSettings.Sandbox);

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
        private void UpdateToken(QBOSettings qboSettings)
        {
            Database db = new Database();
            db.Open();

            string sql = "Update [Franchisor].[dbo].[tbl_Location] set " +
                        " AccessToken = '" + qboSettings.AccessToken + "'," +
                        " RefreshToken = '" + qboSettings.RefreshToken + "'," +
                        " LastQBOSync = '" + DateTime.Now.ToString() + "'" +
                        " Where FranchisorID = '" + qboSettings.FranchisorID + "'" +
                        " And LocationID = '" + qboSettings.LocationID + "'";

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
        public CashFlowInfo GetCashFlowInfo(QBOSettings qbs)
        {
            CashFlowInfo cfInfo = new CashFlowInfo();

            TimeZoneInfo ZoneInfo = TimeZoneInfo.FindSystemTimeZoneById("Pacific Standard Time");
            DateTime dateTime = Convert.ToDateTime(qbs.LastSyncDateTime);
            dateTime = TimeZoneInfo.ConvertTime(dateTime, ZoneInfo);

            string LastUpdatedTime = dateTime.ToString("yyyy-MM-ddTHH:mm:ss");

            ServiceContext context = GetServiceContext(qbs.AccessToken, qbs.FileID, qbs.Sandbox);

            if (qbs.Sandbox)
            {
                context.IppConfiguration.BaseUrl.Qbo = "https://sandbox-quickbooks.api.intuit.com/";
            }

            //Get value from Balance Sheet
            getValueFromBalanceSheet(context, cfInfo);

            try
            {
                string dateStart = DateTime.Today.ToString("yyyy-MM-dd");
                string dateEnd = DateTime.Today.ToString("yyyy-MM-dd");

                dateStart = DateTime.Now.FirstDayOfYear().ToString("yyyy-MM-dd");

                ReportService reportsService = new ReportService(context);
                reportsService.start_date = dateStart;
                reportsService.end_date = dateEnd;

                //Execute Report API call
                Report report = reportsService.ExecuteReport("CashFlow");

                Row[] rows = report.Rows;

                try
                {

                    for (int rowIndex = 0; rowIndex < rows.Length; rowIndex++)
                    {
                        Row row = rows[rowIndex];

                        //Get Row Summary
                        Summary rowSummary = GetRowProperty<Summary>(row, ItemsChoiceType1.Summary);

                        //Append Row Summary
                        if (rowSummary != null && rowSummary.ColData != null)
                        {
                            ColData[] colData = rowSummary.ColData;

                            try
                            {
                                if (colData[0].value.Equals("Cash at end of period", StringComparison.OrdinalIgnoreCase))
                                {
                                    string s = colData[0].value;
                                    cfInfo.CashOnHand = colData[1].value;
                                }
                            }
                            catch (Exception ex)
                            {

                            }

                        }
                    }
                }
                catch (Exception ex)
                {
                }

            }
            catch (Exception ex)
            {

            }

            return cfInfo;

        }
        private T GetRowProperty<T>(Row row, ItemsChoiceType1 itemsChoiceType)
        {
            int choiceElementIndex = GetChoiceElementIndex(row, itemsChoiceType);
            if (choiceElementIndex == -1) { return default(T); } else { return (T)row.AnyIntuitObjects[choiceElementIndex]; }
        }
        private int GetChoiceElementIndex(Row row, ItemsChoiceType1 itemsChoiceType)
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
        private void getValueFromBalanceSheet(ServiceContext context, CashFlowInfo cfInfo)
        {

            try
            {
                string dateStart = DateTime.Now.FirstDayOfYear().ToString("yyyy-MM-dd");
                string dateEnd = DateTime.Today.ToString("yyyy-MM-dd");

                ReportService reportsService = new ReportService(context);
                reportsService.start_date = dateStart;
                reportsService.end_date = dateEnd;

                //Execute Report API call
                Report report = reportsService.ExecuteReport("BalanceSheet");

                Row[] rows = report.Rows;

                try
                {

                    for (int rowIndex = 0; rowIndex < rows.Length; rowIndex++)
                    {
                        Row row = rows[rowIndex];

                        try
                        {
                            if (row.group == "TotalAssets")
                            {
                                Rows rs1 = GetRowProperty<Rows>(row, ItemsChoiceType1.Rows);
                                Row rs1r1 = rs1.Row[0];
                                // start on 5/30/2024 by naz
                                Summary rowSummaryAsset = GetRowProperty<Summary>(rs1r1, ItemsChoiceType1.Summary);
                                cfInfo.TotalCurrentAssets = rowSummaryAsset.ColData[1].value;
                                // end on 5/330/2024 by naz
                                Rows rs2 = GetRowProperty<Rows>(rs1r1, ItemsChoiceType1.Rows);
                                Row rs2r0 = rs2.Row[0];

                                Row rs2r1 = rs2.Row[0];
                                if (rs2.Row.Length > 1)
                                {
                                    rs2r1 = rs2.Row[1];

                                }
                                Summary rowSummary0 = GetRowProperty<Summary>(rs2r0, ItemsChoiceType1.Summary);
                                Summary rowSummary1 = GetRowProperty<Summary>(rs2r1, ItemsChoiceType1.Summary);
                                ColData[] ScolData = rowSummary1.ColData;
                                string bs1 = ScolData[0].value;
                                cfInfo.AccountReceivable = ScolData[1].value;

                                ScolData = rowSummary0.ColData;
                                cfInfo.TotalBankAccounts = ScolData[1].value;
                            }
                        }
                        catch (Exception ex)
                        {

                        }

                        try
                        {
                            if (row.group == "TotalLiabilitiesAndEquity")
                            {
                                Rows rs1 = GetRowProperty<Rows>(row, ItemsChoiceType1.Rows);
                                Row rs1r1 = rs1.Row[0];
                                // start on 5/30/2024 by naz
                                Summary rowSummaryAsset = GetRowProperty<Summary>(rs1r1, ItemsChoiceType1.Summary);
                                cfInfo.TotalCurrentLiabilities = rowSummaryAsset.ColData[1].value;
                                // end on 5/330/2024 by naz
                                Rows rs2 = GetRowProperty<Rows>(rs1r1, ItemsChoiceType1.Rows);
                                Row rs2r1 = rs2.Row[0];
                                Rows rs3 = GetRowProperty<Rows>(rs2r1, ItemsChoiceType1.Rows);
                                Row rs3r1 = rs3.Row[0];
                                Summary rowSummary1 = GetRowProperty<Summary>(rs3r1, ItemsChoiceType1.Summary);
                                ColData[] ScolData = rowSummary1.ColData;
                                string bs1 = ScolData[0].value;
                                cfInfo.AccountPayable = ScolData[1].value;

                                Row rs3r2 = rs3.Row[1];
                                rowSummary1 = GetRowProperty<Summary>(rs3r2, ItemsChoiceType1.Summary);
                                ScolData = rowSummary1.ColData;
                                bs1 = ScolData[0].value;
                                cfInfo.CCBalance = ScolData[1].value;
                            }
                        }
                        catch (Exception ex)
                        {

                        }

                    }
                }
                catch (Exception ex)
                {

                }

            }
            catch (Exception ex)
            {

            }

        }
        public ProfitAndLossInfo GetProfitAndLossInfo(string TimePeriod, QBOSettings qbs)
        {
            ProfitAndLossInfo pnl = new ProfitAndLossInfo();
            if (qbs == null)
            {
                return pnl;
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
                string dateStart = DateTime.Today.ToString("yyyy-MM-dd");
                //string dateEnd = DateTime.Today.AddDays(1).ToString("yyyy-MM-dd");
                string dateEnd = DateTime.Today.ToString("yyyy-MM-dd");

                switch (TimePeriod)
                {
                    case "Today":
                        PopulateProfitAndLossInfo(dateStart, dateEnd, context, pnl, "Today");
                        break;
                    case "Week":
                        dateStart = DateTime.Now.FirstDayOfWeek().ToString("yyyy-MM-dd");
                        PopulateProfitAndLossInfo(dateStart, dateEnd, context, pnl, "Week");
                        break;
                    case "Month":
                        dateStart = DateTime.Now.FirstDayOfMonth().ToString("yyyy-MM-dd");
                        PopulateProfitAndLossInfo(dateStart, dateEnd, context, pnl, "Month");
                        break;
                    case "Year":
                        dateStart = DateTime.Now.FirstDayOfYear().ToString("yyyy-MM-dd");
                        PopulateProfitAndLossInfo(dateStart, dateEnd, context, pnl, "Year");
                        break;

                }

            }
            catch (Exception ex)
            {

            }

            if (string.IsNullOrEmpty(pnl.TodayTotalExpenses))
            {
                pnl.TodayTotalExpenses = "0.00";
            }
            if (string.IsNullOrEmpty(pnl.YearTotalExpenses))
            {
                pnl.YearTotalExpenses = "0.00";
            }
            if (string.IsNullOrEmpty(pnl.MonthTotalExpenses))
            {
                pnl.MonthTotalExpenses = "0.00";
            }
            if (string.IsNullOrEmpty(pnl.WeekTotalExpenses))
            {
                pnl.WeekTotalExpenses = "0.00";
            }
            return pnl;
        }
        private void PopulateProfitAndLossInfo(string startDate, string endDate, ServiceContext context, ProfitAndLossInfo pnl, string period)
        {
            string CostOfGoods = "0.0";
            string GrossProfit = "0.0";
            string CostOfLabor = "0.0";
            string GrossRevenue = "0.0";
            string TotalExpenses = "0.0";
            string TotalIncome = "0.0";

            ReportService reportsService = new ReportService(context);
            reportsService.start_date = startDate;
            reportsService.end_date = endDate;

            //Execute Report API call
            Report report = reportsService.ExecuteReport("ProfitAndLoss");

            Row[] rows = report.Rows;

            try
            {

                for (int rowIndex = 0; rowIndex < rows.Length; rowIndex++)
                {
                    Row row = rows[rowIndex];

                    if (row.group != null)
                    {
                        if (row.group == "Income")
                        {
                            //Get Child Rows
                            Rows childRows = GetRowProperty<Rows>(row, ItemsChoiceType1.Rows);

                            if (childRows != null)
                            {
                                foreach (Row r in childRows.Row)
                                {
                                    if (r.type == RowTypeEnum.Section)
                                    {
                                        Rows r1 = GetRowProperty<Rows>(r, ItemsChoiceType1.Rows);

                                        if (r1 != null)
                                        {
                                            foreach (Row r2 in r1.Row)
                                            {
                                                Summary rSummary = GetRowProperty<Summary>(r2, ItemsChoiceType1.Summary);
                                                if (rSummary != null && rSummary.ColData != null)
                                                {
                                                    ColData[] colData = rSummary.ColData;

                                                    try
                                                    {
                                                        if (colData[0].value == "Total Labor")
                                                        {
                                                            CostOfLabor = colData[1].value;
                                                        }

                                                    }
                                                    catch (Exception ex)
                                                    {

                                                    }

                                                }
                                            }
                                        }
                                    }

                                }
                            }
                        }
                    }


                    //Get Row Summary
                    Summary rowSummary = GetRowProperty<Summary>(row, ItemsChoiceType1.Summary);

                    //Append Row Summary
                    if (rowSummary != null && rowSummary.ColData != null)
                    {
                        ColData[] colData = rowSummary.ColData;

                        try
                        {
                            switch (colData[0].value)
                            {
                                case "Total Cost of Goods Sold":
                                    CostOfGoods = colData[1].value;
                                    break;
                                case "Gross Profit":
                                    GrossProfit = colData[1].value;
                                    break;
                                case "Total Labor":
                                    CostOfLabor = colData[1].value;
                                    break;
                                case "Net Income":
                                    GrossRevenue = colData[1].value;
                                    break;
                                case "Total Expenses":
                                    TotalExpenses = colData[1].value;
                                    break;
                                case "Total Income":
                                    TotalIncome = colData[1].value;
                                    break;
                            }
                        }
                        catch (Exception ex)
                        {

                        }

                    }
                }
            }
            catch (Exception ex)
            {

            }

            switch (period)
            {
                case "Today":
                    pnl.TodayCostOfGoods = CostOfGoods;
                    pnl.TodayCostOfLabor = CostOfLabor;
                    pnl.TodayGrossProfit = GrossProfit;
                    pnl.TodayGrossRevenue = GrossRevenue;
                    pnl.TodayTotalExpenses = TotalExpenses;
                    pnl.TodayTotalIncome = TotalIncome;
                    break;
                case "Week":
                    pnl.WeekCostOfGoods = CostOfGoods;
                    pnl.WeekCostOfLabor = CostOfLabor;
                    pnl.WeekGrossProfit = GrossProfit;
                    pnl.WeekGrossRevenue = GrossRevenue;
                    pnl.WeekTotalExpenses = TotalExpenses;
                    pnl.WeekTotalIncome = TotalIncome;
                    break;
                case "Month":
                    pnl.MonthCostOfGoods = CostOfGoods;
                    pnl.MonthCostOfLabor = CostOfLabor;
                    pnl.MonthGrossProfit = GrossProfit;
                    pnl.MonthGrossRevenue = GrossRevenue;
                    pnl.MonthTotalExpenses = TotalExpenses;
                    pnl.MonthTotalIncome = TotalIncome;
                    break;
                case "Year":
                    pnl.YearCostOfGoods = CostOfGoods;
                    pnl.YearCostOfLabor = CostOfLabor;
                    pnl.YearGrossProfit = GrossProfit;
                    pnl.YearGrossRevenue = GrossRevenue;
                    pnl.YearTotalExpenses = TotalExpenses;
                    pnl.YearTotalIncome = TotalIncome;
                    break;
            }

        }

    }

    public class FinanCialRatioData
    {
        public string RevenuePerEmployee = "0.00";
        public string GrossProfitMargin = "0.00";
        public string ProfitMargin = "0.00";
        public string QuickRatio = "0.00";
        public string GrossProfit = "0.00";
        public string TotalRevenue = "0.00";
        public string Royalties = "0.00";

    }

    
  
    //public class FinanCialRatioDataTBK
    //{
    //    public string RevenuePerEmployee { get; set; }
    //    public string GrossProfitMargin { get; set; }
    //    public string ProfitMargin { get; set; }
    //    public string QuickRatio { get; set; }
    //    public string GrossProfit { get; set; }
    //    public string TotalRevenue { get; set; }
    //    public string Royalties { get; set; }
    //    public string CostOfGoods { get; set; }

    //    // KPI for TBK
    //    public string Sales { get; set; }
    //    public string Expense { get; set; }
    //    public string ProfitMarginPercent { get; set; }

    //    public string LabourEfficiency { get; set; }
    //    public string FoodCostEfficiency { get; set; }
    //    public string EventSuppliesEfficiency { get; set; }
    //    public string PassThru { get; set; }
    //    public string CCFees { get; set; }
    //    public string Payroll { get; set; }
    //    public string FoodCost { get; set; }
    //    public string QBOPassThru { get; set; }
    //    public string QBOCCFee { get; set; }
    //    public string EventSupply { get; set; }

    //}
}
