using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FranchisorEXE
{
    class FinancialRatioManager
    {
        public DataTable GetLocationList()
        {
            DataTable dt = new DataTable();
            string strSQL = @"SELECT  l.FranchisorID,l.LocationID,l.LocationName, fg.[QuickRatioDaily],fg.[RevenuePerEmployeeDaily],fg.[GrossProfitMarginDaily],fg.[ProfitMarginDaily]
                                FROM   Franchisor.dbo.tbl_Financial_Goal as fg Right JOIN
                                            Franchisor.dbo.tbl_Location as l ON fg.LocationID = l.LocationID; ";

            Database db = new Database();
            db.Execute(strSQL, out dt);


            return dt;
        }

        public void Get_FinancialRatio_Data()
        {
            //string FranchisorID = HttpContext.Current.Session["FranchisorID"].ToString();
            string RevenuePerEmployeeDaily = "0";
            //List<FinancialRatioDataCategoryWise> financialRatioData = new List<FinancialRatioDataCategoryWise>();
            string qboFinancialData = "";

            qboFinancialData = " truncate table Franchisor.dbo.tbl_FinancialRatio ; ";

            DataTable dt = GetLocationList();

            string[] Date = { "Today", "Week", "Month", "Year" };

            foreach (DataRow dr2 in dt.Rows)
            {
                string FranchisorID = dr2["FranchisorID"].ToString();

                if (dr2["RevenuePerEmployeeDaily"] != null)
                {
                    RevenuePerEmployeeDaily = dr2["RevenuePerEmployeeDaily"].ToString();

                }

                //FinancialRatioDataCategoryWise financialRatioDataLocationWise = new FinancialRatioDataCategoryWise();

                string LocationID = dr2["LocationID"].ToString();
                string LocationName = dr2["LocationName"].ToString();
                List<FinanCialRatioData> _Result = new List<FinanCialRatioData>();
                FinanCialRatioData _FinanCialRatioData = new FinanCialRatioData();
                _FinanCialRatioData.GrossProfitMargin = "0.00";
                _FinanCialRatioData.ProfitMargin = "0.00";
                _FinanCialRatioData.QuickRatio = "0.00";
                _FinanCialRatioData.RevenuePerEmployee = "0.00";

                _FinanCialRatioData.TotalRevenue = "0.00";
                _FinanCialRatioData.GrossProfit = "0.00";
                _FinanCialRatioData.Royalties = "0.00";

                _Result.Add(_FinanCialRatioData);
                //DataTable dtDisplayItems = new DataTable();
                //dtDisplayItems = QBODataProcessor.GetDisplayItemsFromDB(FranchisorID, LocationID);
                DataSet dataSet = QBODataProcessor.LoadGoalData(FranchisorID, LocationID);

                try
                {
                    for (int i = 0; i < Date.Length; i++)
                    {
                        string TimePeriod = Date[i];

                        try
                        {

                            QBOSettings qBOSettings = QBODataProcessor.GetQBOSettingsLocationWise(FranchisorID, LocationID);

                            if (qBOSettings == null)
                            {
                                continue;
                            }

                            Int32 NumberOfEmployees = Convert.ToInt16(RevenuePerEmployeeDaily);

                            QBODataProcessor qbdp = new QBODataProcessor();
                            CashFlowInfo cfInfo = new CashFlowInfo();
                            ProfitAndLossInfo pnl = new ProfitAndLossInfo();

                            bool bQBOToken = qbdp.RefreshAccessToken(qBOSettings);

                            if (bQBOToken)
                            {
                                cfInfo = qbdp.GetCashFlowInfo(FranchisorID, LocationID, qBOSettings);
                                pnl = qbdp.GetProfitAndLossInfo(TimePeriod, qBOSettings);
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
                            string QuickRatio = "0.00";

                            try
                            {
                                QuickRatio = ((Convert.ToDouble(cfInfo.TotalBankAccounts) + Convert.ToDouble(cfInfo.AccountReceivable)) / Convert.ToDouble(cfInfo.AccountPayable)).ToString("0.00");
                            }
                            catch { }
                            if (QuickRatio == "NaN") QuickRatio = "0.00";
                            if (QuickRatio == "∞") QuickRatio = "0.00";

                            _FinanCialRatioData.QuickRatio = QuickRatio.ToString();

                            string YearRevenuePerEmployee = "0.00";

                            string YearGrossProfitMargin = "0.00";

                            switch (TimePeriod)
                            {
                                case "Today":

                                    string TodayRevenuePerEmployee = "0.00";
                                    string TodayRevenue = "0.00";

                                    try
                                    {
                                        //TodayRevenue = pnl.TodayGrossRevenue;
                                        TodayRevenue = pnl.TodayTotalIncome;

                                        TodayRevenuePerEmployee = (Convert.ToDouble(pnl.TodayTotalIncome) / NumberOfEmployees).ToString("0.00");
                                        if (TodayRevenuePerEmployee == "NaN") TodayRevenuePerEmployee = "0.00";
                                    }
                                    catch { }
                                    _FinanCialRatioData.RevenuePerEmployee = TodayRevenuePerEmployee;
                                    _FinanCialRatioData.TotalRevenue = TodayRevenue;
                                    _FinanCialRatioData.GrossProfit = Math.Round(((Convert.ToDouble(pnl.TodayTotalIncome) - Convert.ToDouble(pnl.TodayCostOfGoods)))).ToString();


                                    string TodayGrossProfitMargin = "0.00";
                                    try
                                    {

                                        //TodayGrossProfitMargin = Math.Round(((Convert.ToDouble(pnl.TodayGrossRevenue) - Convert.ToDouble(pnl.TodayCostOfGoods)) / Convert.ToDouble(pnl.TodayGrossRevenue)) * 100).ToString();
                                        TodayGrossProfitMargin = Math.Round(((Convert.ToDouble(pnl.TodayTotalIncome) - Convert.ToDouble(pnl.TodayCostOfGoods)) / Convert.ToDouble(pnl.TodayTotalIncome)) * 100).ToString();


                                        if (TodayGrossProfitMargin == "NaN") TodayGrossProfitMargin = "0.00";
                                    }
                                    catch { }
                                    _FinanCialRatioData.GrossProfitMargin = TodayGrossProfitMargin;


                                    string TodayProfitMargin = "0.00";
                                    try
                                    {
                                        double NetRevenue = Convert.ToDouble(pnl.TodayNetIncome);
                                        double Revenue = Convert.ToDouble(pnl.TodayTotalIncome);

                                        TodayProfitMargin = Math.Round((NetRevenue / Revenue) * 100).ToString();

                                        if (TodayProfitMargin == "NaN") TodayProfitMargin = "0.00";
                                    }
                                    catch { }
                                    _FinanCialRatioData.ProfitMargin = TodayProfitMargin;

                                    break;
                                case "Week":


                                    string WeekRevenuePerEmployee = "0.00";

                                    try
                                    {
                                        // WeekRevenuePerEmployee = (Convert.ToDouble(pnl.WeekGrossRevenue) / NumberOfEmployees).ToString("0.00"); ;
                                        WeekRevenuePerEmployee = (Convert.ToDouble(pnl.WeekTotalIncome) / NumberOfEmployees).ToString("0.00"); ;

                                        if (WeekRevenuePerEmployee == "NaN") WeekRevenuePerEmployee = "0.00";
                                    }
                                    catch { }
                                    _FinanCialRatioData.RevenuePerEmployee = WeekRevenuePerEmployee;
                                    _FinanCialRatioData.TotalRevenue = pnl.WeekTotalIncome;
                                    _FinanCialRatioData.GrossProfit = Math.Round(((Convert.ToDouble(pnl.WeekTotalIncome) - Convert.ToDouble(pnl.WeekCostOfGoods)))).ToString();

                                    //_FinanCialRatioData.TotalRevenue = pnl.WeekGrossRevenue;
                                    //_FinanCialRatioData.GrossProfit = pnl.WeekGrossRevenue;

                                    string WeekGrossProfitMargin = "0.00";

                                    try
                                    {
                                        //WeekGrossProfitMargin = Math.Round(((Convert.ToDouble(pnl.WeekGrossRevenue) - Convert.ToDouble(pnl.WeekCostOfGoods)) / Convert.ToDouble(pnl.WeekGrossRevenue)) * 100).ToString();
                                        WeekGrossProfitMargin = Math.Round(((Convert.ToDouble(pnl.WeekTotalIncome) - Convert.ToDouble(pnl.WeekCostOfGoods)) / Convert.ToDouble(pnl.WeekTotalIncome)) * 100).ToString();

                                        if (WeekGrossProfitMargin == "NaN") WeekGrossProfitMargin = "0.00";
                                    }
                                    catch { }
                                    _FinanCialRatioData.GrossProfitMargin = WeekGrossProfitMargin;

                                    string WeekProfitMargin = "0.00";

                                    try
                                    {
                                        //double NetRevenue = Convert.ToDouble(pnl.WeekGrossRevenue) - (Convert.ToDouble(pnl.WeekCostOfGoods) + Convert.ToDouble(pnl.WeekTotalExpenses));
                                        //double Revenue = Convert.ToDouble(pnl.WeekGrossRevenue);

                                        double NetRevenue = Convert.ToDouble(pnl.WeekNetIncome);
                                        double Revenue = Convert.ToDouble(pnl.WeekTotalIncome);

                                        WeekProfitMargin = Math.Round((NetRevenue / Revenue) * 100).ToString();

                                        if (WeekProfitMargin == "NaN") WeekProfitMargin = "0.00";

                                    }
                                    catch { }
                                    _FinanCialRatioData.ProfitMargin = WeekProfitMargin;
                                    break;
                                case "Month":
                                    string MonthRevenuePerEmployee = "0.00";

                                    try
                                    {
                                        MonthRevenuePerEmployee = (Convert.ToDouble(pnl.MonthTotalIncome) / NumberOfEmployees).ToString("0.00");

                                        if (MonthRevenuePerEmployee == "NaN") MonthRevenuePerEmployee = "0.00";
                                    }
                                    catch { }
                                    _FinanCialRatioData.RevenuePerEmployee = MonthRevenuePerEmployee;
                                    _FinanCialRatioData.TotalRevenue = pnl.MonthTotalIncome;
                                    _FinanCialRatioData.GrossProfit = Math.Round(((Convert.ToDouble(pnl.MonthTotalIncome) - Convert.ToDouble(pnl.MonthCostOfGoods)))).ToString();


                                    string MonthGrossProfitMargin = "0.00";

                                    try
                                    {
                                        //MonthGrossProfitMargin = Math.Round(((Convert.ToDouble(pnl.MonthGrossRevenue) - Convert.ToDouble(pnl.MonthCostOfGoods)) / Convert.ToDouble(pnl.MonthGrossRevenue)) * 100).ToString();
                                        MonthGrossProfitMargin = Math.Round(((Convert.ToDouble(pnl.MonthTotalIncome) - Convert.ToDouble(pnl.MonthCostOfGoods)) / Convert.ToDouble(pnl.MonthTotalIncome)) * 100).ToString();

                                        if (MonthGrossProfitMargin == "NaN") MonthGrossProfitMargin = "0.00";
                                    }
                                    catch { }
                                    _FinanCialRatioData.GrossProfitMargin = MonthGrossProfitMargin;

                                    string MonthProfitMargin = "0.00";

                                    try
                                    {
                                        //double NetRevenue = Convert.ToDouble(pnl.MonthGrossRevenue) - (Convert.ToDouble(pnl.MonthCostOfGoods) + Convert.ToDouble(pnl.MonthTotalExpenses));
                                        //double Revenue = Convert.ToDouble(pnl.MonthGrossRevenue);

                                        double NetRevenue = Convert.ToDouble(pnl.MonthNetIncome);
                                        double Revenue = Convert.ToDouble(pnl.MonthTotalIncome);

                                        MonthProfitMargin = Math.Round((NetRevenue / Revenue) * 100).ToString();

                                        if (MonthProfitMargin == "NaN") MonthProfitMargin = "0.00";
                                    }
                                    catch { }
                                    _FinanCialRatioData.ProfitMargin = MonthProfitMargin;


                                    break;
                                case "Year":

                                    try
                                    {
                                        //YearGrossProfitMargin = Math.Round(((Convert.ToDouble(pnl.YearGrossRevenue) - Convert.ToDouble(pnl.YearCostOfGoods)) / Convert.ToDouble(pnl.YearGrossRevenue)) * 100).ToString();
                                        YearGrossProfitMargin = Math.Round(((Convert.ToDouble(pnl.YearTotalIncome) - Convert.ToDouble(pnl.YearCostOfGoods)) / Convert.ToDouble(pnl.YearTotalIncome)) * 100).ToString();

                                        //int percentage = (int)Math.Round(((Convert.ToDouble(pnl.YearGrossRevenue) - Convert.ToDouble(pnl.YearCostOfGoods)) / Convert.ToDouble(pnl.YearGrossRevenue)) * 100);

                                        if (YearGrossProfitMargin == "NaN") YearGrossProfitMargin = "0.00";
                                    }
                                    catch { }

                                    //var percentage = Convert.ToInt32(Math.Round( Convert.ToDouble(YearGrossProfitMargin), 0));

                                    _FinanCialRatioData.GrossProfitMargin = YearGrossProfitMargin;
                                    string YearProfitMargin = "0.00";

                                    try
                                    {
                                        //double NetRevenue = Convert.ToDouble(pnl.YearGrossRevenue) - (Convert.ToDouble(pnl.YearCostOfGoods) + Convert.ToDouble(pnl.YearTotalExpenses));
                                        //double Revenue = Convert.ToDouble(pnl.YearGrossRevenue);

                                        double NetRevenue = Convert.ToDouble(pnl.YearNetIncome);
                                        double Revenue = Convert.ToDouble(pnl.YearTotalIncome);

                                        YearProfitMargin = Math.Round((NetRevenue / Revenue) * 100).ToString();

                                        if (YearProfitMargin == "NaN") YearProfitMargin = "0.00";
                                    }
                                    catch { }
                                    _FinanCialRatioData.ProfitMargin = YearProfitMargin;

                                    try
                                    {
                                        YearRevenuePerEmployee = (Convert.ToDouble(pnl.YearTotalIncome) / NumberOfEmployees).ToString("0.00"); ;

                                        if (YearRevenuePerEmployee == "NaN") YearRevenuePerEmployee = "0.00";
                                    }
                                    catch { }
                                    _FinanCialRatioData.RevenuePerEmployee = YearRevenuePerEmployee;
                                    _FinanCialRatioData.TotalRevenue = pnl.YearTotalIncome;
                                    _FinanCialRatioData.GrossProfit = Math.Round(((Convert.ToDouble(pnl.YearTotalIncome) - Convert.ToDouble(pnl.YearCostOfGoods)))).ToString();

                                    break;
                            }


                        }
                        catch { }

                        string result = "INSERT INTO dbo.tbl_FinancialRatio (FranchaisorID, LocationID, Date, QuickRatio, TotalRevenue, RevenuePerEmployee, GrossProfitMargin, ProfitMargin, GrossProfit)" +
                            "VALUES('" + FranchisorID + "','" + LocationID + "','" + TimePeriod + "','" +
                            _FinanCialRatioData.QuickRatio + "','" + _FinanCialRatioData.TotalRevenue + "','" + _FinanCialRatioData.RevenuePerEmployee + "','" + _FinanCialRatioData.GrossProfitMargin + "','" + _FinanCialRatioData.ProfitMargin + "','" + _FinanCialRatioData.GrossProfit + "')";

                        qboFinancialData += result;
                    }


                }
                catch (Exception ex)
                {
                }

            }
            Database db = new Database();
            db.Execute(qboFinancialData);
            //return financialRatioData;

        }
    }
}
