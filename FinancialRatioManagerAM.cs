using Intuit.Ipp.Core;
using Intuit.Ipp.Data;
using Intuit.Ipp.ReportService;
using Intuit.Ipp.Security;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FranchisorEXE
{
    class FinancialRatioManagerAM
    {
        public static object HttpContext { get; private set; }

        public static string FormatAsCurrency(string value)
        {
            // Try to parse the string to a double
            if (double.TryParse(value, NumberStyles.Any, CultureInfo.InvariantCulture, out double number))
            {
                // Round up the number to the nearest whole value
                double roundedUpValue = Math.Ceiling(number);

                // Format the rounded-up value as US currency without decimals
                return roundedUpValue.ToString("C0", CultureInfo.GetCultureInfo("en-US"))
                                     .Replace("(", "-").Replace(")", "");
            }
            else
            {
                return "$0"; // Default value if parsing fails
            }
        }

        //  function to safely convert strings to double
        public static double SafeConvertToDouble(string value)
        {
            if (double.TryParse(value, NumberStyles.Any, CultureInfo.InvariantCulture, out double result))
            {
                return result;
            }
            return 0.0; // Default to zero if conversion fails
        }

        //  function to calculate percentage safely
        public static string SafeCalculatePercentage(string numerator, string denominator)
        {
            double num = SafeConvertToDouble(numerator);
            double den = SafeConvertToDouble(denominator);

            // Check for invalid division scenarios
            if (den == 0 || double.IsInfinity(num / den) || double.IsNaN(num / den))
            {
                return "0.0%"; // Return zero percentage if division is invalid
            }

            double percentage = (num / den) * 100;

            // Format the percentage with one decimal place
            return percentage.ToString("N1", CultureInfo.GetCultureInfo("en-US")) + "%";
        }

        public void Get_FinancialRatio_DataAM()
        {
            Common.WriteLog("Process start: " + DateTime.Now.ToString());


            string qboFinancialData = "";
            qboFinancialData = "TRUNCATE TABLE Franchisor.dbo.tbl_FinancialRatioAM;";

            QBODataProcessorAM qbdpAM = new QBODataProcessorAM();
            DataTable dt = GetLocationListAM();

            string[] Date = { "Today", "Week", "Month", "Year", "Year to last month", "Previous year", "Previous month","Previous week",
                "This year quarter", "Previous year quarter", "Last year to date", "Last year to last month","MonthWise" };


            foreach (DataRow drLocation in dt.Rows)
            {
                QBOSettings qBOSettings = GetQBOSettingsAM(drLocation);
                bool bQBOToken = qbdpAM.RefreshAccessTokenAM(qBOSettings);

                try
                {
                    string YOYSales = "0";//this will not change no matter what the time period is
                    string YOYIncome = "0";//this will not change no matter what the time period is

                    for (int i = 0; i < Date.Length; i++)
                    {

                        string TimePeriod = Date[i];

                        try
                        {
                            CashFlowInfo cfInfo = new CashFlowInfo();
                            FinancialData pnl = new FinancialData();

                            if (bQBOToken)
                            {
                                string startDate = "";
                                string endDate = "";

                                // Extract the start date and end date from the returned tuple
                                if (TimePeriod != "MonthWise")
                                {
                                    var dateRange = qbdpAM.GetDateRange(TimePeriod);

                                    startDate = dateRange.FromDate;
                                    endDate = dateRange.ToDate;
                                    pnl = qbdpAM.GetProfitAndLossInfoAM(startDate, endDate, qBOSettings);


                                    // Last Year
                                    if (TimePeriod == "Year to last month")
                                    {
                                        FinancialData fdReportLastYear = new FinancialData();
                                        DateTime today = DateTime.Today;
                                        //previous year date to last completed month of previous year
                                        DateTime firstDateLastYear = new DateTime(today.Year - 1, 1, 1);
                                        DateTime lastDateLastYear = new DateTime(today.Year - 1, today.Month - 1, DateTime.DaysInMonth(today.Year - 1, today.Month - 1));
                                        string lastYearFromDate = firstDateLastYear.ToString("yyyy-MM-dd");
                                        string lastYearToDate = lastDateLastYear.ToString("yyyy-MM-dd");

                                        fdReportLastYear = qbdpAM.GetFinancialReport(qBOSettings, lastYearFromDate, lastYearToDate);
                                        YOYSales = qbdpAM.YoYGrowth(pnl.NetEventSales, fdReportLastYear.NetEventSales);
                                        YOYIncome = qbdpAM.YoYGrowth(pnl.NetOperatingIncome, fdReportLastYear.NetOperatingIncome);

                                    }
                                    pnl.NetOperatingIncomeYoYGrowthRate = YOYIncome;
                                    pnl.NetEventSalesYoYGrowthRate = YOYSales;
                                    qboFinancialData += GenerateQuery(pnl, qBOSettings, TimePeriod, false);
                                    qboFinancialData += "Update Franchisor.dbo.tbl_FinancialRatioAM set NetEventSalesYoYGrowthRate='" + YOYSales + "' ,NetOperatingIncomeYoYGrowthRate='" + YOYIncome + "' where LocationID='" + qBOSettings.LocationID + "';";

                                }
                                else
                                {
                                    DateTime today = DateTime.Today;

                                    for (int m = 0; m < 6; m++)
                                    {
                                        DateTime firstDayOfMonth = new DateTime(today.Year, today.Month, 1).AddMonths(-m);
                                        DateTime lastDayOfMonth = new DateTime(firstDayOfMonth.Year, firstDayOfMonth.Month, DateTime.DaysInMonth(firstDayOfMonth.Year, firstDayOfMonth.Month));

                                        startDate = firstDayOfMonth.ToString("yyyy-MM-dd");
                                        endDate = lastDayOfMonth.ToString("yyyy-MM-dd");
                                        pnl = qbdpAM.GetProfitAndLossInfoAM(startDate, endDate, qBOSettings);
                                        pnl.NetOperatingIncomeYoYGrowthRate = YOYIncome;
                                        pnl.NetEventSalesYoYGrowthRate = YOYSales;
                                        qboFinancialData += GenerateQuery(pnl, qBOSettings, startDate, true);
                                    }

                                }

                            }

                            //qboFinancialData += result;
                        }
                        catch (Exception ex)
                        {
                            Common.WriteLog("Error occurred during processing: " + ex.Message);
                        }
                    }
                }
                catch (Exception ex)
                {
                    Common.WriteLog("Error occurred during location processing: " + ex.Message);
                }
            }

            // Execute the complete batch of SQL insert statements
            Database db = new Database();
            db.Execute(qboFinancialData);

            Common.WriteLog("Process completed: " + DateTime.Now.ToString());
        }
        
        public string GenerateQuery(FinancialData pnl, QBOSettings qBOSettings, string TimePeriod, bool IsMonthWise = false)
        {

            string LabourEfficiency = "0.0";
            string FoodCostEfficiency = "0.0";
            string EventSuppliesEfficiency = "0.0";
            string ProfitMarginPercent = "0.0";
            string ProfitMargin = "0.0";
            string PassThruasPercentage = "0.0";
            string CCFeeasPercentage = "0.0";
            
            string CostOfGoods = FormatAsCurrency(pnl.CostOfGoods);
            string GrossProfit = FormatAsCurrency(pnl.GrossProfit);
            string CostOfLabor = FormatAsCurrency(pnl.CostOfLabor);
            string GrossRevenue = FormatAsCurrency(pnl.GrossRevenue);
            string TotalExpenses = FormatAsCurrency(pnl.TotalExpenses);
            string TotalIncome = FormatAsCurrency(pnl.TotalIncome);
            string NetEventSales = FormatAsCurrency(pnl.NetEventSales);
            string FoodCosts = FormatAsCurrency(pnl.FoodCosts);
            //string EventSupplies = FormatAsCurrency(pnl.EventSupplies);
            string TotalOtherExpenses = FormatAsCurrency(pnl.TotalOtherExpenses);
            string NetIncome = FormatAsCurrency(pnl.NetIncome);
            string TotalPayroll = FormatAsCurrency(pnl.TotalPayroll);
            string PassThroughGratuities = FormatAsCurrency(pnl.PassThroughGratuities);
            string MerchantCreditCardProcessing = FormatAsCurrency(pnl.MerchantCreditCardProcessing);
            string TotalAdvertisingAndMarketing = FormatAsCurrency(pnl.TotalAdvertisingAndMarketing);
            string TotalOccupancyExpenses = FormatAsCurrency(pnl.TotalOccupancyExpenses);
            string RepairsAndMaintenance = FormatAsCurrency(pnl.RepairsAndMaintenance);
            string Utilities = FormatAsCurrency(pnl.Utilities);
            string Royalty = FormatAsCurrency(pnl.Royalty);
            string TechnologyFee = FormatAsCurrency(pnl.TechnologyFee);
            string BrandFundContribution = FormatAsCurrency(pnl.BrandFundContribution);
            string BusinessInsurance = FormatAsCurrency(pnl.BusinessInsurance);
            string OfficeSuppliesAndSubscriptions = FormatAsCurrency(pnl.OfficeSuppliesAndSubscriptions);
            string OtherExpenses = FormatAsCurrency(pnl.OtherExpenses);
            string SalesTax = FormatAsCurrency(pnl.SalesTax);
            string HealthInsurancePremiumsExpenses = FormatAsCurrency(pnl.HealthInsurancePremiumsExpenses);
            string RetirementContributionsExpenses = FormatAsCurrency(pnl.RetirementContributionsExpenses);
            string NetOperatingIncome = FormatAsCurrency(pnl.NetOperatingIncome);
            string OperatingExpense = FormatAsCurrency(pnl.NetOperatingIncome);


            // Derived formulas calculation  
            string NetOperatingIncomePercentage = SafeCalculatePercentage(pnl.NetOperatingIncome, pnl.NetEventSales);
            string NetOperatingIncomeYoYGrowthRate = pnl.NetOperatingIncomeYoYGrowthRate;
            string FoodCostsPercentage = SafeCalculatePercentage(pnl.FoodCosts, pnl.NetEventSales);
            string EventSuppliesPercentage = SafeCalculatePercentage(pnl.EventSupplies, pnl.NetEventSales);
            string TotalPayrollPercentage = SafeCalculatePercentage(pnl.TotalPayroll, pnl.NetEventSales);
            string TotalAdvertisingAndMarketingPercentage = SafeCalculatePercentage(pnl.TotalAdvertisingAndMarketing, pnl.NetEventSales);
            string TotalOccupancyExpensesPercentage = SafeCalculatePercentage(pnl.TotalOccupancyExpenses, pnl.NetEventSales);
            string RepairsAndMaintenancePercentage = SafeCalculatePercentage(pnl.RepairsAndMaintenance, pnl.NetEventSales);
            string UtilitiesPercentage = SafeCalculatePercentage(pnl.Utilities, pnl.NetEventSales);
            string RoyaltyPercentage = SafeCalculatePercentage(pnl.Royalty, pnl.NetEventSales);
            string TechnologyFeePercentage = SafeCalculatePercentage(pnl.TechnologyFee, pnl.NetEventSales);
            string BrandFundContributionPercentage = SafeCalculatePercentage(pnl.BrandFundContribution, pnl.NetEventSales);
            string BusinessInsurancePercentage = SafeCalculatePercentage(pnl.BusinessInsurance, pnl.NetEventSales);
            string MerchantCreditCardProcessingPercentage = SafeCalculatePercentage(pnl.MerchantCreditCardProcessing, pnl.NetEventSales);
            string OfficeSuppliesAndSubscriptionsPercentage = SafeCalculatePercentage(pnl.OfficeSuppliesAndSubscriptions, pnl.NetEventSales);
            string OtherExpensesPercentage = SafeCalculatePercentage(pnl.OtherExpenses, pnl.NetEventSales);
            string NetEventSalesYoYGrowthRate = pnl.NetEventSalesYoYGrowthRate;


            // Calculation Part

            double Expense = string.IsNullOrEmpty(pnl?.TotalExpenses) ? 0.00 : (Convert.ToDouble(pnl.TotalExpenses));
            double Income = string.IsNullOrEmpty(pnl?.NetIncome) ? 0.00 : (Convert.ToDouble(pnl.NetIncome));
            double Payroll = string.IsNullOrEmpty(pnl?.TotalPayroll) ? 0.00 : (Convert.ToDouble(pnl.TotalPayroll));
            double sales = string.IsNullOrEmpty(pnl?.NetEventSales) ? 0.00 : (Convert.ToDouble(pnl.NetEventSales));
            double FoodCost = string.IsNullOrEmpty(pnl?.FoodCosts) ? 0.00 : (Convert.ToDouble(pnl.FoodCosts));
            double EventSupplies = string.IsNullOrEmpty(pnl?.EventSupplies) ? 0.00 : (Convert.ToDouble(pnl.EventSupplies));
            double passThru = string.IsNullOrEmpty(pnl?.PassThroughGratuities) ? 0.00 : (Convert.ToDouble(pnl.PassThroughGratuities));
            double ccFee = string.IsNullOrEmpty(pnl?.MerchantCreditCardProcessing) ? 0.00 : (Convert.ToDouble(pnl.MerchantCreditCardProcessing));
            double grossprofit = string.IsNullOrEmpty(pnl?.GrossProfit) ? 0.00 : (Convert.ToDouble(pnl.GrossProfit));
            double cogs = string.IsNullOrEmpty(pnl?.CostOfGoods) ? 0.00 : (Convert.ToDouble(pnl.CostOfGoods));


            // Profit Margin Calculation
            double _NetOperatingIncome = string.IsNullOrEmpty(pnl?.NetOperatingIncome) ? 0.00 : Math.Ceiling(Convert.ToDouble(pnl.NetOperatingIncome));

            //double profitmargin = _NetOperatingIncome;
            // According to JAY Profit Margin % - (Sales 4200 - COGS - Operating Expenses) / Sales 4200 X 100
            //ProfitMargin = profitmargin.ToString("0").Replace("NaN", "0");
            //ProfitMarginPercent = (profitmargin * 100).ToString("0.0") + "%";

            double _NetEventSales = string.IsNullOrEmpty(pnl?.NetEventSales) ? 0.00 : Math.Ceiling(Convert.ToDouble(pnl.NetEventSales));
            double _CostOfGoods = string.IsNullOrEmpty(pnl?.CostOfGoods) ? 0.00 : Math.Ceiling(Convert.ToDouble(pnl.CostOfGoods));
            double _TotalExpenses = string.IsNullOrEmpty(pnl?.TotalExpenses) ? 0.00 : Math.Ceiling(Convert.ToDouble(pnl.TotalExpenses));

            double _Cost = _NetEventSales - _CostOfGoods - _TotalExpenses;
            double _ProfitMarginPercent = (_Cost / sales) * 100;


            //Saruf
            //Net Income ÷ Revenue × 100

            double netIncome = Convert.ToDouble(pnl.NetIncome);
            double revenue = Convert.ToDouble(pnl.TotalIncome);
            double grossProfit = Convert.ToDouble(pnl.GrossProfit);
            double profitMargin = Math.Round((netIncome / revenue) * 100, 2);

            //ProfitMargin = (netIncome / revenue).ToString("0.00");
            //ProfitMarginPercent = (profitMargin).ToString("N", System.Globalization.CultureInfo.GetCultureInfo("en-US")) + "%";
            
            ProfitMargin = Math.Round(netIncome / revenue * 100).ToString("0.00") + "%";

            ProfitMarginPercent = Math.Round((grossProfit / revenue) * 100).ToString("0.00") + "%";

            //?//

            //ProfitMargin = (_Cost / sales).ToString("0.00");
            //ProfitMarginPercent = Income
            //ProfitMarginPercent = _ProfitMarginPercent.ToString("0.0").Replace("∞", "0.0") + "%";

            // Efficiency Calculations
            LabourEfficiency = SafeCalculatePercentage(Payroll.ToString(), sales.ToString()).ToString();
            FoodCostEfficiency = SafeCalculatePercentage(FoodCost.ToString(), sales.ToString());
            EventSuppliesEfficiency = SafeCalculatePercentage(EventSupplies.ToString(), sales.ToString());
            PassThruasPercentage = SafeCalculatePercentage(passThru.ToString(), sales.ToString());
            CCFeeasPercentage = SafeCalculatePercentage(ccFee.ToString(), sales.ToString());


            string result = "INSERT INTO Franchisor.dbo.tbl_FinancialRatioAM (" +
                "NetEventSales, TotalExpenses, ProfitMarginPercent, LabourEfficiency, FoodCostEfficiency, EventSuppliesEfficiency, " +
                "PassThruAsSalePercentage, CCFeesAsSalePercentage, CostOfGoods, Payroll, FoodCost, EventSupply, CCFee, PassThru, " +
                "FranchaisorID, LocationID, TimePeriod, CreatedDate, ProfitMargin, GrossProfit, CostOfLabor, GrossRevenue, TotalIncome, " +
                "FoodCosts, EventSupplies, TotalOtherExpenses, NetIncome, TotalPayroll, PassThroughGratuities, MerchantCreditCardProcessing, " +
                "TotalAdvertisingAndMarketing, TotalOccupancyExpenses, RepairsAndMaintenance, Utilities, TechnologyFee, BrandFundContribution, " +
                "BusinessInsurance, OfficeSuppliesAndSubscriptions, OtherExpenses, SalesTax, HealthInsurancePremiumsExpenses, RetirementContributionsExpenses, " +
                "NetOperatingIncome, OperatingExpense, NetOperatingIncomePercentage, NetOperatingIncomeYoYGrowthRate, TotalPayrollPercentage, " +
                "TotalAdvertisingAndMarketingPercentage, TotalOccupancyExpensesPercentage, RepairsAndMaintenancePercentage, UtilitiesPercentage, " +
                "RoyaltyPercentage, Royalty, TechnologyFeePercentage, BrandFundContributionPercentage, BusinessInsurancePercentage, " +
                "MerchantCreditCardProcessingPercentage, OfficeSuppliesAndSubscriptionsPercentage, OtherExpensesPercentage, NetEventSalesYoYGrowthRate, IsMonthWise)" +
                " VALUES('" +
                NetEventSales.Replace("NaN", "0") + "','" +
                TotalExpenses.Replace("NaN", "0") + "','" +
                ProfitMarginPercent.Replace("NaN", "0.0") + "','" +
                LabourEfficiency.Replace("NaN", "0.0") + "','" +
                FoodCostEfficiency.Replace("NaN", "0.0") + "','" +
                EventSuppliesEfficiency.Replace("NaN", "0.0") + "','" +
                PassThruasPercentage.Replace("NaN", "0.0") + "','" +
                CCFeeasPercentage.Replace("NaN", "0.0") + "','" +
                CostOfGoods.Replace("NaN", "0") + "','" +
                Payroll.ToString("0.00") + "','" +
                FoodCost.ToString("0.00") + "','" +
                EventSupplies.ToString("0.00") + "','" +
                ccFee.ToString("0.00") + "','" +
                passThru.ToString("0.00") + "','" +
                qBOSettings.FranchisorID + "','" +
                qBOSettings.LocationID + "','" +
                TimePeriod + "','" +
                DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "','" +
                ProfitMargin + "','" +
                GrossProfit + "','" +
                CostOfLabor + "','" +
                GrossRevenue + "','" +
                TotalIncome + "','" +
                FoodCosts + "','" +
                EventSupplies + "','" +
                TotalOtherExpenses + "','" +
                NetIncome + "','" +
                TotalPayroll + "','" +
                PassThroughGratuities + "','" +
                MerchantCreditCardProcessing + "','" +
                TotalAdvertisingAndMarketing + "','" +
                TotalOccupancyExpenses + "','" +
                RepairsAndMaintenance + "','" +
                Utilities + "','" +
                TechnologyFee + "','" +
                BrandFundContribution + "','" +
                BusinessInsurance + "','" +
                OfficeSuppliesAndSubscriptions + "','" +
                OtherExpenses + "','" +
                SalesTax + "','" +
                HealthInsurancePremiumsExpenses + "','" +
                RetirementContributionsExpenses + "','" +
                NetOperatingIncome + "','" +
                OperatingExpense + "','" +
                NetOperatingIncomePercentage + "','" +
                NetOperatingIncomeYoYGrowthRate + "','" +
                TotalPayrollPercentage + "','" +
                TotalAdvertisingAndMarketingPercentage + "','" +
                TotalOccupancyExpensesPercentage + "','" +
                RepairsAndMaintenancePercentage + "','" +
                UtilitiesPercentage + "','" +
                RoyaltyPercentage + "','" +
                Royalty + "','" +
                TechnologyFeePercentage + "','" +
                BrandFundContributionPercentage + "','" +
                BusinessInsurancePercentage + "','" +
                MerchantCreditCardProcessingPercentage + "','" +
                OfficeSuppliesAndSubscriptionsPercentage + "','" +
                OtherExpensesPercentage + "','" +
                NetEventSalesYoYGrowthRate + "','" +

                IsMonthWise + "');";
            Common.WriteLog("Generate query: " + DateTime.Now.ToString());
            return result;
        }

        private DataTable GetLocationListAM()
        {
            DataTable dt = new DataTable();
            string strSQL = @"SELECT   l.LocationID, l.LocationName, f.FranchisorType, l.AccessToken, l.RefreshToken, l.BQOFileID,
                                  l.QuickbookCompanyName, l.ConnectionStatus, l.LastQBOSync, l.FranchisorID
                          FROM     tbl_Franchaisor as f INNER JOIN
                                  tbl_Location as l  ON f.FranchisorID = l.FranchisorID where f.FranchisorType <> 'TBK' and l.ConnectionStatus=1  ";

            //strSQL += " and l.LocationID='14206' ;";
            Database db = new Database();
            db.Open();
            db.Execute(strSQL, out dt);
            db.Close();
            return dt;
        }

        public static QBOSettings GetQBOSettingsAM(DataRow dr)
        {
            QBOSettings qbs = new QBOSettings();

            qbs.FranchisorID = dr["FranchisorID"].ToString();
            qbs.LocationID = dr["LocationID"].ToString();
            qbs.AccessToken = dr["AccessToken"].ToString();
            qbs.RefreshToken = dr["RefreshToken"].ToString();
            qbs.FileID = dr["BQOFileID"].ToString();
            qbs.LocationName = dr["LocationName"].ToString();
            qbs.QuickbookCompanyName = dr["QuickbookCompanyName"].ToString();
            bool sandbox = Convert.ToBoolean(ConfigurationManager.AppSettings["QBOSandBox"].ToString());
            qbs.Sandbox = sandbox;
            Common.WriteLog("Get QBO Settings: " + DateTime.Now.ToString());
            return qbs;
        }

    }
}
