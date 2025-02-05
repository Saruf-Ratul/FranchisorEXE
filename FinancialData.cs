using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FranchisorEXE
{
    public class FinancialData
    {
        public string LocationName { get; set; } = "0.0";
        // Raw data fields from the report
        public string CostOfGoods { get; set; } = "0.00";
        public string GrossProfit { get; set; } = "0.00";
        public string CostOfLabor { get; set; } = "0.00";
        public string GrossRevenue { get; set; } = "0.00";
        public string TotalExpenses { get; set; } = "0.00";
        public string TotalIncome { get; set; } = "0.00";
        public string NetEventSales { get; set; } = "0.00";
        public string FoodCosts { get; set; } = "0.00";
        public string EventSupplies { get; set; } = "0.00";
        public string TotalOtherExpenses { get; set; } = "0.00";
        public string NetIncome { get; set; } = "0.00";
        public string TotalPayroll { get; set; } = "0.00";
        public string PassThroughGratuities { get; set; } = "0.00";
        public string MerchantCreditCardProcessing { get; set; } = "0.00";
        public string TotalAdvertisingAndMarketing { get; set; } = "0.00";
        public string TotalOccupancyExpenses { get; set; } = "0.00";
        public string RepairsAndMaintenance { get; set; } = "0.00";
        public string Utilities { get; set; } = "0.00";
        public string Royalty { get; set; } = "0.00";
        public string TechnologyFee { get; set; } = "0.00";
        public string BrandFundContribution { get; set; } = "0.00";
        public string BusinessInsurance { get; set; } = "0.00";
        public string OfficeSuppliesAndSubscriptions { get; set; } = "0.00";
        public string OtherExpenses { get; set; } = "0.00";
        public string SalesTax { get; set; } = "0.00";
        public string HealthInsurancePremiumsExpenses { get; set; } = "0.00";
        public string RetirementContributionsExpenses { get; set; } = "0.00";
        public string OperatingExpense { get; set; } = "0.00";

        // Calculated or derived fields for reporting
        public string NetOperatingIncome { get; set; } = "0.00";
        public string NetOperatingIncomePercentage { get; set; } = "0.00";
        public string NetOperatingIncomeYoYGrowthRate { get; set; } = "0.00";
        public string FoodCostsPercentage { get; set; } = "0.00";
        public string EventSuppliesPercentage { get; set; } = "0.00";
        public string TotalPayrollPercentage { get; set; } = "0.00";
        public string TotalAdvertisingAndMarketingPercentage { get; set; } = "0.00";
        public string TotalOccupancyExpensesPercentage { get; set; } = "0.00";
        public string RepairsAndMaintenancePercentage { get; set; } = "0.00";
        public string UtilitiesPercentage { get; set; } = "0.00";
        public string RoyaltyPercentage { get; set; } = "0.00";
        public string TechnologyFeePercentage { get; set; } = "0.00";
        public string BrandFundContributionPercentage { get; set; } = "0.00";
        public string BusinessInsurancePercentage { get; set; } = "0.00";
        public string MerchantCreditCardProcessingPercentage { get; set; } = "0.00";
        public string OfficeSuppliesAndSubscriptionsPercentage { get; set; } = "0.00";
        public string OtherExpensesPercentage { get; set; } = "0.00";
        public string NetEventSalesYoYGrowthRate { get; set; } = "0.00";

        // LHG specific fields

        // Franchise Fees
        public string RoyaltyFees { get; set; } = "0.00";

        // Indirect Expenses
        public string IndirectVanInsurance { get; set; } = "0.00";
        public string IndirectToolsSupplies { get; set; } = "0.00";

        // Payroll Breakdown
        public string PayrollTaxAndWCB { get; set; } = "0.00";

        // COGS Breakdown
        public string CostOfDisposal { get; set; } = "0.00";
        public string CostOfMaterials { get; set; } = "0.00";

        // Advertising & Marketing Breakdown
        public string AdvertisingAndPromotion { get; set; } = "0.00";
        public string AdvertisingLeadGen { get; set; } = "0.00";

        // Bank & Merchant Fees
        public string BankCharges { get; set; } = "0.00";
        public string CreditCardProcessing { get; set; } = "0.00";

        // General Expenses
        public string VehicleAndGas { get; set; } = "0.00";
        public string DuesAndLicenses { get; set; } = "0.00";
        public string Meals { get; set; } = "0.00";

        public string OccupancyExpenses { get; set; } = "0.00";
        public string HealthInsurancePremiums { get; set; } = "0.00";
        public string RetirementContributions { get; set; } = "0.00";


    }


}
