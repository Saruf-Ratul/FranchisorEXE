using System;
using System.Collections.Generic;
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
            //Console.WriteLine("End DataLoaded");

            //TBK
            //Saruf Ratul//

            //FinancialRatioManagerTBK frTBK = new FinancialRatioManagerTBK();
            //frTBK.Get_FinancialRatio_DataTBK();

            //Console.WriteLine("End DataLoaded");

            //AM
            //Saruf Ratul//
            FinancialRatioManagerAM frAM = new FinancialRatioManagerAM();
            frAM.Get_FinancialRatio_DataAM();
            //Console.WriteLine("End DataLoaded");
        }
    }
}
