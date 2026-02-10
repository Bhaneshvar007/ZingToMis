using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data.OracleClient;
using System.Linq;
using System.Linq.Expressions;
using System.Text;
using System.Threading.Tasks;

namespace ZingToMIS
{
    class Program
    {

        static void Main(string[] args)
        {
            MISdbprocess obj = new MISdbprocess();
           //Console.ReadKey();
            try
           { 
            obj.WriteError("=============================================================================================================");
            obj.WriteError("ZingToMIS: Utility Start At = " + DateTime.Now.ToString("dd-MM-yyyy hh:mm:ss"));
            obj.WriteError("ZingToMIS: Activity Start:");

            obj.ActivityStart();

            obj.WriteError("ZingToMIS: Utility End At = " + DateTime.Now.ToString("dd-MM-yyyy hh:mm:ss"));
            obj.WriteError("ZingToMIS: Utility Run Successfully");
            Console.WriteLine("ZingToMIS: Utility Run Successfully");
            obj.WriteError("=============================================================================================================");
           }
           catch(Exception ex)
           {
                obj.WriteError("Error In Main"+ex.Message.ToString());
            }
        }
    }
    
}
