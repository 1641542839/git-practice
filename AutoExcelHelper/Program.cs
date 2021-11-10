using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AutoExcelHelper
{
    class Program
    {
        //test commit
        static void Main(string[] args)
        {
            ExcelGeneration run = new ExcelGeneration();
            run.getExcel();
            sendEmail send = new sendEmail();
            _ = send.SendEmails();
            Console.ReadKey();
        }
    }
}
