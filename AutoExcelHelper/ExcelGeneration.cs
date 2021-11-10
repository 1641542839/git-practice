using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace AutoExcelHelper
{
    class ExcelGeneration
    {
        public void getExcel()
        {
            try{
                //get search brand from config 
                String brand = ConfigurationManager.AppSettings["Brand"].ToString();
                if (brand.Contains('/'))
                {
                    String[] brandNames = brand.Split('/');
                    foreach (String item in brandNames)
                    {
                        createExcel(item);
                    }
                }
                else
                {
                    createExcel(brand);
                }
            }
            catch(Exception e)
            {
                MessageBox.Show(e.Message);
            }
           
            
        }


        public void createExcel(String brandName)
        {
            //get sheet name from config
            String sheetName = ConfigurationManager.AppSettings["sheetName"].ToString();
            String[] sheetNames = sheetName.Split('/');



            // create file path
            String path = Directory.GetCurrentDirectory() + "\\" + "excelGeneration";
            // create directory
            System.IO.Directory.CreateDirectory(path);
            var time = DateTime.Now.ToString("yyyyMMdd-hhmmss");

            path += "\\" + brandName + time + ".xls";
            //if (brand.Contains("/"))
            //{
            //    path += "\\" + "MuliBrand" + time + ".xls";
            //}
            //else
            //{
            //    path += "\\" + brand + time + ".xls";
            //}



            //this month
            String cMonth = (DateTime.Now.Month-1).ToString();
            //this year
            String cYear = DateTime.Now.Year.ToString();
            String timeTo = cYear + cMonth;

            String timeFrom = new DateTime(DateTime.Now.Year, 1, 1).ToString("yyyyMM");
            string UnitOfMeasurement = "";


            using (var fs = new FileStream(path, FileMode.Create, FileAccess.Write))
            {
                HSSFWorkbook workbook = new HSSFWorkbook();
                HSSFFont myFont = (HSSFFont)workbook.CreateFont();
                myFont.Boldweight = (short)FontBoldWeight.Bold;
                ISheet sheet;

                // get Supplier from config
                String supplier = ConfigurationManager.AppSettings["Supplier"];

                foreach (string name in sheetNames)
                {
                    if (name.Substring(name.Length - 1) != "$") UnitOfMeasurement = "Sales Unit";
                    else if (name.Substring(name.Length - 1) == "$") UnitOfMeasurement = "Sales Value$";

                    //Microsoft.Office.Interop.Excel.Worksheet sheet = wb.Worksheets.Item[name];
                    try
                    {

                        //sheet = wb.Worksheets[Array.IndexOf(sheetNames, name) + 1];
                        //sheet.Name = name;
                        sheet = workbook.CreateSheet(name);

                        DisplaySht(cYear, brandName, supplier, name, UnitOfMeasurement, timeFrom, timeTo, sheet, myFont);
                        //wb.Sheets.Add();

                    }
                    catch (Exception e)
                    {
                        MessageBox.Show(e.Message);
                    }

                }

                workbook.Write(fs);
                workbook.Close();
            }
        }

        public void DisplaySht(String year, String brand, String supplierCode, String sheetName, String unitOfMeasurement, String timeFrom, String timeTo, ISheet sheet, HSSFFont myFont)
        {
            try
            {
                int modeNo;
                string uomarea = "";
                if (sheetName.ToUpper() == "TOTAL UNITS") { modeNo = 1; uomarea = "TOTAL"; } //TOTAL Units
                else if (sheetName.ToUpper() == "TOTAL $") { modeNo = 2; uomarea = "TOTAL"; } //TOTAL $
                else if ((sheetName.Substring(0, 3).ToUpper() == "ASN") && (sheetName.Substring(0, 5).ToUpper() != "TOTAL") && sheetName.Substring(sheetName.Length - 1).ToUpper() != "$") { modeNo = 3; uomarea = "ASN"; } //ASN Units
                else if ((sheetName.Substring(0, 3).ToUpper() == "ASN") && (sheetName.Substring(0, 5).ToUpper() != "TOTAL") && sheetName.Substring(sheetName.Length - 1).ToUpper() == "$") { modeNo = 4; uomarea = "ASN"; } //ASN $
                else if ((sheetName.Substring(0, 3).ToUpper() == "OTH") && sheetName.Substring(sheetName.Length - 1).ToUpper() != "$") { modeNo = 5; uomarea = "OTH"; }//oth Units
                else if ((sheetName.Substring(0, 3).ToUpper() == "OTH") && sheetName.Substring(sheetName.Length - 1).ToUpper() == "$") { modeNo = 6; uomarea = "OTH"; } //oth Units
                else if ((sheetName.Substring(0, 3).ToUpper() == "SMT") && sheetName.Substring(sheetName.Length - 1).ToUpper() != "$") { modeNo = 7; uomarea = "SMT"; }//SMT Units
                else if ((sheetName.Substring(0, 3).ToUpper() == "SMT") && sheetName.Substring(sheetName.Length - 1).ToUpper() == "$") { modeNo = 8; uomarea = "SMT"; }//SMT Units
                else modeNo = 0;

                //if ((sheetName.Substring(0, 3).ToUpper() == "SMT") && sheetName.Substring(sheetName.Length - 1).ToUpper() != "$") { modeNo = 7; uomarea = "SMT"; }
                //else if ((sheetName.Substring(0, 3).ToUpper() == "SMT") && sheetName.Substring(sheetName.Length - 1).ToUpper() == "$") { modeNo = 8; uomarea = "SMT"; }
                //else modeNo = 0;



                var dataTable = GetSQlTable(brand, supplierCode, modeNo, sheetName, "CY", timeFrom, timeTo);
                var dataTable2 = GetSQlTable(brand, supplierCode, modeNo, sheetName, "LY", timeFrom, timeTo);

                FormatSetup(dataTable, dataTable2, sheetName, year, uomarea + " " + unitOfMeasurement, sheet, myFont);

            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
            }
        }

        public DataTable GetSQlTable(String brand, String supplierCode, int Sqlmode, String sheetName, String yearType, String timeFrom, String timeTo)
        {

            try
            {
                var sql_Conn = SqlEdittor(brand, supplierCode, sheetName, Sqlmode, yearType, timeFrom, timeTo);

                // SQL will be use in sqlhelper
                var dataTable = SqlHelper.ExecSQLSales(sql_Conn);//get the data by using sql 
                return dataTable;
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
                return null;
            }
        }

        public String SqlEdittor(string Brand, string SupplierCode, string SheetName, int sqlmode, string yearType, String TimeFrom, String TimeTo)
        {

            try
            {
                string dbserver = ConfigurationManager.AppSettings["DataSource"].ToString();
                string dbname = ConfigurationManager.AppSettings["Initial Catalog"].ToString();

                //integrated security = true means the current Windows account credentials are used for authentication.
                string connection = string.Format("Data Source={0} ; Initial Catalog={1} ; integrated security = true ", dbserver, dbname);



                // config the searching filed
                string TimeRange;
                string sumTimeRange;
                string sumTimeRangeLY;
                string LastYrStr = TimeFrom.Substring(0, 4);
                int LastYrint = Int32.Parse(LastYrStr) - 1;

                if (yearType == "CY")
                {
                    TimeRange = "db1.yyyymm between '" + TimeFrom + "' and '" + TimeTo + "'";
                    sumTimeRange = "yyyymm between '" + TimeFrom + "' and '" + TimeTo + " '";
                    sumTimeRangeLY = "left(yyyymm,4) = '" + LastYrint.ToString() + "'";
                }
                else//last year
                {
                    TimeFrom = (Int32.Parse(TimeFrom.Substring(0, 4)) - 1).ToString();
                    TimeRange = "db1.Year = '" + LastYrint.ToString() + "'";
                    sumTimeRange = "left(yyyymm,4) = '" + LastYrint.ToString() + "'";
                    sumTimeRangeLY = "left(yyyymm,4) = '" + LastYrint.ToString() + "'-1";

                }

                //brand filter
                string BrandSelect = "";
                if (Brand == "" || Brand == null)
                {
                    BrandSelect = "";
                }
                else
                {
                    //string[] Brands = Brand.Split('/');
                    //for (int i = 0; i < Brands.Length; i++)
                    //{
                    //    Brands[i] = "'" + Brands[i].ToUpper() + "'";
                    //}
                    //string Brandstr = string.Join(",", Brands);
                    BrandSelect = "AND(item.u_mcs_brand in( " + "'"+Brand+"'" + "))";
                }

                //supplier filter
                string SupplierSellect = "";
                if (SupplierCode == "" || SupplierCode == null)
                {
                    SupplierSellect = "";
                }
                else
                {
                    string[] SupplierCodes = SupplierCode.Split('/');
                    for (int i = 0; i < SupplierCodes.Length; i++)
                    {
                        SupplierCodes[i] = "'" + SupplierCodes[i] + "'";
                    }
                    string SupplierCodestr = string.Join(",", SupplierCodes);
                    SupplierSellect = "AND(item.cardcode in( " + SupplierCodestr + "))";
                }

                //get the market
                string mkt = SheetName.Substring(0, 3);

                string Year = TimeFrom.Substring(0, 4);
                string sql = "";
                /***what does sqlmode stand for ****/
                if (sqlmode == 1)
                {
                    sql = @"SELECT stock_code,
                                   itemname,
                                   cardcode [Co.],
                                   cardname [Supplier],
                                   q01      [Jan],
                                   q02      [Feb],
                                   q03      [Mar],
                                   q04      [Apr],
                                   q05      [May],
                                   q06      [Jun],
                                   q07      [Jul],
                                   q08      [Aug],
                                   q09      [Sep],
                                   q10      [Oct],
                                   q11      [Nov],
                                   q12      [Dec],
                                   subtotal
                                   --'',
                                   --onhand   [Onhand],
                                   --onorder
                            FROM   (SELECT stock_code,
                                           itemname,
                                           cardcode,
                                           cardname,
                                           q01,
                                           q02,
                                           q03,
                                           q04,
                                           q05,
                                           q06,
                                           q07,
                                           q08,
                                           q09,
                                           q10,
                                           q11,
                                           q12,
                                           Isnull(q01, 0) + Isnull(q02, 0) + Isnull(q03, 0)
                                           + Isnull( q04, 0) + Isnull( q05, 0) + Isnull( q06, 0)
                                           + Isnull( q07, 0) + Isnull(q08, 0) + Isnull( q09, 0)
                                           + Isnull ( q10, 0) + Isnull ( q11, 0) + Isnull ( q12, 0)
                                           [SUBTOTAL],
                                           ''                                                       [EM],
                                           onhand,
                                           onorder
                                    FROM   (SELECT stock_code,
                                                   year,
                                                   value,
                                                   month,
                                                   onhand,
                                                   onorder,
                                                   itemname,
                                                   cardcode,
                                                   cardname
                                            FROM   (SELECT  stock_code                    [Stock_Code],
                                                           LEFT(yyyymm, 4)        [Year],
                                                           value,
                                                            Tran_Type,
														   ReserveC1,
                                                           col + RIGHT(yyyymm, 2) [month],
                                                           YYYYMM,
                                                            MARKET
                                                    FROM  xCurrent " + @"
                                                           CROSS apply(VALUES ( qty,'Q' ),(Net,'S')) x(value, col))db1
                                                   LEFT JOIN xOMA_Code" + @" item
                                                          ON item.itemcode = db1.stock_code
                                            WHERE  (" + TimeRange + @"  )
                                                   AND stock_code IS NOT NULL
                                                   " + SupplierSellect + @"
                                                   " + BrandSelect + @"
                                                   And (db1.market = 'CLS' or db1.market = 'WTH'  or db1.market = 'SMT' or db1.market = 'ASN' or db1.market = 'FS')
												   And ((Tran_Type !='ARCDT')  OR(left(ReserveC1,3) NOT IN ('CDL','SMS','CAT','SDC','FRC','DCD')))) tbp


                                           PIVOT (Sum(value)
                                                 FOR [month] IN ([q01],
                                                                 [q02],
                                                                 [q03],
                                                                 [q04],
                                                                 [q05],
                                                                 [q06],
                                                                 [q07],
                                                                 [q08],
                                                                 [q09],
                                                                 [q10],
                                                                 [q11],
                                                                 [q12])) pt
                                    UNION ALL
                                            select 
                                            '',
                                            'Total ' + '" + Year + @"',
                                            '',
                                            '',

                                            [01],[02],[03],[04],[05],[06],[07],[08],[09],[10],[11],[12],
	                                            Sum(Isnull([01], 0) + Isnull([02], 0) + Isnull([03], 0)
	                                              + Isnull( [04], 0) + Isnull ( [05], 0) + Isnull( [06], 0)
	                                              + Isnull( [07], 0) + Isnull ([08], 0) + Isnull( [09], 0)
	                                              + Isnull( [10], 0) + Isnull( [11], 0) + Isnull( [12], 0)),
                                            '',
                                            NULL,
                                            NULL
                                            from(
	                                            select qty,right(yyyymm,2) [mm] 
	                                            from xCurrent" + @" 
	                                            join xOMA_Code" + @" item  
	                                            ON item.itemcode = stock_code
	                                            where (" + sumTimeRange + @"  )
                                                AND stock_code IS NOT NULL              
	                                            " + SupplierSellect + @"
	                                            " + BrandSelect + @"
                                                And (market = 'CLS' or market = 'WTH'  or market = 'SMT' or market = 'ASN' or market = 'FS')
	                                            And ((Tran_Type !='ARCDT')  OR(left(ReserveC1,3) NOT IN ('CDL','SMS','CAT','SDC','FRC','DCD')))) as p
                                            pivot (sum(qty) for [mm] in ([01],[02],[03],[04],[05],[06],[07],[08],[09],[10],[11],[12]) ) as piv

                                            group by [01],[02],[03],[04],[05],[06],[07],[08],[09],[10],[11],[12]
                                    UNION ALL
                                    select 
                                            '',
                                            'Total '+ Cast(Cast('" + Year + @"' AS INT)-1 AS VARCHAR(4)),
                                            '',
                                            '',
                                            [01],[02],[03],[04],[05],[06],[07],[08],[09],[10],[11],[12],
	                                            Sum(Isnull([01], 0) + Isnull([02], 0) + Isnull([03], 0)
	                                              + Isnull( [04], 0) + Isnull ( [05], 0) + Isnull( [06], 0)
	                                              + Isnull( [07], 0) + Isnull ([08], 0) + Isnull( [09], 0)
	                                              + Isnull( [10], 0) + Isnull( [11], 0) + Isnull( [12], 0)),
                                            '',
                                            NULL,
                                            NULL
                                            from(
	                                            select qty,right(yyyymm,2) [mm] 
	                                            from xCurrent" + @" 
	                                            join xOMA_Code" + @" item  
	                                            ON item.itemcode = stock_code
	                                            where (" + sumTimeRangeLY + @"  )
                                                AND stock_code IS NOT NULL              
	                                            " + SupplierSellect + @"
	                                            " + BrandSelect + @"
                                                And (market = 'CLS' or market = 'WTH'  or market = 'SMT' or market = 'ASN' or market = 'FS')
	                                            And ((Tran_Type !='ARCDT')  OR(left(ReserveC1,3) NOT IN ('CDL','SMS','CAT','SDC','FRC','DCD')))) as p
                                            pivot (sum(qty) for [mm] in ([01],[02],[03],[04],[05],[06],[07],[08],[09],[10],[11],[12]) ) as piv

                                            group by [01],[02],[03],[04],[05],[06],[07],[08],[09],[10],[11],[12]
                                                                 )db";
                }
                else if (sqlmode == 2)
                {
                    sql = @"SELECT stock_code,
                                   itemname,
                                   cardcode [Co.],
                                   cardname [Supplier],
                                   s01      [Jan],
                                   s02      [Feb],
                                   s03      [Mar],
                                   s04      [Apr],
                                   s05      [May],
                                   s06      [Jun],
                                   s07      [Jul],
                                   s08      [Aug],
                                   s09      [Sep],
                                   s10      [Oct],
                                   s11      [Nov],
                                   s12      [Dec],
                                   subtotal
                                   --'',
                                   --onhand   [Onhand],
                                   --onorder
                            FROM   (SELECT stock_code,
                                           itemname,
                                           cardcode,
                                           cardname,
                                           s01,
                                           s02,
                                           s03,
                                           s04,
                                           s05,
                                           s06,
                                           s07,
                                           s08,
                                           s09,
                                           s10,
                                           s11,
                                           s12,
                                           Isnull(s01, 0) + Isnull(s02, 0) + Isnull(s03, 0)
                                           + Isnull( s04, 0) + Isnull( s05, 0) + Isnull( s06, 0)
                                           + Isnull( s07, 0) + Isnull(s08, 0) + Isnull( s09, 0)
                                           + Isnull ( s10, 0) + Isnull ( s11, 0) + Isnull ( s12, 0)
                                           [SUBTOTAL],
                                           ''                                                       [EM],
                                           onhand,
                                           onorder
                                    FROM   (SELECT stock_code,
                                                   year,
                                                   value,
                                                   month,
                                                   onhand,
                                                   onorder,
                                                   itemname,
                                                   cardcode,
                                                   cardname
                                            FROM   (SELECT  stock_code            [Stock_Code],
                                                           LEFT(yyyymm, 4)        [Year],
                                                           value,
                                                            Tran_Type,
														   ReserveC1,
                                                           col + RIGHT(yyyymm, 2) [month],
                                                            YYYYMM,
                                                            MARKET
                                                    FROM   xCurrent" + @"
                                                           CROSS apply(VALUES ( qty,
                                                                      'Q' ),
                                                                              (Net,
                                                                      'S')) x(value, col))db1
                                                   LEFT JOIN xOMA_Code " + @" item
                                                          ON item.itemcode = db1.stock_code
                                            WHERE  (" + TimeRange + @"  )
                                                   AND stock_code IS NOT NULL
                                                   " + SupplierSellect + @"
                                                   " + BrandSelect + @"
                                                   And (db1.market = 'CLS' or db1.market = 'WTH'  or db1.market = 'SMT' or db1.market = 'ASN' or db1.market = 'FS')
												   And ((Tran_Type !='ARCDT')  OR(left(ReserveC1,3) NOT IN ('CDL','SMS','CAT','SDC','FRC','DCD')))) tbp
                                           PIVOT (Sum(value)
                                                 FOR [month] IN ([s01],
                                                                 [s02],
                                                                 [s03],
                                                                 [s04],
                                                                 [s05],
                                                                 [s06],
                                                                 [s07],
                                                                 [s08],
                                                                 [s09],
                                                                 [s10],
                                                                 [s11],
                                                                 [s12])) pt
                                    UNION ALL
                                            select 
                                            '',
                                            'Total ' + '" + Year + @"',
                                            '',
                                            '',

                                            [01],[02],[03],[04],[05],[06],[07],[08],[09],[10],[11],[12],
	                                            Sum(Isnull([01], 0) + Isnull([02], 0) + Isnull([03], 0)
	                                              + Isnull( [04], 0) + Isnull ( [05], 0) + Isnull( [06], 0)
	                                              + Isnull( [07], 0) + Isnull ([08], 0) + Isnull( [09], 0)
	                                              + Isnull( [10], 0) + Isnull( [11], 0) + Isnull( [12], 0)),
                                            '',
                                            NULL,
                                            NULL
                                            from(
	                                            select Net,right(yyyymm,2) [mm] 
	                                            from xCurrent" + @" 
	                                            join xOMA_Code" + @" item  
	                                            ON item.itemcode = stock_code
	                                            where (" + sumTimeRange + @"  )
                                                AND stock_code IS NOT NULL              
	                                            " + SupplierSellect + @"
	                                            " + BrandSelect + @"
                                                And (market = 'CLS' or market = 'WTH'  or market = 'SMT' or market = 'ASN' or market = 'FS')
	                                            And ((Tran_Type !='ARCDT')  OR(left(ReserveC1,3) NOT IN ('CDL','SMS','CAT','SDC','FRC','DCD')))) as p
                                            pivot (sum(Net) for [mm] in ([01],[02],[03],[04],[05],[06],[07],[08],[09],[10],[11],[12]) ) as piv

                                            group by [01],[02],[03],[04],[05],[06],[07],[08],[09],[10],[11],[12]
                                    UNION ALL
                                    select 
                                            '',
                                            'Total '+ Cast(Cast('" + Year + @"' AS INT)-1 AS VARCHAR(4)),
                                            '',
                                            '',
                                            [01],[02],[03],[04],[05],[06],[07],[08],[09],[10],[11],[12],
	                                            Sum(Isnull([01], 0) + Isnull([02], 0) + Isnull([03], 0)
	                                              + Isnull( [04], 0) + Isnull ( [05], 0) + Isnull( [06], 0)
	                                              + Isnull( [07], 0) + Isnull ([08], 0) + Isnull( [09], 0)
	                                              + Isnull( [10], 0) + Isnull( [11], 0) + Isnull( [12], 0)),
                                            '',
                                            NULL,
                                            NULL
                                            from(
	                                            select Net,right(yyyymm,2) [mm] 
	                                            from xCurrent" + @" 
	                                            join xOMA_Code" + @" item  
	                                            ON item.itemcode = stock_code
	                                            where (" + sumTimeRangeLY + @"  )
                                                AND stock_code IS NOT NULL              
	                                            " + SupplierSellect + @"
	                                            " + BrandSelect + @"
                                                And (market = 'CLS' or market = 'WTH'  or market = 'SMT' or market = 'ASN' or market = 'FS')
	                                            And ((Tran_Type !='ARCDT')  OR(left(ReserveC1,3) NOT IN ('CDL','SMS','CAT','SDC','FRC','DCD')))) as p
                                            pivot (sum(Net) for [mm] in ([01],[02],[03],[04],[05],[06],[07],[08],[09],[10],[11],[12]) ) as piv

                                            group by [01],[02],[03],[04],[05],[06],[07],[08],[09],[10],[11],[12]
                                                                 )db";
                }
                else if (sqlmode == 3)
                {
                    sql = @"SELECT stock_code,
                                   itemname,
                                   cardcode [Co.],
                                   cardname [Supplier],
                                   q01      [Jan],
                                   q02      [Feb],
                                   q03      [Mar],
                                   q04      [Apr],
                                   q05      [May],
                                   q06      [Jun],
                                   q07      [Jul],
                                   q08      [Aug],
                                   q09      [Sep],
                                   q10      [Oct],
                                   q11      [Nov],
                                   q12      [Dec],
                                   subtotal
                                   --'',
                                   --onhand   [Onhand],
                                   --onorder
                            FROM   (SELECT stock_code,
                                           itemname,
                                           cardcode,
                                           cardname,
                                           q01,
                                           q02,
                                           q03,
                                           q04,
                                           q05,
                                           q06,
                                           q07,
                                           q08,
                                           q09,
                                           q10,
                                           q11,
                                           q12,
                                           Isnull(q01, 0) + Isnull(q02, 0) + Isnull(q03, 0)
                                           + Isnull( q04, 0) + Isnull( q05, 0) + Isnull( q06, 0)
                                           + Isnull( q07, 0) + Isnull(q08, 0) + Isnull( q09, 0)
                                           + Isnull ( q10, 0) + Isnull ( q11, 0) + Isnull ( q12, 0)
                                           [SUBTOTAL],
                                           ''                                                       [EM],
                                           onhand,
                                           onorder
                                    FROM   (SELECT stock_code,
                                                   year,
                                                   value,
                                                   month,
                                                   onhand,
                                                   onorder,
                                                   itemname,
                                                   cardcode,
                                                   cardname
                                            FROM   (SELECT  stock_code                    [Stock_Code],
                                                           LEFT(yyyymm, 4)        [Year],
                                                           value,
                                                           market,
                                                            Tran_Type,
														   ReserveC1,
                                                           col + RIGHT(yyyymm, 2) [month],
                                                            YYYYMM
                                                    FROM  xCurrent " + @"
                                                           CROSS apply(VALUES ( qty,'Q' ),(Net,'S')) x(value, col))db1
                                                   LEFT JOIN xOMA_Code" + @" item
                                                          ON item.itemcode = db1.stock_code
                                            WHERE  (" + TimeRange + @"  )
                                                   AND stock_code IS NOT NULL
                                                   " + SupplierSellect + @"
                                                   " + BrandSelect + @"
                                                   And (db1.market = '" + mkt + @"' OR '" + mkt + @"' ='')
												   And ((Tran_Type !='ARCDT')  OR(left(ReserveC1,3) NOT IN ('CDL','SMS','CAT','SDC','FRC','DCD')))) tbp
                                           PIVOT (Sum(value)
                                                 FOR [month] IN ([q01],
                                                                 [q02],
                                                                 [q03],
                                                                 [q04],
                                                                 [q05],
                                                                 [q06],
                                                                 [q07],
                                                                 [q08],
                                                                 [q09],
                                                                 [q10],
                                                                 [q11],
                                                                 [q12])) pt
                                    UNION ALL
                                            select 
                                            '',
                                            'Total ' + '" + Year + @"',
                                            '',
                                            '',

                                            [01],[02],[03],[04],[05],[06],[07],[08],[09],[10],[11],[12],
	                                            Sum(Isnull([01], 0) + Isnull([02], 0) + Isnull([03], 0)
	                                              + Isnull( [04], 0) + Isnull ( [05], 0) + Isnull( [06], 0)
	                                              + Isnull( [07], 0) + Isnull ([08], 0) + Isnull( [09], 0)
	                                              + Isnull( [10], 0) + Isnull( [11], 0) + Isnull( [12], 0)),
                                            '',
                                            NULL,
                                            NULL
                                            from(
	                                            select Qty,right(yyyymm,2) [mm] 
	                                            from xCurrent" + @" 
	                                            join xOMA_Code" + @" item  
	                                            ON item.itemcode = stock_code
	                                            where (" + sumTimeRange + @"  )
                                                AND stock_code IS NOT NULL              
	                                            " + SupplierSellect + @"
	                                            " + BrandSelect + @"
                                                And (market = '" + mkt + @"' OR '" + mkt + @"' ='')
	                                            And ((Tran_Type !='ARCDT')  OR(left(ReserveC1,3) NOT IN ('CDL','SMS','CAT','SDC','FRC','DCD')))) as p
                                            pivot (sum(Qty) for [mm] in ([01],[02],[03],[04],[05],[06],[07],[08],[09],[10],[11],[12]) ) as piv
                                            group by [01],[02],[03],[04],[05],[06],[07],[08],[09],[10],[11],[12]
                                    UNION ALL
                                    select 
                                            '',
                                            'Total '+ Cast(Cast('" + Year + @"' AS INT)-1 AS VARCHAR(4)),
                                            '',
                                            '',
                                            [01],[02],[03],[04],[05],[06],[07],[08],[09],[10],[11],[12],
	                                            Sum(Isnull([01], 0) + Isnull([02], 0) + Isnull([03], 0)
	                                              + Isnull( [04], 0) + Isnull ( [05], 0) + Isnull( [06], 0)
	                                              + Isnull( [07], 0) + Isnull ([08], 0) + Isnull( [09], 0)
	                                              + Isnull( [10], 0) + Isnull( [11], 0) + Isnull( [12], 0)),
                                            '',
                                            NULL,
                                            NULL
                                            from(
	                                            select Qty,right(yyyymm,2) [mm] 
	                                            from xCurrent" + @" 
	                                            join xOMA_Code" + @" item  
	                                            ON item.itemcode = stock_code
	                                            where (" + sumTimeRangeLY + @"  )
                                                AND stock_code IS NOT NULL              
	                                            " + SupplierSellect + @"
	                                            " + BrandSelect + @"
                                                And (market = '" + mkt + @"' OR '" + mkt + @"' ='')
	                                            And ((Tran_Type !='ARCDT')  OR(left(ReserveC1,3) NOT IN ('CDL','SMS','CAT','SDC','FRC','DCD')))) as p
                                            pivot (sum(Qty) for [mm] in ([01],[02],[03],[04],[05],[06],[07],[08],[09],[10],[11],[12]) ) as piv

                                            group by [01],[02],[03],[04],[05],[06],[07],[08],[09],[10],[11],[12]
                                                                )db";
                }
                else if (sqlmode == 4)
                {
                    sql = @"SELECT stock_code,
                                   itemname,
                                   cardcode [Co.],
                                   cardname [Supplier],
                                   s01      [Jan],
                                   s02      [Feb],
                                   s03      [Mar],
                                   s04      [Apr],
                                   s05      [May],
                                   s06      [Jun],
                                   s07      [Jul],
                                   s08      [Aug],
                                   s09      [Sep],
                                   s10      [Oct],
                                   s11      [Nov],
                                   s12      [Dec],
                                   subtotal
                                   --'',
                                   --onhand   [Onhand],
                                   --onorder
                            FROM   (SELECT stock_code,
                                           itemname,
                                           cardcode,
                                           cardname,
                                           s01,
                                           s02,
                                           s03,
                                           s04,
                                           s05,
                                           s06,
                                           s07,
                                           s08,
                                           s09,
                                           s10,
                                           s11,
                                           s12,
                                           Isnull(s01, 0) + Isnull(s02, 0) + Isnull(s03, 0)
                                           + Isnull( s04, 0) + Isnull( s05, 0) + Isnull( s06, 0)
                                           + Isnull( s07, 0) + Isnull(s08, 0) + Isnull( s09, 0)
                                           + Isnull ( s10, 0) + Isnull ( s11, 0) + Isnull ( s12, 0)
                                           [SUBTOTAL],
                                           ''                                                       [EM],
                                           onhand,
                                           onorder
                                    FROM   (SELECT stock_code,
                                                   year,
                                                   value,
                                                   month,
                                                   onhand,
                                                   onorder,
                                                   itemname,
                                                   cardcode,
                                                   cardname
                                            FROM   (SELECT  stock_code                    [Stock_Code],
                                                           LEFT(yyyymm, 4)        [Year],
                                                           value,
                                                           market,
                                                            Tran_Type,
														   ReserveC1,
                                                           col + RIGHT(yyyymm, 2) [month],
                                                            YYYYMM
                                                    FROM  xCurrent " + @"
                                                           CROSS apply(VALUES ( qty,'Q' ),(Net,'S')) x(value, col))db1
                                                   LEFT JOIN xOMA_Code" + @" item
                                                          ON item.itemcode = db1.stock_code
                                            WHERE  (" + TimeRange + @"  )
                                                   AND stock_code IS NOT NULL
                                                   " + SupplierSellect + @"
                                                   " + BrandSelect + @"
                                                   And (db1.market = '" + mkt + @"' OR '" + mkt + @"' ='')
												   And ((Tran_Type !='ARCDT')  OR(left(ReserveC1,3) NOT IN ('CDL','SMS','CAT','SDC','FRC','DCD')))) tbp
                                           PIVOT (Sum(value)
                                                 FOR [month] IN ([s01],
                                                                 [s02],
                                                                 [s03],
                                                                 [s04],
                                                                 [s05],
                                                                 [s06],
                                                                 [s07],
                                                                 [s08],
                                                                 [s09],
                                                                 [s10],
                                                                 [s11],
                                                                 [s12])) pt
                                    UNION ALL
                                            select 
                                            '',
                                            'Total ' + '" + Year + @"',
                                            '',
                                            '',

                                            [01],[02],[03],[04],[05],[06],[07],[08],[09],[10],[11],[12],
	                                            Sum(Isnull([01], 0) + Isnull([02], 0) + Isnull([03], 0)
	                                              + Isnull( [04], 0) + Isnull ( [05], 0) + Isnull( [06], 0)
	                                              + Isnull( [07], 0) + Isnull ([08], 0) + Isnull( [09], 0)
	                                              + Isnull( [10], 0) + Isnull( [11], 0) + Isnull( [12], 0)),
                                            '',
                                            NULL,
                                            NULL
                                            from(
	                                            select Net,right(yyyymm,2) [mm] 
	                                            from xCurrent" + @" 
	                                            join xOMA_Code" + @" item  
	                                            ON item.itemcode = stock_code
	                                            where (" + sumTimeRange + @"  )
                                                AND stock_code IS NOT NULL              
	                                            " + SupplierSellect + @"
	                                            " + BrandSelect + @"
                                                And (market = '" + mkt + @"' OR '" + mkt + @"' ='')
	                                            And ((Tran_Type !='ARCDT')  OR(left(ReserveC1,3) NOT IN ('CDL','SMS','CAT','SDC','FRC','DCD')))) as p
                                            pivot (sum(Net) for [mm] in ([01],[02],[03],[04],[05],[06],[07],[08],[09],[10],[11],[12]) ) as piv
                                            group by [01],[02],[03],[04],[05],[06],[07],[08],[09],[10],[11],[12]
                                    UNION ALL
                                    select 
                                            '',
                                            'Total '+ Cast(Cast('" + Year + @"' AS INT)-1 AS VARCHAR(4)),
                                            '',
                                            '',
                                            [01],[02],[03],[04],[05],[06],[07],[08],[09],[10],[11],[12],
	                                            Sum(Isnull([01], 0) + Isnull([02], 0) + Isnull([03], 0)
	                                              + Isnull( [04], 0) + Isnull ( [05], 0) + Isnull( [06], 0)
	                                              + Isnull( [07], 0) + Isnull ([08], 0) + Isnull( [09], 0)
	                                              + Isnull( [10], 0) + Isnull( [11], 0) + Isnull( [12], 0)),
                                            '',
                                            NULL,
                                            NULL
                                            from(
	                                            select Net,right(yyyymm,2) [mm] 
	                                            from xCurrent" + @" 
	                                            join xOMA_Code" + @" item  
	                                            ON item.itemcode = stock_code
	                                            where (" + sumTimeRangeLY + @" )
                                                AND stock_code IS NOT NULL              
	                                            " + SupplierSellect + @"
	                                            " + BrandSelect + @"
                                                And (market = '" + mkt + @"' OR '" + mkt + @"' ='')
	                                            And ((Tran_Type !='ARCDT')  OR(left(ReserveC1,3) NOT IN ('CDL','SMS','CAT','SDC','FRC','DCD')))) as p
                                            pivot (sum(Net) for [mm] in ([01],[02],[03],[04],[05],[06],[07],[08],[09],[10],[11],[12]) ) as piv

                                            group by [01],[02],[03],[04],[05],[06],[07],[08],[09],[10],[11],[12]
                                                                    )db";
                }
                else if (sqlmode == 5)
                {

                    sql = @"SELECT stock_code,
                                   itemname,
                                   cardcode [Co.],
                                   cardname [Supplier],
                                   q01      [Jan],
                                   q02      [Feb],
                                   q03      [Mar],
                                   q04      [Apr],
                                   q05      [May],
                                   q06      [Jun],
                                   q07      [Jul],
                                   q08      [Aug],
                                   q09      [Sep],
                                   q10      [Oct],
                                   q11      [Nov],
                                   q12      [Dec],
                                   subtotal
                                   --'',
                                   --onhand   [Onhand],
                                   --onorder
                            FROM   (SELECT stock_code,
                                           itemname,
                                           cardcode,
                                           cardname,
                                           q01,
                                           q02,
                                           q03,
                                           q04,
                                           q05,
                                           q06,
                                           q07,
                                           q08,
                                           q09,
                                           q10,
                                           q11,
                                           q12,
                                           Isnull(q01, 0) + Isnull(q02, 0) + Isnull(q03, 0)
                                           + Isnull( q04, 0) + Isnull( q05, 0) + Isnull( q06, 0)
                                           + Isnull( q07, 0) + Isnull(q08, 0) + Isnull( q09, 0)
                                           + Isnull ( q10, 0) + Isnull ( q11, 0) + Isnull ( q12, 0)
                                           [SUBTOTAL],
                                           ''                                                       [EM],
                                           onhand,
                                           onorder
                                    FROM   (SELECT stock_code,
                                                   year,
                                                   value,
                                                   month,
                                                   onhand,
                                                   onorder,
                                                   itemname,
                                                   cardcode,
                                                   cardname
                                            FROM   (SELECT  stock_code                    [Stock_Code],
                                                           LEFT(yyyymm, 4)        [Year],
                                                           value,
                                                           market,
                                                            Tran_Type,
														   ReserveC1,
                                                           col + RIGHT(yyyymm, 2) [month],YYYYMM
                                                    FROM  xCurrent " + @"
                                                           CROSS apply(VALUES ( qty,'Q' ),(Net,'S')) x(value, col))db1
                                                   LEFT JOIN xOMA_Code" + @" item
                                                          ON item.itemcode = db1.stock_code
                                            WHERE  (" + TimeRange + @"  )
                                                   AND stock_code IS NOT NULL
                                                   " + SupplierSellect + @"
                                                   " + BrandSelect + @"
                                                   And (db1.market = 'FS')
												   And ((Tran_Type !='ARCDT')  OR(left(ReserveC1,3) NOT IN ('CDL','SMS','CAT','SDC','FRC','DCD')))) tbp

                                           PIVOT (Sum(value)
                                                 FOR [month] IN ([q01],
                                                                 [q02],
                                                                 [q03],
                                                                 [q04],
                                                                 [q05],
                                                                 [q06],
                                                                 [q07],
                                                                 [q08],
                                                                 [q09],
                                                                 [q10],
                                                                 [q11],
                                                                 [q12])) pt
                                    UNION ALL
                                            select 
                                            '',
                                            'Total ' + '" + Year + @"',
                                            '',
                                            '',

                                            [01],[02],[03],[04],[05],[06],[07],[08],[09],[10],[11],[12],
	                                            Sum(Isnull([01], 0) + Isnull([02], 0) + Isnull([03], 0)
	                                              + Isnull( [04], 0) + Isnull ( [05], 0) + Isnull( [06], 0)
	                                              + Isnull( [07], 0) + Isnull ([08], 0) + Isnull( [09], 0)
	                                              + Isnull( [10], 0) + Isnull( [11], 0) + Isnull( [12], 0)),
                                            '',
                                            NULL,
                                            NULL
                                            from(
	                                            select Qty,right(yyyymm,2) [mm] 
	                                            from xCurrent" + @" 
	                                            join xOMA_Code" + @" item  
	                                            ON item.itemcode = stock_code
	                                            where (" + sumTimeRange + @"  )
                                                AND stock_code IS NOT NULL              
	                                            " + SupplierSellect + @"
	                                            " + BrandSelect + @"
                                                And (market = 'FS')
	                                            And ((Tran_Type !='ARCDT')  OR(left(ReserveC1,3) NOT IN ('CDL','SMS','CAT','SDC','FRC','DCD')))) as p
                                            pivot (sum(Qty) for [mm] in ([01],[02],[03],[04],[05],[06],[07],[08],[09],[10],[11],[12]) ) as piv
                                            group by [01],[02],[03],[04],[05],[06],[07],[08],[09],[10],[11],[12]

                                    UNION ALL
                                    select 
                                            '',
                                            'Total '+ Cast(Cast('" + Year + @"' AS INT)-1 AS VARCHAR(4)),
                                            '',
                                            '',
                                            [01],[02],[03],[04],[05],[06],[07],[08],[09],[10],[11],[12],
	                                            Sum(Isnull([01], 0) + Isnull([02], 0) + Isnull([03], 0)
	                                              + Isnull( [04], 0) + Isnull ( [05], 0) + Isnull( [06], 0)
	                                              + Isnull( [07], 0) + Isnull ([08], 0) + Isnull( [09], 0)
	                                              + Isnull( [10], 0) + Isnull( [11], 0) + Isnull( [12], 0)),
                                            '',
                                            NULL,
                                            NULL
                                            from(
	                                            select Qty,right(yyyymm,2) [mm] 
	                                            from xCurrent" + @" 
	                                            join xOMA_Code" + @" item  
	                                            ON item.itemcode = stock_code
	                                            where (" + sumTimeRangeLY + @"  )
                                                AND stock_code IS NOT NULL              
	                                            " + SupplierSellect + @"
	                                            " + BrandSelect + @"
                                                And (market = 'FS')
	                                            And ((Tran_Type !='ARCDT')  OR(left(ReserveC1,3) NOT IN ('CDL','SMS','CAT','SDC','FRC','DCD')))) as p
                                            pivot (sum(Qty) for [mm] in ([01],[02],[03],[04],[05],[06],[07],[08],[09],[10],[11],[12]) ) as piv

                                            group by [01],[02],[03],[04],[05],[06],[07],[08],[09],[10],[11],[12]
                                                    )db";
                }
                else if (sqlmode == 6)
                {
                    sql = @"SELECT stock_code,
                                   itemname,
                                   cardcode [Co.],
                                   cardname [Supplier],
                                   s01      [Jan],
                                   s02      [Feb],
                                   s03      [Mar],
                                   s04      [Apr],
                                   s05      [May],
                                   s06      [Jun],
                                   s07      [Jul],
                                   s08      [Aug],
                                   s09      [Sep],
                                   s10      [Oct],
                                   s11      [Nov],
                                   s12      [Dec],
                                   subtotal
                                   --'',
                                   --onhand   [Onhand],
                                   --onorder
                            FROM   (SELECT stock_code,
                                           itemname,
                                           cardcode,
                                           cardname,
                                           s01,
                                           s02,
                                           s03,
                                           s04,
                                           s05,
                                           s06,
                                           s07,
                                           s08,
                                           s09,
                                           s10,
                                           s11,
                                           s12,
                                           Isnull(s01, 0) + Isnull(s02, 0) + Isnull(s03, 0)
                                           + Isnull( s04, 0) + Isnull( s05, 0) + Isnull( s06, 0)
                                           + Isnull( s07, 0) + Isnull(s08, 0) + Isnull( s09, 0)
                                           + Isnull ( s10, 0) + Isnull ( s11, 0) + Isnull ( s12, 0)
                                           [SUBTOTAL],
                                           ''                                                       [EM],
                                           onhand,
                                           onorder
                                    FROM   (SELECT stock_code,
                                                   year,
                                                   value,
                                                   month,
                                                   onhand,
                                                   onorder,
                                                   itemname,
                                                   cardcode,
                                                   cardname
                                            FROM   (SELECT  stock_code                    [Stock_Code],
                                                           LEFT(yyyymm, 4)        [Year],
                                                           value,
                                                           market,
                                                            Tran_Type,
														   ReserveC1,
                                                           col + RIGHT(yyyymm, 2) [month],YYYYMM
                                                    FROM  xCurrent " + @"
                                                           CROSS apply(VALUES ( qty,'Q' ),(Net,'S')) x(value, col))db1
                                                   LEFT JOIN xOMA_Code" + @" item
                                                          ON item.itemcode = db1.stock_code
                                            WHERE  (" + TimeRange + @"  )
                                                   AND stock_code IS NOT NULL
                                                   " + SupplierSellect + @"
                                                   " + BrandSelect + @"
                                                   And (db1.market = 'FS')
												   And ((Tran_Type !='ARCDT')  OR(left(ReserveC1,3) NOT IN ('CDL','SMS','CAT','SDC','FRC','DCD')))) tbp

                                           PIVOT (Sum(value)
                                                 FOR [month] IN ([s01],
                                                                 [s02],
                                                                 [s03],
                                                                 [s04],
                                                                 [s05],
                                                                 [s06],
                                                                 [s07],
                                                                 [s08],
                                                                 [s09],
                                                                 [s10],
                                                                 [s11],
                                                                 [s12])) pt
                                    UNION ALL
                                            select 
                                            '',
                                            'Total ' + '" + Year + @"',
                                            '',
                                            '',

                                            [01],[02],[03],[04],[05],[06],[07],[08],[09],[10],[11],[12],
	                                            Sum(Isnull([01], 0) + Isnull([02], 0) + Isnull([03], 0)
	                                              + Isnull( [04], 0) + Isnull ( [05], 0) + Isnull( [06], 0)
	                                              + Isnull( [07], 0) + Isnull ([08], 0) + Isnull( [09], 0)
	                                              + Isnull( [10], 0) + Isnull( [11], 0) + Isnull( [12], 0)),
                                            '',
                                            NULL,
                                            NULL
                                            from(
	                                            select Net,right(yyyymm,2) [mm] 
	                                            from xCurrent" + @" 
	                                            join xOMA_Code" + @" item  
	                                            ON item.itemcode = stock_code
	                                            where (" + sumTimeRange + @"  )
                                                AND stock_code IS NOT NULL              
	                                            " + SupplierSellect + @"
	                                            " + BrandSelect + @"
                                                And (market = 'FS')
	                                            And ((Tran_Type !='ARCDT')  OR(left(ReserveC1,3) NOT IN ('CDL','SMS','CAT','SDC','FRC','DCD')))) as p
                                            pivot (sum(Net) for [mm] in ([01],[02],[03],[04],[05],[06],[07],[08],[09],[10],[11],[12]) ) as piv
                                            group by [01],[02],[03],[04],[05],[06],[07],[08],[09],[10],[11],[12]

                                    UNION ALL
                                    select 
                                            '',
                                            'Total '+ Cast(Cast('" + Year + @"' AS INT)-1 AS VARCHAR(4)),
                                            '',
                                            '',
                                            [01],[02],[03],[04],[05],[06],[07],[08],[09],[10],[11],[12],
	                                            Sum(Isnull([01], 0) + Isnull([02], 0) + Isnull([03], 0)
	                                              + Isnull( [04], 0) + Isnull ( [05], 0) + Isnull( [06], 0)
	                                              + Isnull( [07], 0) + Isnull ([08], 0) + Isnull( [09], 0)
	                                              + Isnull( [10], 0) + Isnull( [11], 0) + Isnull( [12], 0)),
                                            '',
                                            NULL,
                                            NULL
                                            from(
	                                            select Net,right(yyyymm,2) [mm] 
	                                            from xCurrent " + @" 
	                                            join xOMA_Code" + @" item  
	                                            ON item.itemcode = stock_code
	                                            where (" + sumTimeRangeLY + @"  )
                                                AND stock_code IS NOT NULL              
	                                            " + SupplierSellect + @"
	                                            " + BrandSelect + @"
                                                And (market = 'FS')
	                                            And ((Tran_Type !='ARCDT')  OR(left(ReserveC1,3) NOT IN ('CDL','SMS','CAT','SDC','FRC','DCD')))) as p
                                            pivot (sum(Net) for [mm] in ([01],[02],[03],[04],[05],[06],[07],[08],[09],[10],[11],[12]) ) as piv

                                            group by [01],[02],[03],[04],[05],[06],[07],[08],[09],[10],[11],[12]
                                                                 )db";
                }
                else if (sqlmode == 7)
                {

                    sql = @"SELECT stock_code,
                                   itemname,
                                   cardcode [Co.],
                                   cardname [Supplier],
                                   q01      [Jan],
                                   q02      [Feb],
                                   q03      [Mar],
                                   q04      [Apr],
                                   q05      [May],
                                   q06      [Jun],
                                   q07      [Jul],
                                   q08      [Aug],
                                   q09      [Sep],
                                   q10      [Oct],
                                   q11      [Nov],
                                   q12      [Dec],
                                   subtotal
                                   --'',
                                   --onhand   [Onhand],
                                   --onorder
                            FROM   (SELECT stock_code,
                                           itemname,
                                           cardcode,
                                           cardname,
                                           q01,
                                           q02,
                                           q03,
                                           q04,
                                           q05,
                                           q06,
                                           q07,
                                           q08,
                                           q09,
                                           q10,
                                           q11,
                                           q12,
                                           Isnull(q01, 0) + Isnull(q02, 0) + Isnull(q03, 0)
                                           + Isnull( q04, 0) + Isnull( q05, 0) + Isnull( q06, 0)
                                           + Isnull( q07, 0) + Isnull(q08, 0) + Isnull( q09, 0)
                                           + Isnull ( q10, 0) + Isnull ( q11, 0) + Isnull ( q12, 0)
                                           [SUBTOTAL],
                                           ''                                                       [EM],
                                           onhand,
                                           onorder
                                    FROM   (SELECT stock_code,
                                                   year,
                                                   value,
                                                   month,
                                                   onhand,
                                                   onorder,
                                                   itemname,
                                                   cardcode,
                                                   cardname
                                            FROM   (SELECT  stock_code                    [Stock_Code],
                                                           LEFT(yyyymm, 4)        [Year],
                                                           value,
                                                           market,
                                                            Tran_Type,
														   ReserveC1,
                                                           col + RIGHT(yyyymm, 2) [month],YYYYMM
                                                    FROM  xCurrent " + @"
                                                           CROSS apply(VALUES ( qty,'Q' ),(Net,'S')) x(value, col))db1
                                                   LEFT JOIN xOMA_Code" + @" item
                                                          ON item.itemcode = db1.stock_code
                                            WHERE  (" + TimeRange + @"  )
                                                   AND stock_code IS NOT NULL
                                                   " + SupplierSellect + @"
                                                   " + BrandSelect + @"
                                                   And (db1.market = 'CLS' or db1.market = 'WTH'  or db1.market = 'SMT')
												   And ((Tran_Type !='ARCDT')  OR(left(ReserveC1,3) NOT IN ('CDL','SMS','CAT','SDC','FRC','DCD')))) tbp
                                          PIVOT (Sum(value)
                                                 FOR [month] IN ([q01],
                                                                 [q02],
                                                                 [q03],
                                                                 [q04],
                                                                 [q05],
                                                                 [q06],
                                                                 [q07],
                                                                 [q08],
                                                                 [q09],
                                                                 [q10],
                                                                 [q11],
                                                                 [q12])) pt
                                    UNION ALL
                                            select 
                                            '',
                                            'Total ' + '" + Year + @"',
                                            '',
                                            '',

                                            [01],[02],[03],[04],[05],[06],[07],[08],[09],[10],[11],[12],
	                                            Sum(Isnull([01], 0) + Isnull([02], 0) + Isnull([03], 0)
	                                              + Isnull( [04], 0) + Isnull ( [05], 0) + Isnull( [06], 0)
	                                              + Isnull( [07], 0) + Isnull ([08], 0) + Isnull( [09], 0)
	                                              + Isnull( [10], 0) + Isnull( [11], 0) + Isnull( [12], 0)),
                                            '',
                                            NULL,
                                            NULL
                                            from(
	                                            select Qty,right(yyyymm,2) [mm] 
	                                            from xCurrent" + @" 
	                                            join xOMA_Code" + @" item  
	                                            ON item.itemcode = stock_code
	                                            where (" + sumTimeRange + @"  )
                                                AND stock_code IS NOT NULL              
	                                            " + SupplierSellect + @"
	                                            " + BrandSelect + @"
                                                And (market = 'CLS' or market = 'WTH'  or market = 'SMT')
	                                            And ((Tran_Type !='ARCDT')  OR(left(ReserveC1,3) NOT IN ('CDL','SMS','CAT','SDC','FRC','DCD')))) as p
                                            pivot (sum(Qty) for [mm] in ([01],[02],[03],[04],[05],[06],[07],[08],[09],[10],[11],[12]) ) as piv
                                            group by [01],[02],[03],[04],[05],[06],[07],[08],[09],[10],[11],[12]

                                    UNION ALL
                                    select 
                                            '',
                                            'Total '+ Cast(Cast('" + Year + @"' AS INT)-1 AS VARCHAR(4)),
                                            '',
                                            '',
                                            [01],[02],[03],[04],[05],[06],[07],[08],[09],[10],[11],[12],
	                                            Sum(Isnull([01], 0) + Isnull([02], 0) + Isnull([03], 0)
	                                              + Isnull( [04], 0) + Isnull ( [05], 0) + Isnull( [06], 0)
	                                              + Isnull( [07], 0) + Isnull ([08], 0) + Isnull( [09], 0)
	                                              + Isnull( [10], 0) + Isnull( [11], 0) + Isnull( [12], 0)),
                                            '',
                                            NULL,
                                            NULL
                                            from(
	                                            select Qty,right(yyyymm,2) [mm] 
	                                            from xCurrent" + @" 
	                                            join xOMA_Code" + @" item  
	                                            ON item.itemcode = stock_code
	                                            where (" + sumTimeRangeLY + @"  )
                                                AND stock_code IS NOT NULL              
	                                            " + SupplierSellect + @"
	                                            " + BrandSelect + @"
                                                And (market = 'CLS' or market = 'WTH'  or market = 'SMT')
	                                            And ((Tran_Type !='ARCDT')  OR(left(ReserveC1,3) NOT IN ('CDL','SMS','CAT','SDC','FRC','DCD')))) as p
                                            pivot (sum(Qty) for [mm] in ([01],[02],[03],[04],[05],[06],[07],[08],[09],[10],[11],[12]) ) as piv

                                            group by [01],[02],[03],[04],[05],[06],[07],[08],[09],[10],[11],[12]
                                                                )db";
                }
                else if (sqlmode == 8)
                {
                    sql = @"SELECT stock_code,
                                   itemname,
                                   cardcode [Co.],
                                   cardname [Supplier],
                                   s01      [Jan],
                                   s02      [Feb],
                                   s03      [Mar],
                                   s04      [Apr],
                                   s05      [May],
                                   s06      [Jun],
                                   s07      [Jul],
                                   s08      [Aug],
                                   s09      [Sep],
                                   s10      [Oct],
                                   s11      [Nov],
                                   s12      [Dec],
                                   subtotal
                                   --'',
                                   --onhand   [Onhand],
                                  -- onorder
                            FROM   (SELECT stock_code,
                                           itemname,
                                           cardcode,
                                           cardname,
                                           s01,
                                           s02,
                                           s03,
                                           s04,
                                           s05,
                                           s06,
                                           s07,
                                           s08,
                                           s09,
                                           s10,
                                           s11,
                                           s12,
                                           Isnull(s01, 0) + Isnull(s02, 0) + Isnull(s03, 0)
                                           + Isnull( s04, 0) + Isnull( s05, 0) + Isnull( s06, 0)
                                           + Isnull( s07, 0) + Isnull(s08, 0) + Isnull( s09, 0)
                                           + Isnull ( s10, 0) + Isnull ( s11, 0) + Isnull ( s12, 0)
                                           [SUBTOTAL],
                                           ''                                                       [EM],
                                            onhand,
                                           onorder
                                    FROM   (SELECT stock_code,
                                                   year,
                                                   value,
                                                   month,
                                                    onhand,
                                                   onorder,
                                                   itemname,
                                                   cardcode,
                                                   cardname
                                            FROM   (SELECT  stock_code                    [Stock_Code],
                                                           LEFT(yyyymm, 4)        [Year],
                                                           value,
                                                           market,
                                                            Tran_Type,
														   ReserveC1,
                                                           col + RIGHT(yyyymm, 2) [month],YYYYMM
                                                    FROM  xCurrent " + @"
                                                           CROSS apply(VALUES ( qty,'Q' ),(Net,'S')) x(value, col))db1
                                                   LEFT JOIN xOMA_Code" + @" item
                                                          ON item.itemcode = db1.stock_code
                                            WHERE  (" + TimeRange + @"  )
                                                   AND stock_code IS NOT NULL
                                                   " + SupplierSellect + @"
                                                   " + BrandSelect + @"
                                                   And (db1.market = 'CLS' or db1.market = 'WTH'  or db1.market = 'SMT')
												   And ((Tran_Type !='ARCDT')  OR(left(ReserveC1,3) NOT IN ('CDL','SMS','CAT','SDC','FRC','DCD')))) tbp

                                           PIVOT (Sum(value)
                                                 FOR [month] IN ([s01],
                                                                 [s02],
                                                                 [s03],
                                                                 [s04],
                                                                 [s05],
                                                                 [s06],
                                                                 [s07],
                                                                 [s08],
                                                                 [s09],
                                                                 [s10],
                                                                 [s11],
                                                                 [s12])) pt
                                    UNION ALL
                                            select 
                                            '',
                                            'Total ' + '" + Year + @"',
                                            '',
                                            '',

                                            [01],[02],[03],[04],[05],[06],[07],[08],[09],[10],[11],[12],
	                                            Sum(Isnull([01], 0) + Isnull([02], 0) + Isnull([03], 0)
	                                              + Isnull( [04], 0) + Isnull ( [05], 0) + Isnull( [06], 0)
	                                              + Isnull( [07], 0) + Isnull ([08], 0) + Isnull( [09], 0)
	                                              + Isnull( [10], 0) + Isnull( [11], 0) + Isnull( [12], 0)),
                                            '',
                                            NULL,
                                            NULL
                                            from(
	                                            select Net,right(yyyymm,2) [mm] 
	                                            from xCurrent" + @" 
	                                            join xOMA_Code" + @" item  
	                                            ON item.itemcode = stock_code
	                                            where (" + sumTimeRange + @"  )
                                                AND stock_code IS NOT NULL              
	                                            " + SupplierSellect + @"
	                                            " + BrandSelect + @"
                                                And (market = 'CLS' or market = 'WTH'  or market = 'SMT')
	                                            And ((Tran_Type !='ARCDT')  OR(left(ReserveC1,3) NOT IN ('CDL','SMS','CAT','SDC','FRC','DCD')))) as p
                                            pivot (sum(Net) for [mm] in ([01],[02],[03],[04],[05],[06],[07],[08],[09],[10],[11],[12]) ) as piv
                                            group by [01],[02],[03],[04],[05],[06],[07],[08],[09],[10],[11],[12]

                                    UNION ALL
                                    select 
                                            '',
                                            'Total '+ Cast(Cast('" + Year + @"' AS INT)-1 AS VARCHAR(4)),
                                            '',
                                            '',
                                            [01],[02],[03],[04],[05],[06],[07],[08],[09],[10],[11],[12],
	                                            Sum(Isnull([01], 0) + Isnull([02], 0) + Isnull([03], 0)
	                                              + Isnull( [04], 0) + Isnull ( [05], 0) + Isnull( [06], 0)
	                                              + Isnull( [07], 0) + Isnull ([08], 0) + Isnull( [09], 0)
	                                              + Isnull( [10], 0) + Isnull( [11], 0) + Isnull( [12], 0)),
                                            '',
                                            NULL,
                                            NULL
                                            from(
	                                            select Net,right(yyyymm,2) [mm] 
	                                            from xCurrent" + @" 
	                                            join xOMA_Code" + @" item  
	                                            ON item.itemcode = stock_code
	                                            where (" + sumTimeRangeLY + @"  )
                                                AND stock_code IS NOT NULL              
	                                            " + SupplierSellect + @"
	                                            " + BrandSelect + @"
                                                And (market = 'CLS' or market = 'WTH'  or market = 'SMT')
	                                            And ((Tran_Type !='ARCDT')  OR(left(ReserveC1,3) NOT IN ('CDL','SMS','CAT','SDC','FRC','DCD')))) as p
                                            pivot (sum(Net) for [mm] in ([01],[02],[03],[04],[05],[06],[07],[08],[09],[10],[11],[12]) ) as piv

                                            group by [01],[02],[03],[04],[05],[06],[07],[08],[09],[10],[11],[12]
                                                                 )db";
                }
                return sql + '|' + connection;
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
                return null;
            }





        }

        public void FormatSetup(DataTable dataTable, DataTable dataTable2, string SheetName, String CYear, string unitofMeasure, ISheet sheet, HSSFFont myFont)
        {
            try
            {
                string Year = CYear.ToString();
                IRow row;
                ICell rowCell;
                //first row
                row = sheet.CreateRow(0);
                //first row, fourth cell
                rowCell = row.CreateCell(3);
                //title
                var source = new HSSFRichTextString("Source");
                source.ApplyFont(myFont);
                rowCell.SetCellValue(source);
                //first row, fifth cell
                rowCell = row.CreateCell(4);
                var omes = new HSSFRichTextString("OM Ex-Warehouse Sales");
                omes.ApplyFont(myFont);
                rowCell.SetCellValue(omes);
                //merge colomn
                sheet.AddMergedRegion(new CellRangeAddress(0, 0, 4, 6));
                //second row
                row = sheet.CreateRow(1);
                //second row, fourth cell
                rowCell = row.CreateCell(3);
                var uom = new HSSFRichTextString("Unit of Measurement:");
                uom.ApplyFont(myFont);
                rowCell.SetCellValue(uom);
                //second row, fifth cell
                rowCell = row.CreateCell(4);
                var UOM = new HSSFRichTextString(unitofMeasure);
                UOM.ApplyFont(myFont);
                rowCell.SetCellValue(UOM);
                sheet.AddMergedRegion(new CellRangeAddress(1, 1, 4, 6));
                //third row
                row = sheet.CreateRow(2);
                //third row, fourth cell
                rowCell = row.CreateCell(3);
                var year1 = new HSSFRichTextString("Year:");
                year1.ApplyFont(myFont);
                rowCell.SetCellValue(year1);
                //third row, fifth cell
                rowCell = row.CreateCell(4);
                var year = new HSSFRichTextString(Year);
                year.ApplyFont(myFont);
                rowCell.SetCellValue(year);
                //forth row
                row = sheet.CreateRow(3);
                //forth row, fourth cell
                rowCell = row.CreateCell(3);
                var area = new HSSFRichTextString("Area:");
                // test.UnicodeString = "";
                area.ApplyFont(myFont);
                rowCell.SetCellValue(area);
                //forth row, fifth cell
                rowCell = row.CreateCell(4);
                var au = new HSSFRichTextString("AU");
                au.ApplyFont(myFont);
                rowCell.SetCellValue(au);
                

                // myFont.FontHeight = 10 * 10;
                //rowCell.CellStyle.SetFont(myFont);


                //for(int i = 1; i < 40; i++)
                //{
                //    sheet.AutoSizeColumn(i);
                //}

                int startRow = 4;


                /////First table/////
                TotalFormatSetup(dataTable, SheetName, startRow, sheet, myFont);

                /////second table////
                int table1rows = dataTable.Rows.Count;
                int lastrow = (table1rows + startRow - 1);
                int StartRow2 = lastrow + 4;
                string LYear = (Int32.Parse(CYear) - 1).ToString();
                row = sheet.CreateRow(StartRow2);
                row.CreateCell(3).SetCellValue("Year:");
                row.CreateCell(4).SetCellValue(LYear);
                StartRow2++;
                TotalFormatSetup(dataTable2, SheetName, StartRow2, sheet, myFont);

            

            }
            catch(Exception e)
            {
                MessageBox.Show(e.Message);
            }
        }

        public void TotalFormatSetup(DataTable dataTable, string SheetName, int startRow, ISheet sheet, HSSFFont myFont)
        {
            try
            {
                string[] columnName = { "Stock_Code", "Item Name","Supplier Co.","Supplier",
                            "Jan",
                            "Feb",
                            "Mar",
                            "Apr",
                            "May",
                            "Jun",
                            "Jul",
                            "Aug",
                            "Sep",
                            "Oct",
                            "Nov",
                            "Dec",
                            "Total",
                            "",
                            "Onhand ",
                            "OnOrder"};

                int rows = dataTable.Rows.Count; //not the sheet itself
                int cols = dataTable.Columns.Count;
                if (rows == 0 || cols == 0)
                {
                   // throw new Exception("data is empty");
                    return;
                }

                //object[,] data = StoreTableToarray(dataTable);

                IRow row;
                ICell rowCell;
                // the fifth row title
                // if(startRow)

                //if(startRow == 4)
                //{
                //    row = sheet.CreateRow(startRow);
                //}
                //else
                //{
                //    startRow = rows + startRow - 1;
                //    row = sheet.CreateRow(startRow);
                //}
                row = sheet.CreateRow(startRow);



                for (int i = 0; i < cols; i++)//column name put on the sheet
                {
                    //rowCell = row.CreateCell(i);
                    //rowCell.SetCellValue(columnName[i]);

                    rowCell = row.CreateCell(i);

                    var title = new HSSFRichTextString(columnName[i]);
                    title.ApplyFont(myFont);
                    rowCell.SetCellValue(title);



                }

             

                //----------------------------------------------------------------------------------------
                List<string> columns = new List<string>();
               


                int columnIndex = 0;

                foreach (System.Data.DataColumn column in dataTable.Columns)
                {
                    columns.Add(column.ColumnName);
                   // row.CreateCell(columnIndex).SetCellValue(column.ColumnName);
                    columnIndex++;
                }

                int rowIndex = startRow + 1;
                // write data from datatable to excel

                foreach (DataRow dsrow in dataTable.Rows)
                {
                    row = sheet.CreateRow(rowIndex);
                    int cellIndex = 0;
                    foreach (String col in columns)
                    {
                        
                        double b;
                        bool a = Double.TryParse(dsrow[col].ToString(),out b);
                        if(a == false)
                        {
                            //String cellValue = dsrow[col] == null ? " " : dsrow[col].ToString();
                            row.CreateCell(cellIndex).SetCellValue(dsrow[col].ToString());
                        }
                        else
                        {
                            row.CreateCell(cellIndex).SetCellValue(b);
                        }



                        #region add bold 
                        //if (dataTable.Rows.IndexOf(dsrow) == dataTable.Rows.Count-2 || dataTable.Rows.IndexOf(dsrow) == dataTable.Rows.Count -1)
                        //{
                        //    rowCell = row.CreateCell(cellIndex);
                        //    var total = new HSSFRichTextString(dsrow[col].ToString());
                        //    //total.ApplyFont(myFont);
                        //    //rowCell.SetCellValue(total);
                        //    total.ApplyFont(myFont);
                        //    //rowCell.SetCellValue(dsrow[col].ToString());
                        //    string totalString = total.ToString()==""?"": total.ToString();


                        //    //add bold to the cell
                        //    if (totalString.Contains("Total") || totalString =="")
                        //    {
                        //        //total = new HSSFRichTextString(dsrow[col].ToString());
                        //        //total.ApplyFont(myFont);

                        //        rowCell.SetCellValue(total);
                        //    }
                        //    else
                        //    {
                        //        //total = new HSSFRichTextString(dsrow[col].ToString());
                        //        //total.ApplyFont(myFont);
                        //        total = new HSSFRichTextString(total.ToString());


                        //        rowCell.SetCellValue(Convert.ToDouble(totalString));
                        //        // rowCell.SetCellValue(total);
                        //        Console.WriteLine(total);

                        //    }


                        //   // var total1 = new HSSFRichTextString(totalString);


                        //}
                        #endregion
                        cellIndex++;

                    }

                    rowIndex++;
                }

                row = sheet.CreateRow(rowIndex);
                row.CreateCell(1).SetCellValue("Growth");

                      

                for (int i = 4; i < 17; i++)
                {
                    if(sheet.GetRow(rowIndex - 2).GetCell(i).ToString() != "" && sheet.GetRow(rowIndex - 1).GetCell(i).ToString() != "")
                    {
                        string a = sheet.GetRow(rowIndex - 2).GetCell(i).ToString();
                        string b = sheet.GetRow(rowIndex - 1).GetCell(i).ToString();

                        if (Convert.ToDouble(sheet.GetRow(rowIndex - 2).GetCell(i).ToString()) >= Convert.ToDouble(sheet.GetRow(rowIndex - 1).GetCell(i).ToString()))
                        {
                            double  growth = (Convert.ToDouble(sheet.GetRow(rowIndex - 2).GetCell(i).ToString()) - Convert.ToDouble(sheet.GetRow(rowIndex - 1).GetCell(i).ToString())) / Convert.ToDouble(sheet.GetRow(rowIndex - 1).GetCell(i).ToString()) * 100;
                            String rate = growth.ToString("0.00") + "%";
                            row.CreateCell(i).SetCellValue(rate);
                        }
                        else if (Convert.ToDouble(sheet.GetRow(rowIndex - 2).GetCell(i).ToString()) < Convert.ToDouble(sheet.GetRow(rowIndex - 1).GetCell(i).ToString()))
                        {
                           // string b = sheet.GetRow(rowIndex - 2).GetCell(i).ToString();
                            double growth = (Convert.ToDouble(sheet.GetRow(rowIndex - 1).GetCell(i).ToString()) - Convert.ToDouble(sheet.GetRow(rowIndex - 2).GetCell(i).ToString())) / Convert.ToDouble(sheet.GetRow(rowIndex - 1).GetCell(i).ToString()) * 100;
                            String rate = "-" + growth.ToString("0.00") + "%";
                            row.CreateCell(i).SetCellValue(rate);
                        }
                    }
                    else if (sheet.GetRow(rowIndex - 1).GetCell(i).ToString() == "")
                    {

                        row.CreateCell(i).SetCellValue("");
                    }
                    else
                    {
                        row.CreateCell(i).SetCellValue("-100.00%");
                    }


                }

                for (int i = 1; i < 20; i++)
                {
                    sheet.AutoSizeColumn(i);
                }


            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
            }
        }
    }
}
