using System;
using System.IO;
using System.Data;
using System.Data.OleDb;
using System.Collections.Generic;
using System.Text;

using Excel = Microsoft.Office.Interop.Excel;

namespace XlsToCsv
{
    class Program
    {
        static void Main(string[] args)
        {
            //string sourceFile = @"C:\Users\ssjoshua\Desktop\Current\Data_INFO\SessionHistory_06-01-2014_To_07-31-2014.xlsx";
            string sourceFile = @"C:\Users\SJOSHUA\Desktop\Localization Doc\Zones.xlsx"; 
            DateTime st = DateTime.Now;
            Console.WriteLine("Process Started ...");
            convertExcelToCSV(sourceFile);    //278   
            Console.WriteLine("Total Time : " +(DateTime.Now - st).Milliseconds.ToString());
            Console.WriteLine("Press any key to Close!");
            Console.ReadKey();
        }

        private static void testExcel(string path)
        {
             Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(path);

            int ct = xlWorkbook.Sheets.Count;
            for (int k = 1; k <= ct; k++)
            {
                Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[k];
                Excel.Range xlRange = xlWorksheet.UsedRange;

                int rowCount = xlRange.Rows.Count;
                int colCount = xlRange.Columns.Count;

                bool isExists = System.IO.Directory.Exists(Environment.CurrentDirectory + "\\CSVFiles");
                if (!isExists)
                    System.IO.Directory.CreateDirectory(Environment.CurrentDirectory + "\\CSVFiles");
                StreamWriter wrtr = new StreamWriter(Environment.CurrentDirectory + "\\CSVFiles\\" + xlWorksheet.Name + ".csv");
                for (int i = 1; i <= rowCount; i++)
                {
                    string rowString = "";
                    for (int j = 1; j <= colCount; j++)
                    {
                        try
                        {
                            if (xlRange.Cells[i, j].Value2 != "")
                            {
                                //Console.WriteLine(xlRange.Cells[i, j].Value2.ToString());
                                if (xlRange.Cells[i, j].Value2.ToString().Contains(","))
                                    rowString = rowString + "\"" + xlRange.Cells[i, j].Value2.ToString() + "\",";
                                else
                                    rowString = rowString + xlRange.Cells[i, j].Value2.ToString() + ",";
                            }
                            else if (j == 1 && xlRange.Cells[i, j].Value2 != "")
                                rowString = rowString + ",";

                        }
                        catch (Exception e)
                        {
                            // Console.WriteLine(e.ToString());
                            break;
                        }
                    }
                    //wrtr.WriteLine(rowString);
                    //wrtr.AutoFlush = true;
                }
                wrtr.Close();
            }
        }



        static void convertExcelToCSV(string sourceFile)
        {
            List<string> worksheetName = new List<string>();
            string strConn = String.Format(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + sourceFile + ";Extended Properties=\"Excel 12.0 Xml;HDR=NO;IMEX=1\"");
            OleDbConnection conn = null;
            StreamWriter wrtr = null;
            StreamWriter lg = new StreamWriter("Log.txt");
            OleDbCommand cmd = null;
            OleDbDataAdapter da = null;
            try
            {
                conn = new OleDbConnection(strConn);
                conn.Open();


                DataTable dtk = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                if (dtk != null)
                {
                    for (int i = 0; i < dtk.Rows.Count; i++ )
                        {
                            DataRow row = dtk.Rows[i];
                            string sheetname = row["TABLE_NAME"].ToString().Replace("'", "");
                            sheetname = sheetname.Replace("$", "");
                            //Console.WriteLine(sheetname);
                            worksheetName.Add(sheetname);
                        }
                }

                //dtk.Dispose();


                foreach (string sheetname in worksheetName)
                {
                    cmd = new OleDbCommand("SELECT * FROM ["+sheetname+"$]", conn);
                    cmd.CommandType = CommandType.Text;

                    bool isExists = System.IO.Directory.Exists(Environment.CurrentDirectory + "\\CSVFiles");
                    if (!isExists)
                        System.IO.Directory.CreateDirectory(Environment.CurrentDirectory + "\\CSVFiles");
                    wrtr = new StreamWriter(Environment.CurrentDirectory + "\\CSVFiles\\" + sheetname + ".csv");

                    da = new OleDbDataAdapter(cmd);


                    DataTable dt = new DataTable();

                    da.Fill(dt);

                    int count = dt.Rows.Count;

                    int xvalue = 0;                    

                    //for (int x = 0; x < dt.Rows.Count; x++)
                        foreach (DataRow dR in dt.Rows)
                    {
                        
                        //count--;
                  //  Here : 
                        string rowString = "";
                        for (int y = 0; y < dt.Columns.Count; y++)
                        {
                            if (dt.Rows[xvalue][y].ToString().Trim().Contains(","))
                                rowString += "\"" + dt.Rows[xvalue][y].ToString().Trim() + "\",";
                            else
                                rowString += dt.Rows[xvalue][y].ToString().Trim() + ",";
                        }

                        xvalue++;
                        wrtr.WriteLine(rowString);
                        //if (x == (count))
                        //    Console.WriteLine("Read all Rows");
                    }

                    //if(xvalue != (dt.Rows.Count - 1))
                    //{
                    //    Console.WriteLine(xvalue + " not equal to " + dt.Rows.Count);
                    //    //goto Here;
                    //}
                    //if (count != 1)
                    Console.WriteLine("Read " + (xvalue+1) + " of all " + dt.Rows.Count + " rows in the sheet " + sheetname);

                   
                    cmd.Dispose();
                    da.Dispose();
                    dt.Clear();
                    dt.Dispose();
                    wrtr.Close();
                    wrtr.Dispose();
                }
                Console.WriteLine();
                Console.WriteLine("Done! Your " + sourceFile + " has been converted.");
                Console.WriteLine();
            }
            catch (Exception exc)
            {
                Console.WriteLine(exc.ToString());
                Console.ReadLine();
            }
            finally
            {
                if (conn.State == ConnectionState.Open)
                    conn.Close();
                conn.Dispose();
                cmd.Dispose();
                da.Dispose();
                
                wrtr.Close();
                wrtr.Dispose();
            }
        }       
    }
}