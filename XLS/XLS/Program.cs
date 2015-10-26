using System;
using System.Configuration;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Threading.Tasks;
using System.Drawing;
using Microsoft.Office.Interop.Excel;
namespace XLS
{
    class Program
    {


        static void Main(string[] args)
        {
            String dataFolder = ConfigurationManager.AppSettings["DataFolder"];
            Int32 startRow = Int32.Parse(ConfigurationManager.AppSettings["StartRow"]);
            Int32 maxNumberRows = Int32.Parse(ConfigurationManager.AppSettings["MaxNumberRows"]);
            string list_dir = dataFolder + @"\lists";
            string docs_dir = dataFolder + @"\recepiet";
            string temp_dir = dataFolder + @"\temp";
            string result_dir = dataFolder + @"\result";

            Dictionary<string, string> Bills = new Dictionary<string, string>();
            Dictionary<string, Int32> names = new Dictionary<string, Int32>();
            string[] filePaths = null;
            FillBills(ref filePaths, ref Bills, ref names, list_dir);
            Workbook[] workbooks = new Workbook[filePaths.Length];
            Worksheet[] worksheets = new Worksheet[filePaths.Length];
            Int32[] ranges = new Int32[filePaths.Length];
            Application excel = new Application();
            for (Int16 t = 0; t < filePaths.Length; t++)
            {
                workbooks[t] = excel.Workbooks.Add();
                worksheets[t] = workbooks[t].Worksheets.get_Item(1);
                ranges[t] = 2;
            }

            int count = 0;
            string[] filePaths1 = Directory.GetFiles(docs_dir, "*.xls");
            Range header = null;

            Workbook lostwb = excel.Workbooks.Add();
            Worksheet lostws = lostwb.Worksheets.get_Item(1);
            Int32 lostCounter = 2;
            Int32 copiedRows = 0;
            Int32 lostRows = 0;
            foreach (string path in filePaths1)
            {
                Workbook wb = excel.Workbooks.Open(path);
                Console.WriteLine(path);
                Console.WriteLine("    sheets count - " + wb.Worksheets.Count);
                Worksheet ws1 = (Worksheet)wb.Worksheets.get_Item(3);
                header = ws1.get_Range("A1", "Q1");
                header.Copy(lostws.get_Range("A1", "Q1"));
                for (int i = 1; i <= wb.Worksheets.Count; i++)
                {
                    Worksheet ws = (Worksheet)wb.Worksheets.get_Item(i);
                    Int32 rows = ws.Cells.Rows.Count;
                    Int32 cols = ws.Cells.Columns.Count;
                    if (rows == 0)
                    {
                        continue;
                    }
                    Console.WriteLine(cols + " x " + rows);

                    for (int z = startRow; z < maxNumberRows; z++)
                    {
                        Range rng1 = ws.get_Range("A" + z, "Q" + z);
                        Console.WriteLine(z + ": " + rng1.Cells[1, 3].Text);
                        String key1 = ((String)rng1.Cells[1, 3].Text).Trim();
                        if (key1 != "")
                        {
                            String val = String.Empty;
                            if (Bills.TryGetValue(key1, out val))
                            {
                                Int32 idx = 999;
                                names.TryGetValue(val, out idx);
                                Range to = worksheets[idx].get_Range("A" + (ranges[idx]).ToString(), "Q" + (ranges[idx]).ToString());
                                Range rangeToInsertRow = rng1.EntireRow;
                                rangeToInsertRow.Copy(to);
                                ranges[idx]++;
                                copiedRows++;
                            }
                            else
                            {
                                Range to = lostws.get_Range("A" + (lostCounter).ToString(), "Q" + (lostCounter).ToString());
                                Range rangeToInsertRow = rng1.EntireRow;
                                rangeToInsertRow.Copy(to);
                                lostCounter++;
                                lostRows++;
                            }
                        }
                    }
                }
            }

            Workbook resultwb = excel.Workbooks.Add();
            Worksheet resultws = resultwb.Worksheets.get_Item(1);
            header.Copy(resultws.get_Range("A1", "Q1"));
            Int32 resultCounter = 1;
            for (Int16 t = 0; t < filePaths.Length; t++)
            {
                Worksheet ws = worksheets[t];
                resultCounter++;
                Range resultR = resultws.get_Range("A" + resultCounter, "A" + resultCounter);
                resultR.Style.Font.Size = 14;
                resultR.Interior.Color = Color.YellowGreen;
                resultR.Cells[1, 1] = "Счёт: " + Path.GetFileNameWithoutExtension(filePaths[t]);
                resultCounter++;

                for (int i = 2; i <= ranges[t]; i++)
                {
                    Range rng = ws.get_Range("A" + i, "Q" + i);
                    Range to = resultws.get_Range("A" + resultCounter, "Q" + resultCounter);
                    rng.Copy(to);
                    resultCounter++;
                    Console.WriteLine("file: " + (t + 1).ToString() +" string: " + i + " done");
                }


            }

            for (Int16 t = 0; t < filePaths.Length; t++)
            {
                String name = result_dir + "\\" + Path.GetFileNameWithoutExtension(filePaths[t]) + ".xlsx";
                header.Copy(worksheets[t].get_Range("A1", "Q1"));
                worksheets[t].Columns.AutoFit();
                workbooks[t].SaveAs(name);
                //workbooks[t].SaveAs(name, XlFileFormat.xlWorkbookNormal, null, null, null, null, XlSaveAsAccessMode.xlExclusive, null, null, null, null, null);
                Console.WriteLine(name + " saved");
                //workbooks[t].Close();
            }
            String name1 = result_dir + "\\" + "Report" + ".xlsx";
            String name2 = result_dir + "\\" + "Lost" + ".xlsx";
            resultws.Columns.AutoFit();
            lostws.Columns.AutoFit();
            resultwb.SaveAs(name1);
            lostwb.SaveAs(name2);
            resultwb.Close();
            lostwb.Close();

            excel.Quit();

            Console.WriteLine("All done");
            Console.WriteLine("copied rows: " + copiedRows);
            Console.WriteLine("lost rows: " + lostRows);
        }
        private static bool CheckRange(Range r, out String key, int t)
        {
            key = ((String)r.Cells[1, 3].Text).Trim();
            return true;
        }
        private static void FillBills(ref string[] filePaths, ref Dictionary<string, string> Bills, ref Dictionary<string, Int32> names, String list_dir)
        {
            filePaths = Directory.GetFiles(list_dir, "*.txt");
            int t = 0;
            foreach (string path in filePaths)
            {
                StreamReader streamReader = new StreamReader(path);
                string text = streamReader.ReadToEnd();
                streamReader.Close();
                string bill_name = Path.GetFileNameWithoutExtension(path);

                names.Add(bill_name, t);
                t++;
                string[] text_list = text.Split(',');
                foreach (string depart in text_list)
                {
                    try
                    {
                        Bills.Add(depart.Trim(), bill_name);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex);
                    }
                }
            }
        }

    }
}
