using CsvHelper;
using Nito.AsyncEx;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OD4BUrlsUpdateExcel
{
    class Program
    {
        static void Main(string[] args)
        {
            AsyncContext.Run(() => Async(args));
        }

        static async void Async(string[] args)
        {
            FileInfo file = new FileInfo("C:\\charlie\\MS Graph\\CsvHelper\\CustodianList_AfterJuly12th.xlsx");

            using (var package = new ExcelPackage(file))
            {
                System.Console.WriteLine("StartTime:" + DateTime.Now);
                GraphService graphService = new GraphService();
                ExcelWorkbook workBook = package.Workbook;
                ExcelWorksheet currentWorksheet = workBook.Worksheets.SingleOrDefault(w => w.Name == "custodians");
                int count = 0;
                int totalRows = currentWorksheet.Dimension.End.Row;
                int totalCols = currentWorksheet.Dimension.End.Column;

                for (int i = 2; i <= totalRows; i++)
                {
                    try
                    {
                        string msAlias = string.Empty;
                        if (currentWorksheet.Cells[i, 6].Value != null)
                            msAlias = currentWorksheet.Cells[i, 6].Value.ToString();
                        if (currentWorksheet.Cells[i, 7].Value.ToString() == "NULL")
                        {
                            var driveUrl = await graphService.GetOD4BUrl(msAlias);
                            currentWorksheet.Cells[i, 7].Value = driveUrl;
                            count++;
                            if (count % 100 == 0)
                            {
                                System.Console.WriteLine("100 Updated at:" + DateTime.Now);
                                package.Save();
                            }
                        }

                    }
                    catch (Exception ex)
                    {

                        package.Save();
                        System.Console.WriteLine(ex.ToString());
                    }
                }

                package.Save();
            }
        }
        static async void MainAsync(string[] args)
        {

            System.Console.WriteLine("StartTime:" + DateTime.Now);
            //Stream reader will read test.csv file in current folder
            StreamReader sr = new StreamReader("C:\\charlie\\MS Graph\\CsvHelper\\custodians.csv", Encoding.UTF8);
            StreamWriter write = new StreamWriter("C:\\charlie\\MS Graph\\CsvHelper\\custodiansOutput.csv");
            //Csv reader reads the stream
            CsvReader csvread = new CsvReader(sr);
            CsvWriter csw = new CsvWriter(write);
            //csvread will fetch all record in one go to the IEnumerable object record
            IEnumerable<Custodian> custodians = csvread.GetRecords<Custodian>().ToList();
            GraphService graphService = new GraphService();
            int i = 0;
            try
            {
                foreach (var rec in custodians) // Each record will be fetched and printed on the screen
                {
                    var driveUrl = await graphService.GetOD4BUrl(rec.MSAlias);

                    rec.OD4BUrls = driveUrl;
                    csw.WriteRecord<Custodian>(rec);
                    i++;
                    if (i % 100 == 0)
                    {
                        System.Console.WriteLine("100 Updated at:" + DateTime.Now);
                    }
                }
                sr.Close();
                write.Close(); 

                System.Console.WriteLine("FinishTime:" + DateTime.Now);
            }
            catch (Exception ex)
            {
                sr.Close();
                write.Close();
                System.Console.WriteLine(ex.ToString()); 
            }

        }

    }

    public class Custodian // Test record class
    {
        public string ID { get; set; }
        public string FirstName { get; set; }
        public string LastName { get; set; }
        public string StatusID { get; set; }
        public string PersonnelID { get; set; }
        public string MSAlias { get; set; }
        public string IsEmployee { get; set; }
        public string SharePointUrls { get; set; }
        public string OD4BUrls { get; set; }

    }
}
