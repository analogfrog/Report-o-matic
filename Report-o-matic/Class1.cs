using Microsoft.VisualBasic.FileIO;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Report_o_matic
{
    
    class Record
    {
        private string name;
        private string clockNum;
        private DateTime submitDate;
        private bool approved;
        private float total;

        List<string> garble;

        //Getters and setters.
        public string Name { get => name; set => name = value; }
        public string ClockNum { get => clockNum; set => clockNum = value; }
        public DateTime SubmitDate { get => submitDate; set => submitDate = value; }
        public bool Approved { get => approved; set => approved = value; }
        public float Total { get => total; set => total = value; }

        public Record()
        {
            garble = new List<string>();
        }

        public void addField( string entry)
        {
            garble.Add(entry);
        }

        public void SetData(List<string> headers)
        {

            Name = garble.ElementAt(headers.IndexOf("Role 1 :: Recipient Name"));
            string strSubDate = garble.ElementAt(headers.IndexOf("Role 1 :: DateSigned"));
            submitDate = Convert.ToDateTime(strSubDate.Substring(strSubDate.Length - 4));
            ClockNum = garble.ElementAt(headers.IndexOf("Role 1 :: Clock #"));
            if (garble.ElementAt(headers.IndexOf("Role 2 :: Radio Group dd410e1b-d430-4c8d-a0a8-534e161da7fd")) == "Approve") // TODO: this element needs to be renamed.
                Approved = true;
            else
                Approved = false;

                 
            
            //This sets the vars for the purchase date of the order.
            //This will help determine if the entries need to go in previous or current year during the month of January.
            string strPurchDate = garble.ElementAt(headers.IndexOf("Role 1 :: Date 1"));
            DateTime date = Convert.ToDateTime(strPurchDate);

        }

    }


    class Report
    {


        List<string> headers;
        List<Record> dataList;

        //TODO: Find what needs to be on report and make vars based on this.

        public Report()
        {
            System.Console.WriteLine("WARNING: No file specified!");
        }

        public Report(string inFile)
        {
            headers = new List<string>();
            dataList = new List<Record>();

            ReadList(inFile, headers, dataList);
        }

        public static void ReadList(string inFile, List<string> headers, List<Record> dataList)
        {


            using (TextFieldParser parser = new TextFieldParser(inFile))
            {
                parser.TextFieldType = FieldType.Delimited;
                parser.SetDelimiters(",");

                string[] fields = parser.ReadFields();
                foreach (string field in fields)
                    headers.Add(field);


                //Process entire file
                while (!parser.EndOfData)
                {
                    //Process row
                    fields = parser.ReadFields();
                    Record newRecord = new Record();
                    foreach (string field in fields)
                    {
                        newRecord.addField(field);
                    }
                    dataList.Add(newRecord);
                    dataList.Last().SetData(headers);
                }
            }
        }


        public static void SetupWorkbook()
        {
            int day = DateTime.Today.Day;
            int month = DateTime.Today.Month;
            int year = DateTime.Today.Year;

            string fileName = "Wellness Subsidy Report " + month + "." + day + "." + year;

            using (FileStream stream = new FileStream(@"C:\Temp\" + fileName + ".xlsx", FileMode.Create, FileAccess.Write))
            {
                IWorkbook wb = new XSSFWorkbook();
                ISheet sheet = wb.CreateSheet("Subsidy Data");

                wb.Write(stream);
            }
        }

        static void Main(string[] args)
        {

            System.Console.WriteLine("Reading List...");
            //ReadList(@"c:\Temp\testcsv.csv");
            new Report(@"c:\Temp\testcsv.csv");
            System.Console.WriteLine("List assembled.");
            System.Console.WriteLine("Writing report...");
            Console.ReadKey();
            SetupWorkbook();
            System.Console.WriteLine("Report written.");
            System.Console.WriteLine("Press any key to exit.");
            Console.ReadKey();
        }
    }
}
