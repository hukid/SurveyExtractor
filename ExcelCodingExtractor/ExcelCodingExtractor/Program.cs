using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;

namespace ExcelCodingExtractor
{
    class Program
    {
        static void Main(string[] args)
        {
            // Define the template
            string[,] codingformats = new string[,] {
                //{ "3", "ISE COASSTer survey", "2012", "You are receiving this survey because you participate in COASST. Please tell us briefly (1-2 sentences is fine) why you continue to be involved with this program", "ISE COASSTer survey"},
                //{ "5", "ISE post training survey", "2012", "You are receiving this survey because you have signed up as a COASSTer. Please tell us briefly (1-2 sentences is fine) why you chose to be involved with this program. ", "ISE post training survey"},
                //{ "8", "Satisfaction2009", "2009", "In a few words, please tell us why you participate in the COASST program", "Satisfaction"},
                //{ "9", "ISE COASSTer survey", "2012", " You are receiving this survey because you participate in COASST. Please tell us briefly (1-2 sentences is fine) why you continue to be involved with this program.", "ISE COASSTer survey"},
                { "16", "Profile Form", "TBD", "Why do you want to be involved?", "Profile Form"},
                //{ "17", "Satisfaction2004", "2004", "Q1 - How did you first hear about COASST, and why did you decide to join the program?", "Satisfaction Q1"},
                //{ "23", "AISL post training survey", "2016", "People sign up to participate in citizen science programs for many different reasons. Why did you join COASST?", "AISL post training survey" },
                //{ "24", "Satisfaction2004", "2004", "Q2 - Why do you continue volunteering with COASST?",  "Satisfaction Q2" },
                //{ "28", "AISL COASSTer Survey", "2016", "People participate in citizen science programs for many different reasons. Why do you continue to survey for COASST?", "AISL COASSTer Survey" },
            };

            // Get the excel file:
            Application app = new Application();
            //Workbook xlWorkbook = app.Workbooks.Open(@"D:\Projects\ExcelCodingExtractor\codingTables.xlsx");
            Workbook xlWorkbook = app.Workbooks.Open(@"D:\Projects\ExcelCodingExtractor\MDPost2016.xlsx");
            //Workbook xlWorkbook = app.Workbooks.Open(@"D:\Projects\ExcelCodingExtractor\1.xlsx");

            Sheets sheets = xlWorkbook.Sheets;
            _Worksheet worksheet = (_Worksheet)sheets[2];
            Range xlRange = worksheet.UsedRange;

            // Read excel file and format each coding file
            int rowCount = xlRange.Rows.Count;
            int columnCount = xlRange.Columns.Count;

            List<string> codingsTitle = new List<string>();
            List<string> codingsBody = new List<string>();

            //for (int j = 0; j < codingformats.GetLength(0); j++)
            //{
            //    string year = codingformats[j, 2];
            //    string question = codingformats[j, 3];

            //    string source = codingformats[j, 1];
            //    int answerIndex = Int32.Parse(codingformats[j, 0]);
            //    string titleSource = codingformats[j, 4];

            //    for (int i = 2; i <= rowCount; i++)
            //    {
            //        if (year == "TBD" && xlRange.Cells[i, 10] != null && xlRange.Cells[i, 10].Value2 != null)
            //        {
            //            year = xlRange.Cells[i, 10].Value2.ToString();
            //        }

            //        string answer = string.Empty;
            //        if (xlRange.Cells[i, answerIndex] != null && xlRange.Cells[i, answerIndex].Value2 != null)
            //        {
            //            answer = xlRange.Cells[i, answerIndex].Value2.ToString();
            //        }

            //        if (answer != "#N/A" && answer != string.Empty && answer != "0" && answer != "blank")
            //        {
            //            string id = "unknown";
            //            string name = "unknown";
            //            if (xlRange.Cells[i, 1] != null && xlRange.Cells[i, 1].Value2 != null)
            //            {
            //                id = xlRange.Cells[i, 1].Value2.ToString();
            //            }

            //            if (xlRange.Cells[i, 4] != null && xlRange.Cells[i, 4].Value2 != null)
            //            {
            //                name = xlRange.Cells[i, 4].Value2.ToString();
            //            }

            //            string body = $"Volunteer ID: {id}\nFull name: {name}\n\nData source: {source}\nYear: {year}\n\nSurven question:\n{question}\n\n\nParticipant answer:\n{answer}\n";
            //            string fileName = $".\\output\\{id}-{name}-{titleSource}-{year}.txt";

            //            codingsBody.Add(body);
            //            codingsTitle.Add(fileName);
            //        }
            //    }

            //}

            if (!System.IO.Directory.Exists(".\\output"))
            {
                System.IO.Directory.CreateDirectory(".\\output");
            }


            for (int i = 2; i <= rowCount; i++)
            {
                Console.WriteLine("{0} rows", i);
                string id = "unknown";
                string name = "unknown";
                if (xlRange.Cells[i, 1] != null && xlRange.Cells[i, 1].Value2 != null)
                {
                    id = xlRange.Cells[i, 1].Value2.ToString();
                }

                // name is in col 4
                //if (xlRange.Cells[i, 4] != null && xlRange.Cells[i, 4].Value2 != null)
                //{
                //    name = xlRange.Cells[i, 4].Value2.ToString();
                //}

                //// name is in col 2
                //if (xlRange.Cells[i, 2] != null && xlRange.Cells[i, 2].Value2 != null)
                //{
                //    name = xlRange.Cells[i, 2].Value2.ToString();
                //}

                // for orginal table
                //for (int j = 0; j < codingformats.GetLength(0); j++)
                //{
                //    string year = codingformats[j, 2];
                //    if (year == "TBD" && xlRange.Cells[i, 10] != null && xlRange.Cells[i, 10].Value2 != null)
                //    {
                //        year = xlRange.Cells[i, 10].Value2.ToString();
                //    }

                //    string question = codingformats[j, 3];

                //    string source = codingformats[j, 1];
                //    int answerIndex = Int32.Parse(codingformats[j, 0]);
                //    string titleSource = codingformats[j, 4];

                //    string answer = string.Empty;
                //    if (xlRange.Cells[i, answerIndex] != null && xlRange.Cells[i, answerIndex].Value2 != null)
                //    {
                //        answer = xlRange.Cells[i, answerIndex].Value2.ToString();
                //    }

                //    if (answer != "#N/A" && answer != string.Empty && answer != "0" && answer != "blank" && answer != "-2146826246")
                //    {
                //        string body = $"Volunteer ID: {id}\r\nFull name: {name}\r\n\r\nData source: {source}\r\nYear: {year}\r\n\r\nSurvey question:\r\n{question}\r\n\r\n\r\nParticipant answer:\r\n{answer}\r\n";
                //        string fileName = $".\\output\\{id}-{name}-{titleSource}-{year}.txt";

                //        //codingsBody.Add(body);
                //        //codingsTitle.Add(fileName);

                //        System.IO.File.WriteAllText(fileName, body);
                //    }
                //}

                // for post MD
                string year = "2016";
                name = xlRange.Cells[i, 2].Value2.ToString();

                string question = "People sign up to participate in citizen science programs for many different reasons. Why did you join COASST ?";

                string source = "AISL post training survey_MD";
                int answerIndex = 3;
                string titleSource = "AISL post training survey_MD";

                string answer = string.Empty;
                if (xlRange.Cells[i, answerIndex] != null && xlRange.Cells[i, answerIndex].Value2 != null)
                {
                    answer = xlRange.Cells[i, answerIndex].Value2.ToString();
                }

                if (answer != "#N/A" && answer != string.Empty && answer != "0" && answer != "blank" && answer != "-2146826246")
                {
                    string body = $"Volunteer ID: {id}\r\nFull name: {name}\r\n\r\nData source: {source}\r\nYear: {year}\r\n\r\nSurvey question:\r\n{question}\r\n\r\n\r\nParticipant answer:\r\n{answer}\r\n";
                    string fileName = $".\\output\\{id}-{name}-{titleSource}-{year}.txt";

                    //codingsBody.Add(body);
                    //codingsTitle.Add(fileName);

                    System.IO.File.WriteAllText(fileName, body);
                }
            }

            // Write to txt files
            //for (int i = 0; i < codingsBody.Count; i++)
            //{
            //    System.IO.File.WriteAllText(codingsTitle[i], codingsBody[i]);
            //}
        }
    }
}
