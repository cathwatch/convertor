using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Text.Json;
using System.Web;
using System.Text.RegularExpressions;
using System.Text.Encodings.Web;
using System.Text.Unicode;

namespace D
{
    internal class Program
    {
        static void Main(string[] args)
        {
            string CurrentList = "12 тиждень"; // назва сторінки
            const string PathXlsx = @"C:\Users\zarec\OneDrive\Рабочий стол\DAvid\in.xlsx"; // шлях до файлу
            const string PathJson = @"C:\Users\zarec\OneDrive\Рабочий стол\DAvid\info.json";
            List<days> Days = new List<days>();
            Rozks Main = new Rozks();

            //DaysData
            string day = "";

            int i, j;
            List<lesson> lessons = new List<lesson>();
            string starttime, name, teacher, description, location, endtime, lastDay = "";
            bool isOnline = false;
            int currentIndex = 0;

            name = GetCellValue(PathXlsx, CurrentList, $"AS{10}");

            for (i = 0; i < 6; i++)
            {
                day = GetCellValue(PathXlsx, CurrentList, $"AQ{(i * 19) + 10}");

                if (day == null)
                    break;
                else
                    lastDay = day;

                day = day.Split(' ')[1];

                DateTime d = Convert.ToDateTime(day).ToUniversalTime();

                day = d.ToString("o");

                lessons = new List<lesson>();
                currentIndex = 0;

                for (j = 10 + (19 * i); j <= 25 + (19 * i); j += 3)
                {
                    isOnline = false;
                    name = GetCellValue(PathXlsx, CurrentList, $"AS{j}");
                    if(name != "")
                    {
                        starttime = GetCellValue(PathXlsx, CurrentList, $"AR{j}");
                        teacher = GetCellValue(PathXlsx, CurrentList, $"AS{j + 2}");
                        description = GetCellValue(PathXlsx, CurrentList, $"AY{j}");
                        location = GetCellValue(PathXlsx, CurrentList, $"AZ{j}");
                        starttime = starttime.Replace('-', ':');
                        endtime = endTime(starttime);

                        if (location == "")
                        {
                            location = @"calendar.google.com/calendar";
                            isOnline = true;
                        }

                        lessons.Add(new lesson
                        {
                            starttime = Convert.ToString(starttime),
                            name = name,
                            teacher = teacher.ToString(),
                            description = description,
                            location = location.ToString(),
                            id = currentIndex++.ToString(),
                            type = 0,
                            endtime = endtime,
                            isOnline = isOnline
                        }) ;
                    }
                }

                Days.Add(new days
                {
                    date = day,
                    lessons = lessons
                });
            }

            if(day == null)
            {
                DateTime d = Convert.ToDateTime(lastDay).ToUniversalTime().AddDays(1);

                day = d.ToString("o");

                Days.Add(new days
                {
                    date = day
                });

                d = Convert.ToDateTime(lastDay).ToUniversalTime().AddDays(2);

                day = d.ToString("o");

                Days.Add(new days
                {
                    date = day
                });
            }
            else
            {
                DateTime d = Convert.ToDateTime(day).ToUniversalTime().AddDays(1);

                day = d.ToString("o");

                Days.Add(new days
                {
                    date = day
                });
            }

            Main.days = Days;

            var jsonString = JsonSerializer.Serialize(Main);
            var options = new JsonSerializerOptions
            {
                Encoder = JavaScriptEncoder.Create(UnicodeRanges.BasicLatin, UnicodeRanges.Cyrillic),
                WriteIndented = true
            };

            jsonString = JsonSerializer.Serialize(Main, options);
            File.WriteAllText(PathJson, jsonString);
        }

        private static string endTime(string time)
        {
            switch (time)
            {
                case "9:00":
                    return "10:10";
                case "10:10":
                    return "11:10";
                case "11:20":
                    return "12:20";
                case "12:30":
                    return "13:30";
                case "13:40":
                    return "14:40";
                case "14:50":
                    return "15:50";
                case "16:00":
                    return "17:00";
            }

            return null;
        }
        
        public class Rozks
        {
            public List<days> days { get; set; }
        }

        public class days
        {
            public string date { get; set; }
            public List<lesson> lessons { get; set; }
        }

        public class lesson
        {
            public string description { get; set; }
            public string endtime { get; set; }
            public string id { get; set; }
            public bool isOnline { get; set; }
            public string location { get; set; }
            public string name { get; set; }
            public string starttime { get; set; }
            public string teacher { get; set; }
            public ushort type { get; set; }
        }

        public static string GetCellValue(string fileName, string sheetName, string addressName)
        {
            string value = null;

            // Open the spreadsheet document for read-only access.
            using (SpreadsheetDocument document =
                SpreadsheetDocument.Open(fileName, false))
            {
                // Retrieve a reference to the workbook part.
                WorkbookPart wbPart = document.WorkbookPart;

                // Find the sheet with the supplied name, and then use that 
                // Sheet object to retrieve a reference to the first worksheet.
                Sheet theSheet = wbPart.Workbook.Descendants<Sheet>().
                  Where(s => s.Name == sheetName).FirstOrDefault();

                // Throw an exception if there is no sheet.
                if (theSheet == null)
                {
                    throw new ArgumentException("sheetName");
                }

                // Retrieve a reference to the worksheet part.
                WorksheetPart wsPart =
                    (WorksheetPart)(wbPart.GetPartById(theSheet.Id));

                // Use its Worksheet property to get a reference to the cell 
                // whose address matches the address you supplied.
                Cell theCell = wsPart.Worksheet.Descendants<Cell>().
                  Where(c => c.CellReference == addressName).FirstOrDefault();

                // If the cell does not exist, return an empty string.
                if (theCell != null)
                {
                    value = theCell.InnerText;

                    // If the cell represents an integer number, you are done. 
                    // For dates, this code returns the serialized value that 
                    // represents the date. The code handles strings and 
                    // Booleans individually. For shared strings, the code 
                    // looks up the corresponding value in the shared string 
                    // table. For Booleans, the code converts the value into 
                    // the words TRUE or FALSE.
                    if (theCell.DataType != null)
                    {
                        switch (theCell.DataType.Value)
                        {
                            case CellValues.SharedString:

                                // For shared strings, look up the value in the
                                // shared strings table.
                                var stringTable =
                                    wbPart.GetPartsOfType<SharedStringTablePart>()
                                    .FirstOrDefault();

                                // If the shared string table is missing, something 
                                // is wrong. Return the index that is in
                                // the cell. Otherwise, look up the correct text in 
                                // the table.
                                if (stringTable != null)
                                {
                                    value =
                                        stringTable.SharedStringTable
                                        .ElementAt(int.Parse(value)).InnerText;
                                }
                                break;

                            case CellValues.Boolean:
                                switch (value)
                                {
                                    case "0":
                                        value = "FALSE";
                                        break;
                                    default:
                                        value = "TRUE";
                                        break;
                                }
                                break;
                        }
                    }
                }
                document.Close();
            }
            return value;
        }
    }
}
