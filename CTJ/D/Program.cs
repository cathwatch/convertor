using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Text.Json;
using IronXL;
using System.Text.Encodings.Web;
using System.Text.Unicode;

namespace D
{
    internal class Program
    {
        static GetCellInfo[] GetInfo = new GetCellInfo[4];

        static void SetInfo()
        {
            GetInfo[0] = new GetCellInfo()
            {
                Time = "AR",
                Descriprion = "AY",
                Name = "AS",
                Teacher = "AS",
                Location = "AZ"
            };

            GetInfo[1] = new GetCellInfo()
            {
                Time = "AR",
                Descriprion = "BG",
                Name = "BA",
                Teacher = "BA",
                Location = "BH"
            };

            GetInfo[2] = new GetCellInfo()
            {
                Time = "AR",
                Descriprion = "BO",
                Name = "BI",
                Teacher = "BI",
                Location = "BP"
            };

            GetInfo[3] = new GetCellInfo()
            {
                Time = "AR",
                Descriprion = "BW",
                Name = "BQ",
                Teacher = "BQ",
                Location = "BX"
            };
        }

        static void Main(string[] args)
        {
            SetInfo();

            string FilePath = @"C:\Users\zarec\OneDrive\Рабочий стол\DAvid\in.xlsx";
            const string PathJson = @"C:\Users\zarec\OneDrive\Рабочий стол\DAvid\info.json";
            string SheetName = "12 тиждень";
            WorkBook workbook = WorkBook.Load(FilePath);
            WorkSheet sheet = workbook.GetWorkSheet(SheetName);

            masLessons[] lessons = new masLessons[4];
            List<masDays> Days = new List<masDays>();

            int i, j, g, l, currentIndex, countSplit = 0;
            string starttime = "", name, teacher = "", description = "", location = "", endtime = "", lastDay = "";
            bool isOnline = false;
            string day;

            for (i = 0; i < 6; i++)
            {
                day = sheet[$"AQ{(i * 19) + 10}"].ToString();

                if (day == "")
                    continue;

                lessons = new masLessons[4];

                for (j = 0; j < lessons.Length; j++)
                {
                    lessons[j] = new masLessons();
                }

                currentIndex = 0;

                for (j = 10 + (19 * i); j <= 25 + (19 * i); j += 3)
                {
                    for (g = 0; g < 4; g++)
                    {
                        //if (g == 0)
                        {
                            isOnline = false;
                            name = sheet[$"{GetInfo[g].Name}{j}"].ToString();
                            countSplit = 0;

                            if (name != "")
                            {
                                starttime = sheet[$"{GetInfo[g].Time}{j}"].ToString();
                                teacher = sheet[$"{GetInfo[g].Teacher}{j + 2}"].ToString();
                                description = sheet[$"{GetInfo[g].Descriprion}{j}"].ToString();
                                location = sheet[$"{GetInfo[g].Location}{j}"].ToString();
                                starttime = starttime.Replace('-', ':');
                                endtime = endTime(starttime);

                                if (location == "")
                                {
                                    location = @"calendar.google.com/calendar";
                                    isOnline = true;
                                }

                                if (description == "")
                                {
                                    for (l = g; l < 4; l++)
                                    {
                                        description = sheet[$"{GetInfo[l].Descriprion}{j}"].ToString();
                                        if (description != "")
                                        {
                                            while (countSplit >= 0)
                                            {
                                                lessons[l].group.Add(new lesson
                                                {
                                                    starttime = starttime,
                                                    name = name,
                                                    teacher = teacher.ToString(),
                                                    description = description,
                                                    location = location.ToString(),
                                                    id = currentIndex++.ToString(),
                                                    type = 0,
                                                    endtime = endtime,
                                                    isOnline = isOnline
                                                });
                                                l--;
                                                countSplit--;
                                            }
                                        }
                                        else
                                        {
                                            countSplit++;
                                        }

                                        if (countSplit == -1)
                                        {
                                            break;
                                        }
                                    }
                                }
                                else
                                {
                                    lessons[g].group.Add(new lesson
                                    {
                                        starttime = starttime,
                                        name = name,
                                        teacher = teacher.ToString(),
                                        description = description,
                                        location = location.ToString(),
                                        id = currentIndex++.ToString(),
                                        type = 0,
                                        endtime = endtime,
                                        isOnline = isOnline
                                    });
                                }
                            }
                        }
                    }
                }

                day = day.Split(' ')[1];
                DateTime d = Convert.ToDateTime(day).ToUniversalTime();
                day = d.ToString("o");

                Days.Add(new masDays());

                for (g = 0; g < 4; g++)
                {
                    Days[i].group.Add(new days
                    {
                        date = day,
                        lessons = lessons[g].group
                    });
                }
            }

            List<days> PI20 = new List<days>();
            List<days> KN20 = new List<days>();
            List<days> EBI20 = new List<days>();
            List<days> IN20 = new List<days>();
            for (i = 0; i < Days.Count; i++)
            {
                PI20.Add(new days()
                {
                    date = Days[i].group[0].date,
                    lessons = Days[i].group[0].lessons,
                });
                KN20.Add(new days()
                {
                    date = Days[i].group[1].date,
                    lessons = Days[i].group[1].lessons,
                });
                EBI20.Add(new days()
                {
                    date = Days[i].group[3].date,
                    lessons = Days[i].group[3].lessons,
                });
                IN20.Add(new days()
                {
                    date = Days[i].group[2].date,
                    lessons = Days[i].group[2].lessons,
                });

            }

            Shedule shedule = new Shedule();
            Groups ListLessons = new Groups()
            {
                PI20 = PI20,
                KN20 = KN20,
                EBI20 = EBI20,
                IN20 = IN20
            };

            shedule.shedule = ListLessons;


            var jsonString = JsonSerializer.Serialize(shedule);
            var options = new JsonSerializerOptions
            {
                Encoder = JavaScriptEncoder.Create(UnicodeRanges.BasicLatin, UnicodeRanges.Cyrillic),
                WriteIndented = true
            };

            jsonString = JsonSerializer.Serialize(shedule, options);
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

        public class GetCellInfo
        {
            public string Time { get; set; }
            public string Name { get; set; }
            public string Descriprion { get; set; }
            public string Teacher { get; set; }
            public string Location { get; set; }
        }

        public class Shedule
        {
            public Groups shedule { get; set; }
        }

        public class masDays
        {
            public List<days> group { get; set; }

            public masDays()
            {
                group = new List<days>();
            }
        }

        public class masLessons
        {
            public List<lesson> group { get; set; }

            public masLessons()
            {
                group = new List<lesson>();
            }
        }

        public class Groups
        {
            public List<days> PI20 { get; set; }
            public List<days> KN20 { get; set; }
            public List<days> EBI20 { get; set; }
            public List<days> IN20 { get; set; }
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
    }
}
