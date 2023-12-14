using OfficeOpenXml;
using System;
using System.Data;
using System.Runtime.CompilerServices;
using System.Security;
using System.Threading.Channels;

class Program
{
    static void Main()
    {
        
        Console.WriteLine("Hello, Welcome to Anro's activity tracker");
        Console.WriteLine("Please ensure that all files submissions are .xlsx");

        //Import
        Console.WriteLine("What is your file path?");

        var filePath = Console.ReadLine();

        // Check if the file exists
        if (!File.Exists(filePath))
        {
            Console.WriteLine("The Excel file does not exist.");
            return;
        }
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        // Load the Excel file
        using (var package = new ExcelPackage(new FileInfo(filePath)))
        {

            // Assuming the Excel file has a worksheet named "Sheet1"
            var worksheet = package.Workbook.Worksheets["Sheet1"];
            var rows = worksheet.Dimension.Rows;
            var cols = worksheet.Dimension.Columns;

            int TimestampMap = 0;
            int AccelXMap = 0;
            int AccelYMap = 0;
            int AccelZMap = 0;
            int ActivityMap = 0;
            int ConfidenceMap = 0;
            int TrueHeadingMap = 0;
            int AccelUserXMap = 0;
            int AccelUserYMap = 0;
            int AccelUserZMap = 0;
            int GyroXMap = 0;
            int GyroYMap = 0;
            int GyroZMap = 0;

            //Create mapping for tables accoring to name and column number
            for (int i = 1; i < cols + 1; i++)
            {
                object ColName = worksheet.Cells[1, i].Value;
                if  ((ColName.ToString().ToLower()) == "timestamp")
                {
                    TimestampMap = i;
                }
                else if ((ColName.ToString().ToLower()) == "accelx(g)")
                {
                     AccelXMap = i;
                }
                else if ((ColName.ToString().ToLower()) == "accely(g)")
                {
                    AccelYMap = i;
                }
                else if ((ColName.ToString().ToLower()) == "accelz(g)")
                {
                    AccelZMap = i;
                }
                else if ((ColName.ToString().ToLower()) == "activitytype")
                {
                    ActivityMap = i;
                }
                else if ((ColName.ToString().ToLower()) == "activityconfidence")
                {
                    ConfidenceMap = i;
                }
                else if ((ColName.ToString().ToLower()) == "trueheading")
                {
                    TrueHeadingMap = i;
                }
                else if ((ColName.ToString().ToLower()) == "acceluserx(g)")
                {
                    AccelUserXMap = i;
                }
                else if ((ColName.ToString().ToLower()) == "accelusery(g)")
                {
                    AccelUserYMap = i;
                }
                else if ((ColName.ToString().ToLower()) == "acceluserz(g)")
                {
                    AccelUserZMap = i;
                }
                else if ((ColName.ToString().ToLower()) == "gyrox(rad/s)")
                {
                    GyroXMap = i;
                }
                else if ((ColName.ToString().ToLower()) == "gyroy(rad/s)")
                {
                    GyroYMap = i;
                }
                else if ((ColName.ToString().ToLower()) == "gyroz(rad/s)")
                {
                    GyroZMap = i;
                }
            } 
            
            //Define arrays
            string[] Timestamp = new string[rows + 1];
            string[] Activities = new string[rows + 1];
            string[] Confidence = new string[rows + 1];
            string[] TrueHeading = new string[rows];
            double[] TrueHeadingDouble = new double[rows];
            string[] GyroX = new string[rows + 1];
            string[] GyroY = new string[rows + 1];
            string[] GyroZ = new string[rows + 1];
            string[] AccelX = new string[rows + 1];
            string[] AccelY = new string[rows + 1];
            string[] AccelZ = new string[rows + 1];
            string[] AccelUserX = new string[rows + 1];
            string[] AccelUserY = new string[rows + 1];
            string[] AccelUserZ = new string[rows + 1];

            string[] SecActivities = new string[rows + 1];
            string[] MinActivities = new string[(rows + 1)];
            string[] HrActivities = new string[(rows + 1)];

            //using mapping created above and create data for arrays
            //loop for time
            for (var row = 1; row < Timestamp.Length; row++)
            {
                object currentTime = worksheet.Cells[row, TimestampMap].Value;
                Timestamp[row] = currentTime.ToString();
            }            

            //loop for true heading
            for (var row = 0; row < TrueHeading.Length; row++)
            {
                object current = worksheet.Cells[row + 1, TrueHeadingMap].Value;
                TrueHeading[row] = current.ToString();
            }

            //loop for GyroX
            for (var row = 1; row < GyroX.Length; row++)
            {
                object currentTime = worksheet.Cells[row, GyroXMap].Value;
                GyroX[row] = currentTime.ToString();
            }

            //loop for GyroY
            for (var row = 1; row < GyroY.Length; row++)
            {
                object currentTime = worksheet.Cells[row, GyroYMap].Value;
                GyroY[row] = currentTime.ToString();
            }

            //loop for GyroZ
            for (var row = 1; row < GyroZ.Length; row++)
            {
                object currentTime = worksheet.Cells[row, GyroZMap].Value;
                GyroZ[row] = currentTime.ToString();
            }

            //loop for AccelX
            for (var row = 1; row < AccelX.Length; row++)
            {
                object currentTime = worksheet.Cells[row, AccelXMap].Value;
                AccelX[row] = currentTime.ToString();
            }

            //loop for AccelY
            for (var row = 1; row < AccelY.Length; row++)
            {
                object currentTime = worksheet.Cells[row, AccelYMap].Value;
                AccelY[row] = currentTime.ToString();
            }

            //loop for AccelZ
            for (var row = 1; row < AccelZ.Length; row++)
            {
                object currentTime = worksheet.Cells[row, AccelZMap].Value;
                AccelZ[row] = currentTime.ToString();
            }

            //loop for AccelUserX
            for (var row = 1; row < AccelUserX.Length; row++)
            {
                object currentTime = worksheet.Cells[row, AccelUserXMap].Value;
                AccelUserX[row] = currentTime.ToString();
            }

            //loop for AccelUserY
            for (var row = 1; row < AccelUserY.Length; row++)
            {
                object currentTime = worksheet.Cells[row, AccelUserYMap].Value;
                AccelUserY[row] = currentTime.ToString();
            }

            //loop for AccelUserZ
            for (var row = 1; row < AccelUserZ.Length; row++)
            {
                object currentTime = worksheet.Cells[row, AccelUserZMap].Value;
                AccelUserZ[row] = currentTime.ToString();
            }

            Console.WriteLine("================================================");

            int count = 0;
            int countConvert = 0;
            var secCount = 0;
            var minCount = 0;
            var hrCount = 0;

            foreach (var item in TrueHeading)
            {
                
                string current = TrueHeading[count];
                double prev = 0;
                string activity = "N/A";

                double conversion = 0;

                if (count >= 1)
                {
                    conversion = Convert.ToDouble(current);
                    TrueHeadingDouble[countConvert] = conversion;
                   
                }

                double currentValue = TrueHeadingDouble[countConvert];

                if (countConvert == 0)
                {
                     prev = currentValue;
                }
                else
                {
                     prev = TrueHeadingDouble[countConvert - 1];
                }
                countConvert++;


                if ((prev < (currentValue + 2.5)) && (prev > (currentValue - 2.5)))
                {
                    activity = "Stationary";
                }
                else if ((prev > (currentValue + 2.5)) || (prev < (currentValue - 2.5)))
                {
                    activity = "Moving";
                }
                else
                {
                    activity = "N/A";
                }

                if (count == 1 )
                {
                    Activities[count] = "Activities";
                }
                else
                {
                    Activities[count] = activity;
                }

                count++;

                secCount++;

                if (secCount == 60)
                {
                    secCount = 0;
                    minCount++;
                    string currentTime = "";
                    for (int i = 0; i <= 60; i++)
                    {
                        currentTime = Activities[i];
                        if (currentTime == "Moving")
                        {
                            SecActivities[i] = "Moving";
                        }
                    }

                    var activityTracker = 0;
                    foreach (var sec in SecActivities)
                    {
                        if (sec == "Moving")
                        {
                            activityTracker++;
                        }
                    }

                    if (activityTracker > 30)
                    {
                        MinActivities[secCount] = "Moving";
                    }
                    else
                    {
                        MinActivities[secCount] = "Stationary";
                    }
                    
                }

                if (minCount == 60)
                {
                    minCount = 0;
                    hrCount++;
                    var activityTrackerMin = 0;
                    foreach (var Min in MinActivities)
                    {
                        if (Min == "Moving")
                        {
                            activityTrackerMin++;
                        }
                    }

                    if (activityTrackerMin > 30)
                    {
                        HrActivities[minCount] = "Moving";
                    }
                    else
                    {
                        HrActivities[minCount] = "Stationary";
                    }
                }
            }

            for (int i = 1; i <= rows; i++)
            {
                
                var time = Timestamp[i];
                var act = Activities[i];

                if (i == 1)
                {
                    Console.Write(time + "\t" + "\t");
                    Console.WriteLine(act);
                }
                else
                {
                    Console.Write(time + "\t");
                    Console.WriteLine(act);
                    Console.WriteLine("=");
                }

            }

            var MovingCountMin = 0;
            var StationCountMin = 0;
            var NACountMin = 0;
            Console.WriteLine();
            Console.WriteLine("=====================================================");
            Console.WriteLine("Activities broken down by the minutes");
            Console.WriteLine("=====================================================");

            if (minCount > 0)
            {
                foreach (var min in MinActivities)
                {

                    if (min == "Moving")
                    {
                        MovingCountMin++;
                        Console.WriteLine(min);
                    }
                    else if (min == "Stationary")
                    {
                        StationCountMin++;
                        Console.WriteLine(min);
                    }
                    else
                    {
                        NACountMin++;
                    }
                }
                if (MovingCountMin > (MovingCountMin + StationCountMin) / 2)
                {
                    Console.WriteLine("------------------------------------------------");
                    Console.WriteLine("You are moving enough minutes");
                    Console.WriteLine("------------------------------------------------");
                }
                else
                {
                    Console.WriteLine("------------------------------------------------");
                    Console.WriteLine("You are NOT moving enough minutes");
                    Console.WriteLine("------------------------------------------------");
                }
                var totalMin = MovingCountMin + StationCountMin;
                Console.WriteLine("Total min: " + totalMin);
                Console.WriteLine("Moving Count = " + MovingCountMin);
                Console.WriteLine("Stationary Count = " + StationCountMin);
            }
            else
            {
                Console.WriteLine("The date set dit not have enough records for an min");
            }





            var MovingCountHr = 0;
            var StationCountHr = 0;
            var NACountHr = 0;
            Console.WriteLine();
            Console.WriteLine("=====================================================");
            Console.WriteLine("Activities broken down by the hours");
            Console.WriteLine("=====================================================");

            if (hrCount > 0)
            {
                foreach (var Hr in HrActivities)
                {
                    if (Hr == "Moving")
                    {
                        MovingCountHr++;
                        Console.WriteLine(Hr);
                    }
                    else if (Hr == "Stationary")
                    {
                        StationCountHr++;
                        Console.WriteLine(Hr);
                    }
                    else
                    {
                        NACountHr++;
                    }
                }
                if (MovingCountHr > (MovingCountHr + StationCountHr) / 2)
                {
                    Console.WriteLine("------------------------------------------------");
                    Console.WriteLine("You are moving enough hours");
                    Console.WriteLine("------------------------------------------------");
                }
                else
                {
                    Console.WriteLine("------------------------------------------------");
                    Console.WriteLine("You are NOT moving enough hours");
                    Console.WriteLine("------------------------------------------------");
                }
                var totalHr = MovingCountMin + StationCountMin;
                Console.WriteLine("Total Hr: " + totalHr);
                Console.WriteLine("Moving Count = " + MovingCountHr);
                Console.WriteLine("Stationary Count = " + StationCountHr);
            }
            else
            {
                Console.WriteLine("The date set dit not have enough records for an hour");
            }

           

            Console.ReadLine();
        }       

    }
}
