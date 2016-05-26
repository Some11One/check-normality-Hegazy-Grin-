using System;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.IO;
using System.Text;
using Excel;
using System.Data;
using System.Web.Helpers; // Custom library is used to help read xls/xlsx files
using CourseWork.Models;
using System.Collections.Generic;

namespace CourseWork.Controllers
{
    public class UploadController : Controller
    {

        // Base Index controller - when "Upload" page first called
        public ActionResult Index()
        {
            var model = new Models.UploadModels(); // Creating an model object
            @ViewBag.ErrorMessage = ""; // No error
            return View(model);

        }

        // Additional Index controller - handling file upload
        [HttpPost]
        public ActionResult Index(HttpPostedFileBase file)
        {
            try // Checking if there is any file 
            {
                // Checking file extension
                var allowedExtensions = new[] { ".xls", ".xlsx", ".txt", ".csv" };
                var extension = Path.GetExtension(file.FileName);
                if (!allowedExtensions.Contains(extension))
                {
                    @ViewBag.ErrorMessage = "Недопустимое расширение файла!";
                    return View();
                }
                byte[] uploadedFile = new byte[file.ContentLength];
                // Reading file into a byte array
                file.InputStream.Read(uploadedFile, 0, uploadedFile.Length);
                string uploadResult; // string form of uploaded file
                #region .xls/xlsx extension
                if (extension == ".xlsx" || extension == ".xls")
                {
                    IExcelDataReader excelReader;
                    if (extension == ".xls")
                    {
                        // Reading from a binary Excel file ('97-2003 format; *.xls)
                        excelReader = ExcelReaderFactory.CreateBinaryReader(file.InputStream);
                    }
                    else // xlsx extension
                    {
                        // Reading from a OpenXml Excel file (2007 format; *.xlsx)
                        excelReader = ExcelReaderFactory.CreateOpenXmlReader(file.InputStream);
                    }

                    // DataSet - The result of each spreadsheet will be created in the result.Tables
                    DataSet result = excelReader.AsDataSet();

                    // Free resources (IExcelDataReader is IDisposable)
                    excelReader.Close();

                    int studentsNumber; // Students number = number of Rows
                    int questionsNumber; // Questions number = number of Columns 
                    try
                    {
                        studentsNumber = result.Tables[0].Rows.Count;
                        questionsNumber = result.Tables[0].Columns.Count;
                    }
                    catch // Catch NullReferenseExeption if file is empty
                    {
                        @ViewBag.ErrorMessage = "Файл пуст";
                        return View();
                    }
                    string[,] testResults = new string[studentsNumber, questionsNumber]; // String that will represent data of the uploaded file
                    int row_no = 0; // row number
                    // Cheking for NullReferenceException
                    try
                    {
                        // Reading file into string array
                        while (row_no < result.Tables[0].Rows.Count)
                        {
                            for (int i = 0; i < result.Tables[0].Columns.Count; i++)
                            {
                                testResults[row_no, i] += result.Tables[0].Rows[row_no][i].ToString();
                                // Checking content of the file
                                if ((testResults[row_no, i] != "1") && (testResults[row_no, i] != "0"))
                                {
                                    if (testResults[row_no, i] == "")
                                    {
                                        throw new NullReferenceException();
                                    }
                                    string cacth = testResults[row_no, i];
                                    @ViewBag.ErrorMessage = String.Format("Результаты тестирования должны быть дихотомические!");
                                    return View();
                                }
                            }
                            row_no++;
                        }
                    }
                    catch (NullReferenceException)
                    {
                        @ViewBag.ErrorMessage = String.Format("Матрица ответов должна быть прямоугольная!");
                        return View();
                    }
                    // If file was empty
                    if (testResults == null)
                    {
                        @ViewBag.ErrorMessage = "Файл пуст";
                        return View();
                    }
                    return CheckNormality(testResults); // Calling method that checks data for normality


                }
                #endregion
                else // txt etension 
                    #region .txt Extension
                    if (extension == ".txt")
                    {
                        // Checking if file is empty
                        if (file.ContentLength == 0)
                        {
                            @ViewBag.ErrorMessage = "Файл пуст";
                            return View();
                        }
                        uploadResult = Encoding.UTF8.GetString(uploadedFile); // like this "1 1 1 0 1 0 1\r\n1 0 1 0 1 0..."
                        string[] separator = { "\r\n" };
                        // Remove all spaces
                        string uploadResultWS = uploadResult.Replace(" ", "");
                        // Create array of strings from uploaded file
                        string[] rows = uploadResultWS.Split(separator, StringSplitOptions.RemoveEmptyEntries);
                        int studentsNumber = rows.Length; // Students number = number of Rows
                        int questionsNumber = rows[0].Length; // Questions number = number of Columns 
                        // Checking if Matrix of answers is in form of rectangle
                        for (int i = 0; i < rows.Length; i++)
                        {
                            if (questionsNumber != rows[i].Length)
                            {
                                @ViewBag.ErrorMessage = String.Format("Матрица ответов должна быть прямоугольная!");
                                return View();
                            }
                        }
                        // Creating a table of answers
                        string[,] testResults = new string[studentsNumber, questionsNumber];
                        for (int i = 0; i < studentsNumber; i++)
                        {
                            for (int j = 0; j < questionsNumber; j++)
                            {
                                testResults[i, j] = rows[i][j].ToString();
                                // Checking content of the file
                                if ((testResults[i, j] != "1") && (testResults[i, j] != "0"))
                                {
                                    @ViewBag.ErrorMessage = String.Format("Результаты тестирования должны быть дихотомические!");
                                    return View();
                                }

                            }
                        }
                        return CheckNormality(testResults); // Calling method that checks data for normality
                    }
                    #endregion
                    else // csv extension of the file
                    #region .csv Extension
                    {
                        int studentsNumber; // Students number = number of Rows
                        int questionsNumber; // Questions number = number of Columns 
                        string[] rows; // Array of strings from uploaded file
                        try
                        {
                            uploadResult = Encoding.UTF8.GetString(uploadedFile); // like this "1;1;1;0;1;0;1\r\n1;0;1;0;1;0..."
                            string[] separator = { "\r\n" };
                            // Remove all semicolons
                            string uploadResultWithoutSpaces = uploadResult.Replace(";", "");
                            rows = uploadResultWithoutSpaces.Split(separator, StringSplitOptions.RemoveEmptyEntries);
                            studentsNumber = rows.Length;
                            questionsNumber = rows[0].Length;
                        }
                        catch // Catch NullReferenseExeption if file is empty
                        {
                            @ViewBag.ErrorMessage = "Файл пуст";
                            return View();
                        }
                        // Checking if Matrix of answers is in form of Recrangle
                        for (int i = 0; i < rows.Length; i++)
                        {
                            if (questionsNumber != rows[i].Length)
                            {
                                @ViewBag.ErrorMessage = String.Format("Матрица ответов должна быть прямоугольная!");
                                return View();
                            }
                        }
                        // Creating a table of answers
                        string[,] testResults = new string[studentsNumber, questionsNumber];
                        for (int i = 0; i < studentsNumber; i++)
                        {
                            for (int j = 0; j < questionsNumber; j++)
                            {
                                testResults[i, j] = rows[i][j].ToString();

                                if ((testResults[i, j] != "1") && (testResults[i, j] != "0"))
                                {
                                    @ViewBag.ErrorMessage = String.Format("Результаты тестирования должны быть дихотомические!");
                                    return View();
                                }

                            }
                        }
                        return CheckNormality(testResults); // Calling method that checks data for normality
                    }
                    #endregion
            }
            catch (NullReferenceException)
            {
                @ViewBag.ErrorMessage = "Сначала укажите файл для загрузки!";
                return View();
            }
        } // end of Index Action

        // Method that checks data for normality
        public ActionResult CheckNormality(string[,] testResults)
        {
            int[] data = new int[testResults.GetLength(0)]; // Will be an array of data from table
            Array.Sort(data);
            Array.Reverse(data);
            // Create data array by summing rows of testResult table
            for (int i = 0; i < testResults.GetLength(0); i++)
            {
                data[i] = 0;
                for (int j = 0; j < testResults.GetLength(1); j++)
                {
                    data[i] += int.Parse(testResults[i, j]);
                }
            }
            UploadModels.Answers = data; 
            double x; // Mean value
            double sum = 0; // Sum of array values
            double s; // Standart deviation
            int n = data.Length; // Number of elements
            for (int i = 0; i < n; i++)
            {
                sum += data[i];
            }
            x = (sum / n);
            double sumOfX = 0; // Сумма разниц аргументов и среднего значения
            for (int i = 0; i < n; i++)
            {
                sumOfX += Math.Pow((data[i] - x), 2);
            }
            s = Math.Sqrt(sumOfX * (1.0 / n));
            double[] y = new double[n]; // y - normalized data
            for (int i = 0; i < n; i++)
            {
                y[i] = (data[i] - x) / s;
            }
            double[] h = new double[n]; // Массив аппроксимаций квантилей
            for (int i = 1; i <= n; i++)
            {
                h[i - 1] = 4.91 * (Math.Pow((i / (n + 1.0)), 0.14) - Math.Pow(((n + 1.0 - i) / (n + 1)), 0.14));
            }
            double T1, T2; // Statistics
            double FinSum1 = 0, FinSum2 = 0; // For Stat calculation
            for (int i = 0; i < n; i++)
            {
                FinSum1 += Math.Abs(y[i] - h[i]);
                FinSum2 += Math.Pow(y[i] - h[i], 2);
            }
            T1 = FinSum1 / n;
            T2 = FinSum2 / n;
            // Calculating normal statistics for this n
            double T1Norm, T2Norm; // Normal statistics
            // Данные аппроксимированные формулы взяты из учебника Кобзара "Прикладная математическая статистика"
            T1Norm = 0.6027 - 0.1481 * Math.Log(n) + 0.0090 * (Math.Pow(Math.Log(n), 2));
            T2Norm = 0.0126 + (1.9227 / n) + (5.00677 / Math.Pow(n, 2));
            // Checking our statistics for normality
            bool flag; // Normal / not statistic
            if ((T1 < T1Norm) && (T2 < T2Norm)) flag = true;
            else flag = false;
            // Percentiles
            double[] perc = new double[testResults.GetLength(1)]; // Sum of students, who passed certain amount of answers
            // Calculating sum of answers for each column
            for (int i = 0; i < testResults.GetLength(1); i++)
           {
                perc[i] = 0;
                for (int j = 0; j < testResults.GetLength(0); j++)
                {
                    if (data[j] <= i + 1)
                        perc[i]++;
                }
                perc[i] = (perc[i] / testResults.GetLength(0)) * 100;
            }
            UploadModels.Perc = perc;
            // Saving data for later use
            UploadModels.Data = testResults; 
            UploadModels.Y = y;
            UploadModels.S = s;
            UploadModels.X = x;
            if (flag)
                return RedirectToAction("NormTestResults");
            else
                return RedirectToAction("NeNormTestResults");
        } // end of CheckForNormality

        // Create diagrams, if data is normal
        public ActionResult NormTestResults()
        {
            string[,] data = UploadModels.Data; // data from uploaded file, table of N students and M questions
            int M = data.GetLength(1); // Number of questions
            int N = data.GetLength(0); // Number of students
            int[] answers = new int[UploadModels.Answers.Length]; // Amount of correct answ for each student
            UploadModels.Answers.CopyTo(answers, 0);
            TableModels[] tableArray = new TableModels[N];
            for (int i = 0; i < N; i++)
            {
                tableArray[i] = new TableModels();
                tableArray[i].StudID = i + 1;
                tableArray[i].RawScore = answers[i];
                tableArray[i].Z = UploadModels.Y[i];
                tableArray[i].T = 50 + 10 * double.Parse(string.Format("{0:f1}", UploadModels.Y[i]));
                for (int j = 0; j < M; j++)
                {
                    if (answers[i] == j + 1) // See what percentage assosiates with result of a student
                    {
                        tableArray[i].Perc = UploadModels.Perc[j];
                    }
                }
            }
            // Sort Data
            if (SettingsModels.Sort) Array.Sort(answers, tableArray);
            ViewBag.Data = tableArray; 
            TableModels.tableArray = tableArray;
            return View();
        } // end of NormTestResults()

        // Create diagrams, if data isn't normal
        public ActionResult NeNormTestResults()
        {
            string[,] data = UploadModels.Data; // data from uploaded file, table of N students and M questions
            int M = data.GetLength(1); // Number of questions
            int N = data.GetLength(0); // Number of students
            int[] answers = new int[UploadModels.Answers.Length]; // Amount of correct answ for each student
            UploadModels.Answers.CopyTo(answers, 0);
            TableModels[] tableArray = new TableModels[N];
            for (int i = 0; i < N; i++)
            {
                tableArray[i] = new TableModels();
                tableArray[i].StudID = i + 1;
                tableArray[i].RawScore = answers[i];
            }
            // Sort Data
            if(SettingsModels.Sort) Array.Sort(answers,tableArray);
            ViewBag.Data = tableArray;
            TableModels.tableArray = tableArray;
            return View();
        } // end of NeNormTestResults()

        #region Таблицы

        // Table of raw grades
        public ActionResult absChart()
        {
            string[,] data = UploadModels.Data; // data from uploaded file, table of N students and M questions
            int M = data.GetLength(1); // Number of questions
            int N = data.GetLength(0); // Number of students
            int[] answers = new int[UploadModels.Answers.Length]; // Amount of correct answ for each student
            UploadModels.Answers.CopyTo(answers, 0);
            int[] index = new int[N]; // Students
            for (int i = 0; i < N; i++)
            {
                index[i] = i + 1;
            }
            var dataChart = new Chart(width: 550, height: 400, theme: ChartTheme.Blue)
        .AddTitle("Данные загрузки").SetXAxis("Участники").SetYAxis("Кол-во ответов")
        .AddSeries(
        name: "data",
            xValue: index,
            yValues: answers).GetBytes();
            return base.File(dataChart, "image/png");
        }

        // Draw table with data from file
        public ActionResult dataChart()
        {
            string[,] data = UploadModels.Data; // data from uploaded file, table of N students and M questions
            double x = UploadModels.X; // Mean Value
            double s = UploadModels.S; // Middle Deviation
            int Amount = 3; // Amount of points/2
            int M = data.GetLength(1); // Number of questions
            int N = data.GetLength(0); // Number of students
            int[] answers = new int[UploadModels.Answers.Length]; // Amount of correct answ for each student
            UploadModels.Answers.CopyTo(answers, 0);
            int[] pass = new int[2 * Amount + 1]; // Number of st, who passed certain number of correct answers
            string[] index = new string[2 * Amount + 1]; // number of question
            // Find how many students passed certain number of correct answers
            int count = Amount;
            int stepCount = 0;
            for (double i = x; count < (Amount * 2 + 1); i += s)
            {
                if (i == x) index[count] = "Среднее";
                else
                    index[count] = string.Format("+{0}S", stepCount);
                pass[count] = 0;
                // find how many students passed this amount of correct answers\
                for (int j = 0; j < answers.Length; j++)
                {
                    if (answers[j] == (int)i)
                    {
                        pass[count]++;
                    }
                }
                count++;
                stepCount++;
            }
            count = Amount - 1;
            stepCount = 1;
            for (double i = x - s; count > -1; i -= s)
            {
                index[count] = string.Format("-{0}S", stepCount);
                pass[count] = 0;
                // find how many students passed this amount of correct answers\
                for (int j = 0; j < answers.Length; j++)
                {
                    if (answers[j] == (int)i)
                        pass[count]++;
                }
                count--;
                stepCount++;
            }
            var dataChart = new Chart(width: 550, height: 400, theme: ChartTheme.Blue)
        .AddTitle("Данные загрузки").SetXAxis("Количество вопросов").SetYAxis("Количество участников")
        .AddSeries(
        name: "data",
            xValue: index,
            yValues: pass).GetBytes();
            return base.File(dataChart, "image/png");
        }

        // Z scale
        public ActionResult zScale()
        {
            double Ymin = double.Parse(string.Format("{0:f1}", UploadModels.Y[0]));
            double Ymax = double.Parse(string.Format("{0:f1}", UploadModels.Y[UploadModels.Y.Length - 1]));
            int[] index = new int[UploadModels.Y.Length]; // y scale values
            for (int i = 0; i < index.Length; i++)
            {
                index[i] = i+1;
            }
            var dataChart = new Chart(width: 550, height: 400, theme: ChartTheme.Blue)
        .AddTitle("Z-Шкала").SetXAxis("Участники").SetYAxis("Z-балл", Ymin - 0.1, Ymax + 0.1)
        .AddSeries(
        name: "z",
            xValue: index,
            yValues: UploadModels.Y).GetBytes();
            return base.File(dataChart, "image/png");
        }

        // T scale
        public ActionResult tScale()
        {
            int[] index = new int[UploadModels.Y.Length]; // y scale values
            double[] t = new double[UploadModels.Y.Length]; //values for t scale
            for (int i = 0; i < index.Length; i++)
            {
                index[i] = i+1;
                t[i] = 50 + 10 * double.Parse(string.Format("{0:f1}", UploadModels.Y[i]));
            }
            var dataChart = new Chart(width: 550, height: 400, theme: ChartTheme.Blue)
        .AddTitle("T-Шкала").SetXAxis("Участники").SetYAxis("T-балл")
        .AddSeries(
        name: "t",
            xValue: index,
            yValues: t).GetBytes();
            return base.File(dataChart, "image/png");
        }

        // Percente scale
        public ActionResult percScale()
        {
            string[,] data = UploadModels.Data; // data from uploaded file, table of N students and M questions
            double x = UploadModels.X; // Mean Value
            double s = UploadModels.S; // Middle Deviation
            int M = data.GetLength(1); // Number of questions
            int N = data.GetLength(0); // Number of students
            int[] answers = new int[UploadModels.Answers.Length]; // Amount of correct answ for each student
            UploadModels.Answers.CopyTo(answers, 0);
            int[] index = new int[M]; // number of question
            double[] sum = UploadModels.Perc; // Sum of students, who passed certain amount of answers
            var dataChart = new Chart(width: 550, height: 400, theme: ChartTheme.Blue)
        .AddTitle("Шкала Процентилей").SetXAxis("Балл").SetYAxis("Процентиль")
        .AddSeries(
        name: "perc",
            xValue: index,
            yValues: sum).GetBytes();
            return base.File(dataChart, "image/png");
        }

        // 25-quantile
        public ActionResult FirstQuant()
        {
            string[,] data = UploadModels.Data; // data from uploaded file, table of N students and M questions
            double x = UploadModels.X; // Mean Value
            double s = UploadModels.S; // Middle Deviation
            int M = data.GetLength(1); // Number of questions
            int N = data.GetLength(0); // Number of students       
            int[] answers = new int[UploadModels.Answers.Length]; // Amount of correct answ for each student
            UploadModels.Answers.CopyTo(answers, 0);
            double[] perc = UploadModels.Perc;
            List<int> students = new List<int>();
            List<double> res = new List<double>();
            for (int i = 0; i < N; i++)
            {
                for (int j = 0; j < M; j++)
                {
                    if (answers[i] == j+1) // See what percentage assosiates with result of a student
                    {
                        if (perc[j] <= 25) // Check if it is in first 25%
                        {
                            students.Add(i+1);
                            res.Add(perc[j]);
                        }
                    }
                    else if (answers[i]==0)
                    {
                        students.Add(i + 1);
                        res.Add(0);
                    }
                }
            }
            try
            {
            var dataChart = new Chart(width: 550, height: 400, theme: ChartTheme.Blue)
        .AddTitle("Шкала 1 квартиля").SetXAxis("Участник", students[0]-1, students[res.Count - 1]+1).SetYAxis("Процентиль")
        .AddSeries(
        name: "quant",
            xValue: students,
            yValues: res).GetBytes();
            return base.File(dataChart, "image/png");
            }
            catch
            {
                var dataChart = new Chart(width: 550, height: 400, theme: ChartTheme.Blue)
        .AddTitle("Шкала 1 квартиля - участников с процентилем ниже 25 нет").SetXAxis("Таких участников нет")
        .AddSeries(
        name: "quant",
            xValue: students,
            yValues: res).GetBytes();
                return base.File(dataChart, "image/png");
            }
        }

        // 75-quantile
        public ActionResult ThirdQuant()
        {
            string[,] data = UploadModels.Data; // data from uploaded file, table of N students and M questions
            double x = UploadModels.X; // Mean Value
            double s = UploadModels.S; // Middle Deviation
            int M = data.GetLength(1); // Number of questions
            int N = data.GetLength(0); // Number of students        
            int[] answers = new int[UploadModels.Answers.Length]; // Amount of correct answ for each student
            UploadModels.Answers.CopyTo(answers, 0);
            double[] perc = UploadModels.Perc;
            List<int> students = new List<int>();
            List<double> res = new List<double>();
            for (int i = 0; i < N; i++)
            {
                for (int j = 0; j < M; j++)
                {
                    if (answers[i] == j+1) // See what percentage assosiates with result of a student
                    {
                        if (perc[j] >= 75) // Check if it is in last 25%
                        {
                            students.Add(i+1);
                            res.Add(perc[j]);
                        }
                    }
                }
            }

            try
            {
                var dataChart = new Chart(width: 550, height: 400, theme: ChartTheme.Blue)
            .AddTitle("Шкала 3 квартиля").SetYAxis("Процентиль",70).SetXAxis("Участник", students[0] - 1, students[res.Count - 1] + 1)
            .AddSeries(
            name: "quant",
                xValue: students,
                yValues: res).GetBytes();
                return base.File(dataChart, "image/png");
            }
            catch
            {
                var dataChart = new Chart(width: 550, height: 400, theme: ChartTheme.Blue)
            .AddTitle("Шкала 3 квартиля - участников с процентилем больше 75 нет").SetXAxis("Таких участников нет")
            .AddSeries(
            name: "quant",
                xValue: students,
                yValues: res).GetBytes();
                return base.File(dataChart, "image/png");
            }
        }
        #endregion

        #region save
        // Method to save results of normal data analysis
        public ActionResult SaveNormResult()
        {
            TableModels[] resultsArray = TableModels.tableArray;
            StringBuilder result = new StringBuilder();

            result.Append("Participant;Score;Z-score;T-score;Percentile");
            result.AppendLine();
            for (int i = 0; i < resultsArray.Length; i++)
            {
                string line;
                line = string.Format("{0};{1};{2:f1};{3:f1};{4}", resultsArray[i].StudID, resultsArray[i].RawScore, resultsArray[i].Z, resultsArray[i].T, resultsArray[i].Perc);
                result.Append(line);
                result.AppendLine();
            }
            return File(new System.Text.UTF8Encoding().GetBytes(result.ToString()), "csv/text", "NormResults.csv");
        }

        public ActionResult SaveNeNormResult()
        {
            TableModels[] resultsArray = TableModels.tableArray;
            StringBuilder result = new StringBuilder();

            result.Append("Participant;Score");
            result.AppendLine();
            for (int i = 0; i < resultsArray.Length; i++)
            {
                string line;
                line = string.Format("{0};{1};", resultsArray[i].StudID, resultsArray[i].RawScore);
                result.Append(line);
                result.AppendLine();
            }
            return File(new System.Text.UTF8Encoding().GetBytes(result.ToString()), "csv/text", "AbNormResults.csv");
        }

        
        #endregion

    } // end of UploadController class
} // end of namespace CourseWork.Controlles