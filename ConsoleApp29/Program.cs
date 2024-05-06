using System;
using System.IO;
using System.Net;
using System.Net.Http;
using System.Threading;
using System.Collections.Generic;
using System.Globalization;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using static System.Net.Mime.MediaTypeNames;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Xml.Linq;
using System.Drawing.Printing;
using System.Text;
using ScrapySharp.Network;
using System.Text.RegularExpressions;

namespace CheckPrice
{
    class Program
    {
        struct CellStruct
        {
            public object cellTotalPrice;
            public object cellDate;
           public CellStruct(object Price, object Date)
           { 
                cellTotalPrice = Price;
                cellDate = Date;
           }
        }

        public static object[] FinalizePrice(object[][] objects)
        {
            object[] prices = objects[0];

            return prices;
        }

        public static object[] FinalizeDate(object[][] objects)
        {
            object[] dates = objects[1];

            return dates;
        }

        public static object[][] Universalis(string folderPath, int? columV = 2, int? columH = 4)
        {
            List<CellStruct> cells = new List<CellStruct>();
            List<object> cellDate = new List<object>();
            List<object> cellTotalPrice = new List<object>();
            foreach (string filePath in Directory.GetFiles(folderPath, "*.xlsx"))
            {
                // Открываем файл Excel
                using (ExcelPackage package = new ExcelPackage(new FileInfo(filePath)))
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets[0]; // изменить на нужный лист

                    cells.Add(new CellStruct(float.Parse(worksheet.Cells[(int)columV, (int)columH].Value.ToString()), worksheet.Cells[2, 5].Value)); // value / date
                }
            }

            List<CellStruct> sortedCells = cells.OrderBy(cell => DateTime.Parse(cell.cellDate.ToString())).ToList();
            //Console.WriteLine("---0000---");

            for (int i = 0; i < sortedCells.Count; i++)
            {
                cellTotalPrice.Add(sortedCells[i].cellTotalPrice);
                cellDate.Add(sortedCells[i].cellDate);
            }
            object[][] obj = new object[2][];
            obj[0] = new object[1][];
            obj[0] = cellTotalPrice.Cast<object>().ToArray();
            obj[1] = new object[1][];
            obj[1] = cellDate.Cast<object>().ToArray();

            return obj;
        }

        public static object[] LoadDates(string folderPath)
        {
            int count = Directory.GetFiles(folderPath, "*.xlsx").Length;

            List<object> cellDate = new List<object>();

            foreach (string filePath in Directory.GetFiles(folderPath, "*.xlsx"))
            {
                // Открываем файл Excel
                using (ExcelPackage package = new ExcelPackage(new FileInfo(filePath)))
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets[0]; // изменить на нужный лист

                    cellDate.Add(worksheet.Cells[2, 5].Value);
                }
            }
            return cellDate.Cast<object>().ToArray();
        }
        public static object[] LoadPrices(string folderPath)
        {
            int count = Directory.GetFiles(folderPath, "*.xlsx").Length;
            List<object> cellTotalPrice = new List<object>();

            foreach (string filePath in Directory.GetFiles(folderPath, "*.xlsx"))
            {
                // Открываем файл Excel
                using (ExcelPackage package = new ExcelPackage(new FileInfo(filePath)))
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets[0]; // изменить на нужный лист

                    cellTotalPrice.Add(float.Parse(worksheet.Cells[2, 4].Value.ToString()));
                }
            }

            return cellTotalPrice.Cast<object>().ToArray();
        }

        public static void BuildDiagram(string leftColumName, string rightColumName, object[] arr1, object[] arr2, int? sizeW = 800, int? sizeH = 400)
        {
            string Base = $"https://yequalx.com/ru/chart/line/{leftColumName},{rightColumName};";
            for (int i = 0; i < arr1.Length; i++)
            {
                if(i!= arr1.Length -1)
                {
                    Base += $"{arr1[i].ToString().Replace(",", ".")},{arr2[i].ToString().Replace(",", ".")};";
                }
                else
                {
                    Base += $"{arr1[i].ToString().Replace(",", ".")},{arr2[i].ToString().Replace(",", ".")}";
                }
            }
            Base += $"#w:{sizeW};h:{sizeH};c:4285F4";

            //Console.WriteLine(Base);
            ProcessStartInfo startInfo = new ProcessStartInfo
            {
                FileName = Base, //  имя файла в URL-адрес
                UseShellExecute = true // оболочка системы для запуска процесса
            };
            Process.Start(startInfo);
        }

        public static async Task LoadData(PriceController priceController)
        {
            string filePath = "C:\\Users\\admin\\Desktop\\Prices\\NEW\\info.txt";

            using (StreamReader reader = new StreamReader(filePath))
            {
                string fileContent = reader.ReadToEnd();
                string[] lines = fileContent.Split('^');

                if (lines != null)
                {
                    foreach (string line in lines)
                    {
                        string[] objects = line.Trim().Split('|');

                        Console.WriteLine(objects.Length);
                        for (int i = 0; i < objects.Length / 3; i++)
                        {

                            Console.WriteLine(objects[i] + " / " + objects[i + 1] + " / " + objects[i + 2]);
                            if(!priceController.CheckFile())
                            await priceController.AddProduct(objects[i], objects[i + 1], objects[i + 2]);
                            else
                            await priceController.LoadProduct(objects[i], objects[i + 1], objects[i + 2]);
                        }
                    }
                }
            }
        }

        static async Task Main(string[] args)
        {
            //args = new string[1];
            //args[0] = "diagram_all";
            //Console.WriteLine("Programm");
            string histortFilePath = "C:\\Users\\admin\\Desktop\\Prices\\NEW\\History\\";
            PriceController priceController = new PriceController(histortFilePath);

            if (args.Length > 0)
            {
                if (args[0] == "diagram") { BuildDiagram("Date", "Prices", FinalizeDate(Universalis(histortFilePath)), FinalizePrice(Universalis(histortFilePath))); }
                else if (args[0] == "diagram_all")
                {
                    await LoadData(priceController);
                    Console.WriteLine("All: ");
                    List<Product> productList = priceController.GetProductList();
                    Console.WriteLine(productList.Count);
                    for (int i = 0; i < productList.Count; i++)
                    {
                        Console.WriteLine($"{i + 1}. {productList[i].productType} | {productList[i].name}");
                        BuildDiagram("Date", $"{productList[i].productType}", FinalizeDate(Universalis(histortFilePath)), FinalizePrice(Universalis(histortFilePath, i+2, 2)));
                    }
                }
                else if (args[0] == "menu")
                {
                    string filePath = "C:\\Users\\admin\\Desktop\\Prices\\NEW\\info.txt";

                    
                    using (StreamReader reader = new StreamReader(filePath))
                    {
                        string fileContent = reader.ReadToEnd();
                        string[] lines = fileContent.Split('^');

                        if (lines != null)
                        {
                            foreach (string line in lines)
                            {
                                string[] objects = line.Trim().Split('|');

                                Console.WriteLine(objects.Length);
                                for (int i = 0; i < objects.Length / 3; i++)
                                {

                                    Console.WriteLine(objects[i] + " / " + objects[i + 1] + " / " + objects[i + 2]);
                                    if(!priceController.CheckFile())
                                    await priceController.LoadProduct(objects[i], objects[i + 1], objects[i + 2]);
                                    else
                                    await priceController.AddProduct(objects[i], objects[i + 1], objects[i + 2]);
                                }
                            }
                        }
                    }
                    
                     
                    int choose = 0, choose2 = 0;

                    while (true)
                    {
                        choose = 0;
                        Console.WriteLine("Menu: ");
                        Console.WriteLine("1. Add New Product");
                        Console.WriteLine("2. Show Tracked Products");
                        Console.WriteLine("3. Build Diagramm");

                        choose = int.Parse(Console.ReadLine());

                        if (choose == 1)
                        {
                            Console.Write("Enter Product Type: ");
                            string type = Console.ReadLine();
                            Console.WriteLine();

                            Console.Write("Enter Enter Product Name: ");
                            string name = Console.ReadLine();
                            Console.WriteLine();

                            Console.Write("Enter Enter Product Link: ");
                            string link = Console.ReadLine();
                            Console.WriteLine();

                            await priceController.AddProduct(type, name, link);

                            string filePath2 = "C:\\Users\\admin\\Desktop\\Prices\\NEW\\info.txt";
                            string data = $"{type}|{name}|{link}^";

                            using (StreamWriter writer = new StreamWriter(filePath2, true))
                            {
                                writer.WriteLine(data);
                            }

                        }
                        else if (choose == 2)
                        {
                            priceController.ShowProducts();
                        }
                        else if (choose == 3)
                        {
                            Console.WriteLine("1. Total Price Diagram");
                            Console.WriteLine("2. Choose Product to Buld Diagram");
                            choose2 = int.Parse(Console.ReadLine());

                            if (choose2 == 1)
                            {
                                BuildDiagram("Date", "Prices", FinalizeDate(Universalis(histortFilePath)), FinalizePrice(Universalis(histortFilePath)));
                            }
                            else if (choose2 == 2)
                            {
                                choose = 0;
                                List<Product> productList = priceController.GetProductList();

                                for (int i = 0; i < productList.Count; i++)
                                {
                                    Console.WriteLine($"{i + 1}. {productList[i].productType} | {productList[i].name}");
                                }
                                choose = int.Parse(Console.ReadLine());
                                BuildDiagram("Date", $"{productList[choose - 1].productType}", FinalizeDate(Universalis(histortFilePath)), FinalizePrice(Universalis(histortFilePath, choose + 1, 2)));
                            }

                        }
                    }
                    
                    

                }
            }

            if (priceController.CheckFile())
                Console.WriteLine("Already Exist");
            else
            {
                string filePath = "C:\\Users\\admin\\Desktop\\Prices\\NEW\\info.txt";

                using (StreamReader reader = new StreamReader(filePath))
                {
                    string fileContent = reader.ReadToEnd();
                    string[] lines = fileContent.Split('^');

                    if (lines != null)
                    {
                        foreach (string line in lines)
                        {
                            string[] objects = line.Trim().Split('|');

                            Console.WriteLine(objects.Length);
                            for (int i = 0; i < objects.Length / 3; i++)
                            {

                                Console.WriteLine(objects[i] + " / " + objects[i + 1] + " / " + objects[i + 2]);
                                await priceController.AddProduct(objects[i], objects[i + 1], objects[i + 2]);
                                
                            }
                        }
                    }
                }

                priceController.CalculateDiff();
                priceController.ShowProducts();

                await priceController.SavePricesToExcel();
            }

            

        }

    }

    class PriceController
    {
        private static HttpClient client;
        public static List<Product> productList;
        private static string filePath;
        public PriceController(string filepath)
        {
            client = new HttpClient();
            productList = new List<Product>();
            filePath = filepath + $"Prices[{DateTime.Now.ToString("dd-MM-yyyy")}].xlsx";
        }
        public List<Product> GetProductList()
        {
            return productList;
        }
        public object[] GetPriceArray()
        {
            object[] buf = new object[productList.Count];
            for (int i = 0; i < productList.Count; i++)
            {
                buf[i] = productList[i].price;
            }
            return buf;
        }
        public object[] GetTypeArray()
        {
            object[] buf = new object[productList.Count];
            for (int i = 0; i < productList.Count; i++)
            {
                buf[i] = productList[i].productType;
            }
            return buf;
        }
        
        public string GetFilePath()
        {
            return filePath;
        }

        static string GetLatestFilePath()
        {
            string filePath = @"C:\Users\admin\Desktop\Prices\NEW\History";
            Dictionary<DateTime, FileInfo> fileDict = new Dictionary<DateTime, FileInfo>();
            DateTime date;
            DateTime maxDate = DateTime.MinValue;

            DirectoryInfo directoryInfo = new DirectoryInfo(filePath);

            FileInfo[] fileInfo = directoryInfo.GetFiles("Prices*.xlsx");

            if (fileInfo.Length > 0)
            {
                foreach (FileInfo file in directoryInfo.GetFiles("Prices*.xlsx"))
                {
                    string fileNameWithoutExtension = Path.GetFileNameWithoutExtension(file.Name);

                    fileNameWithoutExtension = fileNameWithoutExtension.Replace("Prices[", "").Replace("]", "");
                    if (DateTime.TryParseExact(fileNameWithoutExtension, "dd-MM-yyyy", null, System.Globalization.DateTimeStyles.None, out date))
                    {
                        fileDict.Add(date, file);

                        if (date > maxDate)
                        {
                            maxDate = date;
                        }
                    }
                }

                if (fileDict.Count > 0)
                {
                    var predMaxDate = fileDict.Keys.Max();
                    return fileDict[predMaxDate].FullName;
                }
                else
                {
                    return "null";
                }
            }
            else { return "null"; }

        }
        public async Task LoadProduct(string productType, string name, string link)
        {
            Product product = new Product();
            product.link = link;
            product.name = name;
            product.productType = productType;
            productList.Add(product);
        }
        public async Task AddProduct(string productType, string name, string link)
        {
            ScrapingBrowser browser = new ScrapingBrowser();
            Product product = new Product();

            product.link = link;
            product.name = name;
            product.productType = productType;

            WebPage page = browser.NavigateToPage(new Uri(link));
            string pattern = @"<span class=""a-offscreen"">€([^<]+)</span>";
            Regex regex = new Regex(pattern);
            Match match = regex.Match(page.Content);

            if (match.Success)
            {

                string price = match.Groups[1].Value;
                Console.WriteLine($"not parced: {price}");
                price = new string(price.Where(c => char.IsDigit(c) || c == '.' || c == ',').ToArray());
                price = price.Replace(',', '.');
                Console.WriteLine($"#DEBUG: price {price}");
                double priceDouble;

                if (double.TryParse(price, out priceDouble))
                    Console.WriteLine("Parced correctly");
                    else
                    {
                        Console.WriteLine("provided special parsing");
                        for (int i = 0; i < price.Length; i++)
                        {
                            if (price[i] == '.')
                            {
                                price = price.Remove(i, 1);
                                Console.WriteLine($"#DEBUG2: With remove {price}");
                                break;
                            }
                        }
                        
                        priceDouble = Double.Parse(price);
                    Console.WriteLine($"#DEBUG3: double value {priceDouble}");
                    }


                await Console.Out.WriteLineAsync($"Matched value: {priceDouble}");
                product.price = priceDouble;
            }
            else
            {
                await Console.Out.WriteLineAsync("Regex match failed.");
            }

            productList.Add(product);
            //Console.WriteLine("Debug");
        }

        //public async Task AddProduct(string productType, string name, string link)
        //{
        //    Product product = new Product();
        //    product.link = link;
        //    product.name = name;
        //    product.productType = productType;
        //    string response = await client.GetStringAsync(link);
        //    Encoding.GetEncoding("ISO-8859-1");
        //    Console.WriteLine(response);
        //    await Console.Out.WriteLineAsync();
        //    product.price = Double.Parse(System.Text.RegularExpressions.Regex.Match(response,
        //        @"<span class=""a-offscreen"">€([0-9]+\.[0-9]+)</span><span aria-hidden=""true""><span class=""a-price-symbol"">€</span><span class=""a-price-whole"">")
        //        .Groups[1].Value, CultureInfo.InvariantCulture);

        //    productList.Add(product);
        //    //Console.WriteLine("Debug");
        //}

        public double GetTotalSum()
        {
            double sum = 0;

            foreach (Product product in productList)
            {
                sum += product.price;
            }

            return sum;
        }


        public bool CheckFile()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            FileInfo fileInfo = new FileInfo(filePath);
            ExcelPackage excelPackage;

            if (fileInfo.Exists)
                return true;
            else
                return false;
        }

        private void ShowProductInfo(Product product)
        {
            Console.WriteLine($"Product Type: {product.productType}");
            Console.WriteLine($"Product Price: {product.price}");
            Console.WriteLine($"Price Difference: {product.diff}");
        }

        public void ShowProducts()
        {
            double totalSum = 0;
            for (int i = 0; i < productList.Count; i++)
            {
                Console.WriteLine($"[{i+1}]");
                Console.WriteLine("------------");
                ShowProductInfo(productList[i]);
                Console.WriteLine("------------");

                totalSum += productList[i].price;
            }

            Console.WriteLine($"Total Price: {totalSum}");
        }

        public async Task SavePricesToExcel()
        {        
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;        
            FileInfo fileInfo = new FileInfo(filePath);
            ExcelPackage excelPackage;

            if (fileInfo.Exists)
            {
                Console.WriteLine("Already Exist File");
            }
            else
            {
                excelPackage = new ExcelPackage();

                var worksheet = excelPackage.Workbook.Worksheets.Add("Prices");
                worksheet.Cells[1, 1].Value = "Product";
                worksheet.Cells[1, 2].Value = "Price";
                worksheet.Cells[1, 3].Value = "Diff";
                worksheet.Cells[1, 4].Value = "Total Price";
                worksheet.Cells[1, 5].Value = "Date";
                worksheet.Cells[1, 6].Value = "Full Product Name";
                worksheet.Cells[1, 7].Value = "Link";
                worksheet.Cells[1, 1, 1, 7].Style.Font.Bold = true;
                worksheet.Cells[1, 1, 1, 7].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                int row = 2;
                foreach (var product in productList)
                {
                    worksheet.Cells[row, 1].Value = product.productType;
                    worksheet.Cells[row, 2].Value = product.price;
                    if (product.diff > 0)
                    {
                        worksheet.Cells[row, 3].Value = "+" + product.diff;
                        worksheet.Cells[row, 3].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        worksheet.Cells[row, 3].Style.Fill.BackgroundColor.SetColor(Color.Red);
                    }
                    else if (product.diff < 0)
                    {
                        worksheet.Cells[row, 3].Value = product.diff;
                        worksheet.Cells[row, 3].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        worksheet.Cells[row, 3].Style.Fill.BackgroundColor.SetColor(Color.LightGreen);
                    }
                    else
                    {
                        worksheet.Cells[row, 3].Value = product.diff;
                        worksheet.Cells[row, 3].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        worksheet.Cells[row, 3].Style.Fill.BackgroundColor.SetColor(Color.Yellow);
                    }
                    worksheet.Cells[row, 6].Value = product.name;
                    worksheet.Cells[row, 7].Value = product.link;

                    row++;
                }
                worksheet.Cells[2, 4].Value = GetTotalSum();
                worksheet.Cells[2, 5].Value = DateTime.Now.ToString("dd-MM-yyyy h:mm:ss tt");

                // Auto-fit the columns to their content
                worksheet.Cells[1, 1, row - 1, 2].AutoFitColumns();

                // Save the Excel file
                var file = new FileInfo(filePath);
                await excelPackage.SaveAsAsync(file);
            }        
        }

        public void CalculateDiff()
        {
            // Load the Excel file
            DateTime date = DateTime.Now;
            if(GetLatestFilePath() != "null")
            {
                Console.WriteLine(GetLatestFilePath());
                Console.WriteLine("yes");
                var file = new FileInfo(GetLatestFilePath());

                using (var excelPackage = new ExcelPackage(file))
                {
                    ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                    var worksheet = excelPackage.Workbook.Worksheets["Prices"];

                    Console.WriteLine(worksheet.Dimension.Columns);
                    for (int row = 2; row <= worksheet.Dimension.Rows; row++)
                    {
                        Product a = new Product();

                        a.price = productList[row - 2].price;
                        a.link = productList[row - 2].link;
                        a.name = productList[row - 2].name;
                        a.productType = productList[row - 2].productType;

                        a.diff = productList[row - 2].price - Double.Parse(worksheet.Cells[row, 2].Value.ToString());

                        productList[row - 2] = a;
                    }
                }
            }
            
        }
    }

    struct Product
    {
        public string link;
        public string name;
        public string productType;
        public double price;
        public double diff;
    }
    
}
