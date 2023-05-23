using HtmlAgilityPack;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;
using System.Windows.Forms;
using HtmlDocument = HtmlAgilityPack.HtmlDocument;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Excel;
using System.Threading;

namespace oboiParser
{
    public partial class Form1 : Form
    {
        object _lock = new object();
        IWebDriver driver;
        WebDriverWait wait;
        string urlbase = "https://belctanko.ru/";
        public Form1()
        {
            InitializeComponent();
        }

        private async void button1_Click(object sender, EventArgs e)
        {
            driver = new ChromeDriver();
            wait = new WebDriverWait(driver, TimeSpan.FromSeconds(3));
            driver.Url = "https://belctanko.ru/catalog";
            wait.Until(x => x.FindElement(By.XPath("//a[@class='category_link']")));
            var doc = driver.PageSource.CreateDocument();
            var root = doc.DocumentNode;
            var CategoryLink = root.SelectNodes("//a[@class='category_link']");
            List<string> categories = new List<string>();
            foreach (var categoryLink in CategoryLink)
            {
                Debug.WriteLine(categoryLink.InnerText.Trim());
                categories.Add(categoryLink.Attributes["href"].Value);
            }


            foreach (var category in categories)
            {
                string nextpage = urlbase + category;
                List<string> products = new List<string>();

           
            }

        }


        public string NextPage(HtmlNode root)
        {
            string nextpage = string.Empty;
            var PageNodes = root.SelectNodes("//a[@class='page_link']");
            var NextpageNode = PageNodes.AsParallel().FirstOrDefault(x => x.InnerText.Contains("вперёд »"));
            if (NextpageNode != null)
            {
                if (NextpageNode.Attributes["href"].Value.Contains("page-all"))
                    nextpage = string.Empty;

                nextpage = urlbase + NextpageNode.Attributes["href"].Value.TrimStart('/');
            }
            else
            {
                nextpage = string.Empty;
            }
            return nextpage;
        }

        private async void button2_Click(object sender, EventArgs e)
        {
            List<string> products = new List<string>();
            
            using (StreamReader reader = new StreamReader("ProductLink.txt"))
            {
                string line;
                while ((line = await reader.ReadLineAsync()) != null)
                {
                    products.Add(line);
                }
            }
            if(driver ==null)
                driver = new ChromeDriver();
           if(wait ==null)
            {
                wait = new WebDriverWait(driver, TimeSpan.FromSeconds(3));
            }
            int i = 0;
            foreach (var profuctLink in products)
            {
              
                try
                {
                    driver.Url = profuctLink;
                    wait.Until(x => x.FindElement(By.XPath("//div[@class='tab_navigation']")));

                    using (StreamWriter writer = new StreamWriter($"ProductSource\\{i}.html", true))
                    {
                        await writer.WriteLineAsync(driver.PageSource);

                    }
                    using (StreamWriter writer = new StreamWriter($"productDownloads.txt", true))
                    {
                        await writer.WriteLineAsync($"{i}   {driver.Url}    ProductSource\\{i}.html");

                    }
                    i++;
                }
                catch 
                {
                    using (StreamWriter writer = new StreamWriter($"Error_productDownloads.txt", true))
                    {
                        await writer.WriteLineAsync($"{i}   {driver.Url}    ProductSource\\{i}.html");

                    }
                }

            }
        }

        private async void button3_Click(object sender, EventArgs e)
        {
            List<SourceHtmlInfo> list = new List<SourceHtmlInfo>();
            List<SourceHtmlInfo> list2 = new List<SourceHtmlInfo>();

            using (StreamReader reader = new StreamReader("productDownloads.txt"))
            {
                string line;
                
                while ((line = await reader.ReadLineAsync()) != null)
                {
                    string sep = "    ";

                    string[] items = line.Split(sep.ToCharArray());
                    SourceHtmlInfo sourceHtmlInfo = new SourceHtmlInfo()
                    {
                        ID = items[0],
                        Link = items[3],
                        Path = items[7]
                    };
                    list.Add(sourceHtmlInfo);

                }
            }

            var a = list.GroupBy(x => x.Link);
            var c = a.Count();
            foreach (var key in a)
            {
                list2.Add(key.First());
            }
            list = null;
            HtmlDocument doc = new HtmlDocument();
            List<Product> products = new List<Product>();
            int j = 0;
            foreach (SourceHtmlInfo sourceHtmlInfo in list2)
            {
               
                using (StreamReader reader = new StreamReader(sourceHtmlInfo.Path))
                {
                    string text = await reader.ReadToEndAsync();
                    doc.LoadHtml(text);
                    try
                    {
                        
                        var product = doc.ToProduct();
                        product.Link = sourceHtmlInfo.Link;
                        
                        products.Add(product);
     
                    }
                   catch (Exception ex) 
                    {
                        Debug.WriteLine(sourceHtmlInfo.Link);
                        j++;
                       // Debug.Print(j.ToString());
                    }
                   

                }
               
            }
            Debug.Print(j.ToString());
            using (FileStream fs = new FileStream("products.json", FileMode.OpenOrCreate))
            {

                await JsonSerializer.SerializeAsync<List<Product>>(fs, products);
                Console.WriteLine("Data has been saved to file");
            }
        }

        private async  void button4_Click(object sender, EventArgs e)
        {


            /*
            List<Product> products_Table = products.Where(x => x.IsTableCharacteristics).ToList();
            Excel.Application excel_app2 = new Excel.Application();
            excel_app2.Visible = false;
            var workbook2 = excel_app2.Workbooks.Add(1);           
            var sheet2 = (Excel.Worksheet)workbook2.Sheets[1];
           
            sheet2.Cells[1, 1] = "Ссылка";
            sheet2.Cells[1, 2] = "Модель";
            sheet2.Cells[1, 3] = "Цена";
            sheet2.Cells[1, 4] = "Изображение";
            sheet2.Cells[1, 5] = "Назначение";
            sheet2.Cells[1, 6] = "Описание";
            sheet2.Cells[1, 7] = "Комплект поставки";
            sheet2.Cells[1, 8] = "Документация";
            sheet2.Cells[1, 9] = "Характеристики";
            sheet2.Cells[1, 10] = "Каталог";
             var row2 = 2;
            foreach (var product in products_Table)
            {
                sheet2.Cells[row2, 1] = "" + product.Link;
                sheet2.Cells[row2, 2] = "" + product.Name;
                sheet2.Cells[row2, 3] = "" + product.Price;
                sheet2.Cells[row2, 4] = "" + string.Join(";", product.ImageUrls);
                sheet2.Cells[row2, 5] = "" + product.Appointment;
                sheet2.Cells[row2, 6] = "" + product.Peculiarities;
                sheet2.Cells[row2, 7] = "" + product.ContentsOfDelivery;
                sheet2.Cells[row2, 8] = "" + product.Documentation;
                sheet2.Cells[row2, 9] = "" + product.OuterHtmlCharacteristics;
                sheet2.Cells[row2, 10] = "" + product.Path;
                row2++;
            }
            sheet2.Columns.AutoFit();
            sheet2.Rows.RowHeight = 20;
            var filename2 = @"E:\oboiParser\oboiParser\bin\Debug\Result_excel\CharacteristicTable.xlsx";
            workbook2.SaveAs(filename2);
            workbook2.Close(0);
            excel_app2.Quit();
            excel_app2 = null;
            */
            MessageBox.Show("Finish");

        }

        private async void button5_Click(object sender, EventArgs e)
        {
            var products = new List<Product>();



            using (FileStream fs = new FileStream("products.json", FileMode.OpenOrCreate))
            {
                products = await JsonSerializer.DeserializeAsync<List<Product>>(fs);

            }

            List<Product> products_NonTable = products.Where(x => !x.IsTableCharacteristics).ToList();
            var Products_ByGroups = products_NonTable.GroupBy(x => x.Path);
            Excel.Application excel_app = new Excel.Application();
            excel_app.Visible = false; ;
            Debug.Print(Products_ByGroups.Count().ToString());

            foreach (var productPath in Products_ByGroups)
            {
                Dictionary<string, string> Characteristics = new Dictionary<string, string>();
                Debug.Print(productPath.Key);


                var workbook = excel_app.Workbooks.Add(1);
                var worksheet = (Excel.Worksheet)workbook.Sheets[1];
                var sheet = (Excel.Worksheet)workbook.Sheets[1];

                foreach (var product in productPath)
                {
                    foreach (var Characteristic in product.Characteristics)
                    {
                        if (!Characteristics.ContainsKey(Characteristic.Key))
                        {
                            Characteristics.Add(Characteristic.Key, "");
                        }
                    }
                }
                var column = 8;
                int columnCount = 8 + Characteristics.Count();
                string[,] rowData = new string[productPath.Count()+1, 8 + Characteristics.Count()];
                rowData[0, 0] = "Ссылка";
                rowData[0, 1] = "Модель";
                rowData[0, 2] = "Цена";
                rowData[0, 3] = "Изображение";
                rowData[0, 4] = "Назначение";
                rowData[0, 5] = "Описание";
                rowData[0, 6] = "Комплект поставки";
                rowData[0, 7] = "Документация";

                foreach (var Characteristic in Characteristics.Keys)
                {
                    rowData[0, column] = Characteristic;
                    column++;
                }
                
               
               
                var row = 1;
                Debug.Print(productPath.Count().ToString());

                foreach (var product in productPath)
                {
                    
                    rowData[row, 0] = "" + product.Link;
                    rowData[row, 1] = "" + product.Name;
                    rowData[row, 2] = "" + product.Price;
                    rowData[row, 3] = "" + string.Join(";", product.ImageUrls);
                    rowData[row, 4] = "" + product.Appointment;
                    rowData[row, 5] = "" + product.Peculiarities;
                    rowData[row, 6] = "" + product.ContentsOfDelivery;
                    rowData[row, 7] = "" + product.Documentation;
                    column = 8;
                    foreach (var CharacteristicKey in Characteristics.Keys)
                    {
                        if (product.Characteristics.ContainsKey(CharacteristicKey))
                        {
                            rowData[row, column] = product.Characteristics[CharacteristicKey];
                        }
                        else
                        {
                            rowData[row, column] = "";
                        }
                        column++;
                    }

                    row++;
                }
                Range rowRange = sheet.Range[worksheet.Cells[1, 1], worksheet.Cells[productPath.Count() + 1, columnCount]];
                rowRange.Value2 = rowData;
                sheet.Columns.AutoFit();
                sheet.Rows.RowHeight = 20;
                var filename = @"E:\oboiParser\oboiParser\bin\Debug\Result_excel\" + productPath.Key + ".xlsx";
                workbook.SaveAs(filename);
                workbook.Close(0);

            }

            excel_app.Quit();
            excel_app = null;
            MessageBox.Show("Finish");
           
        }
    }
}
