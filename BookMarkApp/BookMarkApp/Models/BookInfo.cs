using HtmlAgilityPack;
using Microsoft.AspNetCore.Mvc;
using Microsoft.VisualBasic.FileIO;
using NPOI.HSSF.UserModel;
using NPOI.OpenXmlFormats.Dml.Diagram;
using NPOI.SS.Formula.Functions;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using Org.BouncyCastle.Utilities.Encoders;
using SixLabors.Fonts.Tables.AdvancedTypographic;
using System.Diagnostics.Metrics;
using System.Dynamic;
using System.Formats.Asn1;
using System.Globalization;
using System.Security.Policy;
using System.Text;
using static System.Reflection.Metadata.BlobBuilder;

namespace BookMarkApp.Models
{
    internal static class BookInfo
    {
        public static async Task<Book> GetBookDetails(string url)
        {
            var rokomari = url.Contains("rokomari");
            if (!Uri.IsWellFormedUriString(url, UriKind.Absolute) || !rokomari)
            {
                return new Book() { Title = "Invalid", Link = url };
            }
            var Webget = new HtmlWeb();
            var doc = await Webget.LoadFromWebAsync(url);
            var book = new Book() { Link = url };
            var cols = doc.DocumentNode.SelectNodes("//table[@class='table table-bordered']//tr//td");
            for (int i = 0; cols != null && i < cols.Count; i = i + 2)
            {
                string name = string.Join("", cols[i].InnerText.Trim().Split(default(string[]), StringSplitOptions.RemoveEmptyEntries));
                string value = cols[i + 1].InnerText;
                book.SetValue(name, value);
            }
            book.Price = doc.DocumentNode.SelectNodes("//div[@class='details-book-info__content-book-price']//strike[@class='original-price']")?[0].InnerText.Trim();

            book.SellPrice = doc.DocumentNode.SelectNodes("//div[@class='details-book-info__content-book-price']//span[@class='sell-price']")?[0].InnerText.Trim();

            book.Discount = doc.DocumentNode.SelectNodes("//div[@class='details-book-info__content-book-price']//span[@class='js--save-message']")?[0].InnerText.Trim();

            var category = doc.DocumentNode.SelectNodes("//div[@class='details-book-info__content-category d-flex align-items-center']//a");
            book.Category = category?[0].InnerText.Trim();
            book.CategoryLink = string.Concat(url.Split("com")?[0], "com", category?[0].Attributes["href"].Value);

            book.BookImg = doc.DocumentNode.SelectNodes("//div//div//div//div//div[@class='image-container']//img")?[0].Attributes["src"].Value;

            book.Summary = doc.DocumentNode.SelectNodes("//div[@id='js--summary-description']")?[0].InnerText;

            //var authorName = doc.DocumentNode.SelectNodes("//div[@class='col author_des']//p")?[0].InnerText;

            book.AuthorDescription = doc.DocumentNode.SelectNodes("//div[@class='col author_des']//div")?[0].InnerText;

            return book;
        }

        private static string GetCellValue(ICell cell)
        {
            if (cell == null)
            {
                return null;
            }

            switch (cell.CellType)
            {
                case CellType.String:
                    return cell.StringCellValue;

                case CellType.Numeric:
                    return cell.NumericCellValue.ToString();

                // Handle other cell types as needed

                default:
                    return null;
            }
        }
        public static List<Book> ReadExcel(IFormFile excelFile)
        {
            using (var stream = excelFile.OpenReadStream())
            {
                IWorkbook workbook = new XSSFWorkbook(stream);
                ISheet sheet = workbook.GetSheetAt(0);

                List<Book> books = new List<Book>();

                for (int i = 0; i <= sheet.LastRowNum; i++)
                {
                    var row = sheet.GetRow(i);

                    if (row != null)
                    {
                        var rowData = new Book
                        {
                            BookImg = GetCellValue(row.GetCell(0)),
                            Title = GetCellValue(row.GetCell(0)),
                            Summary = GetCellValue(row.GetCell(0)),
                            Author = GetCellValue(row.GetCell(0)),
                            AuthorDescription = GetCellValue(row.GetCell(0)),
                            Translator = GetCellValue(row.GetCell(0)),
                            Price = GetCellValue(row.GetCell(0)),
                            SellPrice = GetCellValue(row.GetCell(0)),
                            Discount = GetCellValue(row.GetCell(0)),
                            NumberofPages = GetCellValue(row.GetCell(0)),
                            Publisher = GetCellValue(row.GetCell(0)),
                            Category = GetCellValue(row.GetCell(0)),
                            ISBN = GetCellValue(row.GetCell(0)),
                            Edition = GetCellValue(row.GetCell(0)),
                            Country = GetCellValue(row.GetCell(0)),
                            Language = GetCellValue(row.GetCell(0)),
                            CategoryLink = GetCellValue(row.GetCell(0)),
                            Tag = GetCellValue(row.GetCell(0)),
                            Link = GetCellValue(row.GetCell(0)),
                        };

                        books.Add(rowData);
                    }
                }
                return books;
            }
        }
        public static List<Book> ReadCsv(IFormFile csvFile)
        {
            List<Book> data = new List<Book>();

            using (var reader = new System.IO.StreamReader(csvFile.OpenReadStream()))
            {
                using (TextFieldParser parser = new TextFieldParser(reader))
                {
                    parser.TextFieldType = FieldType.Delimited;
                    parser.SetDelimiters(",");

                    while (!parser.EndOfData)
                    {
                        string[] fields = parser.ReadFields();
                        if (fields != null)
                        {
                            var rowData = new Book
                            {
                                BookImg = fields.Length > 0 ? fields[0] : null,
                                Title = fields.Length > 0 ? fields[0] : null,
                                Summary = fields.Length > 0 ? fields[0] : null,
                                Author = fields.Length > 0 ? fields[0] : null,
                                AuthorDescription = fields.Length > 0 ? fields[0] : null,
                                Translator = fields.Length > 0 ? fields[0] : null,
                                Price = fields.Length > 0 ? fields[0] : null,
                                SellPrice = fields.Length > 0 ? fields[0] : null,
                                Discount = fields.Length > 0 ? fields[0] : null,
                                NumberofPages = fields.Length > 0 ? fields[0] : null,
                                Publisher = fields.Length > 0 ? fields[0] : null,
                                Category = fields.Length > 0 ? fields[0] : null,
                                ISBN = fields.Length > 0 ? fields[0] : null,
                                Edition = fields.Length > 0 ? fields[0] : null,
                                Country = fields.Length > 0 ? fields[0] : null,
                                Language = fields.Length > 0 ? fields[0] : null,
                                CategoryLink = fields.Length > 0 ? fields[0] : null,
                                Tag = fields.Length > 0 ? fields[0] : null,
                                Link = fields.Length > 0 ? fields[0] : null,
                            };

                            data.Add(rowData);
                        }
                    }
                }
            }

            return data;
        }
        public static string WriteExcel(List<Book> data)
        {
            IWorkbook workbook = new XSSFWorkbook();
            ISheet sheet = workbook.CreateSheet("Sheet1");

            // Add header row
            IRow headerRow = sheet.CreateRow(0);
            headerRow.CreateCell(0).SetCellValue("BookImg");
            headerRow.CreateCell(1).SetCellValue("Title");
            headerRow.CreateCell(2).SetCellValue("Summary");
            headerRow.CreateCell(3).SetCellValue("Author");
            headerRow.CreateCell(4).SetCellValue("AuthorDescription");
            headerRow.CreateCell(5).SetCellValue("Translator");
            headerRow.CreateCell(6).SetCellValue("Price");
            headerRow.CreateCell(7).SetCellValue("SellPrice");
            headerRow.CreateCell(8).SetCellValue("Discount");
            headerRow.CreateCell(9).SetCellValue("NumberofPages");
            headerRow.CreateCell(10).SetCellValue("Publisher");
            headerRow.CreateCell(11).SetCellValue("Category");
            headerRow.CreateCell(12).SetCellValue("ISBN");
            headerRow.CreateCell(13).SetCellValue("Edition");
            headerRow.CreateCell(14).SetCellValue("Country");
            headerRow.CreateCell(15).SetCellValue("Language");
            headerRow.CreateCell(16).SetCellValue("CategoryLink");
            headerRow.CreateCell(17).SetCellValue("Tag");
            headerRow.CreateCell(18).SetCellValue("Link");
            // Add other header cells as needed

            // Add data rows
            for (int i = 0; i < data.Count; i++)
            {
                IRow dataRow = sheet.CreateRow(i + 1);
                headerRow.CreateCell(0).SetCellValue(data[i].BookImg);
                headerRow.CreateCell(1).SetCellValue(data[i].Title);
                headerRow.CreateCell(2).SetCellValue(data[i].Summary);
                headerRow.CreateCell(3).SetCellValue(data[i].Author);
                headerRow.CreateCell(4).SetCellValue(data[i].AuthorDescription);
                headerRow.CreateCell(5).SetCellValue(data[i].Translator);
                headerRow.CreateCell(6).SetCellValue(data[i].Price);
                headerRow.CreateCell(7).SetCellValue(data[i].SellPrice);
                headerRow.CreateCell(8).SetCellValue(data[i].Discount);
                headerRow.CreateCell(9).SetCellValue(data[i].NumberofPages);
                headerRow.CreateCell(10).SetCellValue(data[i].Publisher);
                headerRow.CreateCell(11).SetCellValue(data[i].Category);
                headerRow.CreateCell(12).SetCellValue(data[i].ISBN);
                headerRow.CreateCell(13).SetCellValue(data[i].Edition);
                headerRow.CreateCell(14).SetCellValue(data[i].Country);
                headerRow.CreateCell(15).SetCellValue(data[i].Language);
                headerRow.CreateCell(16).SetCellValue(data[i].CategoryLink);
                headerRow.CreateCell(17).SetCellValue(data[i].Tag);
                headerRow.CreateCell(18).SetCellValue(data[i].Link);
                // Add other data cells as needed
            }

            // Save the new workbook to a file
            string filePath = @$"bookmark_{DateTime.Now.ToShortDateString()}.xlsx";
            using (var fileStream = System.IO.File.OpenWrite(filePath))
            {
                workbook.Write(fileStream);
            }
            return filePath;
        }
        public static byte[] MakeExcel(List<Book> books) //CSV
        {
            MemoryStream ms = new MemoryStream();
            using (StreamWriter sw = new StreamWriter(ms, Encoding.UTF8))
            {
                var stringCSV = ConvertToCsv(books);
                sw.WriteLine(stringCSV);
                sw.Flush();
            }
            byte[] bytes = ms.ToArray();
            ms.Close();
            return bytes;
        }
        static string ConvertToCsv<T>(IEnumerable<T> data)
        {
            StringWriter csvString = new StringWriter();

            // Write the header
            csvString.WriteLine(string.Join(",", typeof(T).GetProperties().Select(prop => QuoteIfNeeded(prop.Name))));

            // Write the data
            foreach (var item in data)
            {
                var fields = typeof(T).GetProperties()
                    .Select(prop => QuoteIfNeeded(prop.GetValue(item)?.ToString() ?? ""));
                csvString.WriteLine(string.Join(",", fields));
            }

            return csvString.ToString();
        }

        static string QuoteIfNeeded(string value)
        {
            if (value.Contains(',') || value.Contains('"'))
            {
                // Escape double quotes by doubling them and enclose the value in double quotes
                return $"\"{value.Replace("\"", "\"\"")}\"";
            }
            else
            {
                return value;
            }
        }
        public static void SetValue<T>(this T sender, string propertyName, object value)
        {
            var propertyInfo = sender.GetType().GetProperty(propertyName);

            if (propertyInfo is null) return;

            var type = Nullable.GetUnderlyingType(propertyInfo.PropertyType) ?? propertyInfo.PropertyType;

            if (propertyInfo.PropertyType.IsEnum)
            {
                propertyInfo.SetValue(sender, Enum.Parse(propertyInfo.PropertyType, value.ToString()!));
            }
            else
            {
                var safeValue = (value == null) ? null : Convert.ChangeType(value, type);
                propertyInfo.SetValue(sender, safeValue, null);
            }
        }

    }
    public class Book
    {
        public string BookImg { get; set; }
        public string Title { get; set; }
        public string Summary { get; set; }
        public string Author { get; set; }
        public string AuthorDescription { get; set; }
        public string Translator { get; set; }
        public string Price { get; set; }
        public string SellPrice { get; set; }
        public string Discount { get; set; }
        public string NumberofPages { get; set; }
        public string Publisher { get; set; }
        public string Category { get; set; }
        public string ISBN { get; set; }
        public string Edition { get; set; }
        public string Country { get; set; }
        public string Language { get; set; }
        public string CategoryLink { get; set; }
        public string Tag { get; set; }
        public string Link { get; set; }
    }
}
