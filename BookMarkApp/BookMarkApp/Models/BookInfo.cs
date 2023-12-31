using HtmlAgilityPack;
using Microsoft.VisualBasic.FileIO;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using System.Data;
using System.Text;
using System.Text.RegularExpressions;
using static System.Reflection.Metadata.BlobBuilder;

namespace BookMarkApp.Models
{
    internal static class BookInfo
    {
        public static async Task<Book> GetBookDetails(string url)
        {
            var rokomari = url.Contains("rokomari");
            if (!Uri.IsWellFormedUriString(url, UriKind.Absolute))
            {
                return null;
            }
            else if (!rokomari)
            {
                return new Book() { Link = url };
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
            var price = doc.DocumentNode.SelectNodes("//div[@class='details-book-info__content-book-price']//strike[@class='original-price']")?[0].InnerText.Trim();
            book.Price = String.IsNullOrEmpty(price) ? 0 : Convert.ToDecimal(getPrice(price));

            var sellPrice = doc.DocumentNode.SelectNodes("//div[@class='details-book-info__content-book-price']//span[@class='sell-price']")?[0].InnerText.Trim();
            book.SellPrice = string.IsNullOrEmpty(sellPrice) ? 0 : Convert.ToDecimal(getPrice(sellPrice));

            var discount = getDiscount(doc.DocumentNode.SelectNodes("//div[@class='details-book-info__content-book-price']//span[@class='js--save-message']")?[0].InnerText.Trim());
            book.Discount = Convert.ToDecimal(discount.Key);
            book.DiscountRate = Convert.ToDecimal(discount.Value);

            var category = doc.DocumentNode.SelectNodes("//div[@class='details-book-info__content-category d-flex align-items-center']//a");
            book.Category = category?[0].InnerText.Trim();
            book.CategoryLink = string.Concat(url.Split("com")?[0], "com", category?[0].Attributes["href"].Value);

            book.BookImg = doc.DocumentNode.SelectNodes("//div//div//div//div//div[@class='image-container']//img")?[0].Attributes["src"].Value;

            book.Summary = doc.DocumentNode.SelectNodes("//div[@id='js--summary-description']")?[0].InnerText;

            //var authorName = doc.DocumentNode.SelectNodes("//div[@class='col author_des']//p")?[0].InnerText;

            book.AuthorDescription = doc.DocumentNode.SelectNodes("//div[@class='col author_des']//div")?[0].InnerText;

            return book;
        }
        private static string getPrice(string price)
        {
            try
            {
                return price.Trim().Split(" ")[1];
            }
            catch (Exception ex)
            {
                return price;
            }
        }
        private static KeyValuePair<decimal, decimal> getDiscount(string discount)
        {
            try
            {
                if (string.IsNullOrEmpty(discount))
                {
                    return new KeyValuePair<decimal, decimal>(0, 0);
                }
                var d = discount.Trim().Split("TK. ")[1].Split(" ");
                return new KeyValuePair<decimal, decimal>(Convert.ToDecimal(d[0]), Convert.ToDecimal(Regex.Match(d[1], @"\d+").Value));
            }
            catch (Exception ex)
            {
                return new KeyValuePair<decimal, decimal>(0, 0);
            }
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
                            Title = GetCellValue(row.GetCell(1)),
                            Summary = GetCellValue(row.GetCell(2)),
                            Author = GetCellValue(row.GetCell(3)),
                            AuthorDescription = GetCellValue(row.GetCell(4)),
                            Translator = GetCellValue(row.GetCell(5)),
                            Price = Convert.ToDecimal(GetCellValue(row.GetCell(6))),
                            SellPrice = Convert.ToDecimal(GetCellValue(row.GetCell(7))),
                            Discount = Convert.ToDecimal(GetCellValue(row.GetCell(8))),
                            DiscountRate = Convert.ToDecimal(GetCellValue(row.GetCell(9))),
                            NumberofPages = Convert.ToInt32(GetCellValue(row.GetCell(10))),
                            Publisher = GetCellValue(row.GetCell(11)),
                            Category = GetCellValue(row.GetCell(12)),
                            ISBN = GetCellValue(row.GetCell(13)),
                            Edition = GetCellValue(row.GetCell(14)),
                            Country = GetCellValue(row.GetCell(15)),
                            Language = GetCellValue(row.GetCell(16)),
                            CategoryLink = GetCellValue(row.GetCell(17)),
                            Status = GetCellValue(row.GetCell(16)),
                            Rating = GetCellValue(row.GetCell(17)),
                            Tag = GetCellValue(row.GetCell(20)),
                            Link = GetCellValue(row.GetCell(21)),
                        };

                        books.Add(rowData);
                    }
                }
                return books;
            }
        }
        //public static List<Book> ReadCsv(IFormFile csvFile)
        //{
        //    try
        //    {

        //        List<Book> data = new List<Book>();

        //        using (var reader = new System.IO.StreamReader(csvFile.OpenReadStream()))
        //        {
        //            using (TextFieldParser parser = new TextFieldParser(reader))
        //            {
        //                parser.TextFieldType = FieldType.Delimited;
        //                parser.SetDelimiters(",");

        //                while (!parser.EndOfData)
        //                {
        //                    string[] fields = parser.ReadFields();
        //                    if (fields != null)
        //                    {
        //                        var rowData = new Book
        //                        {
        //                            BookImg = fields.Length > 0 ? fields[0] : null,
        //                            Title = fields.Length > 0 ? fields[0] : null,
        //                            Summary = fields.Length > 0 ? fields[0] : null,
        //                            Author = fields.Length > 0 ? fields[0] : null,
        //                            AuthorDescription = fields.Length > 0 ? fields[0] : null,
        //                            Translator = fields.Length > 0 ? fields[0] : null,
        //                            Price = fields.Length > 0 ? (double) (fields[0] ?? 0) : null,
        //                            SellPrice = fields.Length > 0 ? fields[0] : null,
        //                            Discount = fields.Length > 0 ? fields[0] : null,
        //                            NumberofPages = fields.Length > 0 ? fields[0] : null,
        //                            Publisher = fields.Length > 0 ? fields[0] : null,
        //                            Category = fields.Length > 0 ? fields[0] : null,
        //                            ISBN = fields.Length > 0 ? fields[0] : null,
        //                            Edition = fields.Length > 0 ? fields[0] : null,
        //                            Country = fields.Length > 0 ? fields[0] : null,
        //                            Language = fields.Length > 0 ? fields[0] : null,
        //                            CategoryLink = fields.Length > 0 ? fields[0] : null,
        //                            Tag = fields.Length > 0 ? fields[0] : null,
        //                            Link = fields.Length > 0 ? fields[0] : null,
        //                        };

        //                        data.Add(rowData);
        //                    }
        //                }
        //            }
        //        }

        //        return data;

        //    }
        //    catch (Exception ex)
        //    {

        //        return null;
        //    }
        //}
        public static string WriteExcel(List<Book> data)
        {
            IWorkbook workbook = new XSSFWorkbook();
            ISheet sheet = workbook.CreateSheet("Details");

            // Add header row
            IRow headerRow = sheet.CreateRow(0);
            headerRow.CreateCell(0).SetCellValue("Image");
            headerRow.CreateCell(1).SetCellValue("Title");
            headerRow.CreateCell(2).SetCellValue("Summary");
            headerRow.CreateCell(3).SetCellValue("Author");
            headerRow.CreateCell(4).SetCellValue("Author Description");
            headerRow.CreateCell(5).SetCellValue("Translator");
            headerRow.CreateCell(6).SetCellValue("Price");
            headerRow.CreateCell(7).SetCellValue("Sell Price");
            headerRow.CreateCell(8).SetCellValue("Discount");
            headerRow.CreateCell(9).SetCellValue("Discount Rate");
            headerRow.CreateCell(10).SetCellValue("Number of Pages");
            headerRow.CreateCell(11).SetCellValue("Publisher");
            headerRow.CreateCell(12).SetCellValue("Category");
            headerRow.CreateCell(13).SetCellValue("ISBN");
            headerRow.CreateCell(14).SetCellValue("Edition");
            headerRow.CreateCell(15).SetCellValue("Country");
            headerRow.CreateCell(16).SetCellValue("Language");
            headerRow.CreateCell(17).SetCellValue("Category Link");
            headerRow.CreateCell(18).SetCellValue("Status");
            headerRow.CreateCell(19).SetCellValue("Rating");
            headerRow.CreateCell(20).SetCellValue("Tag");
            headerRow.CreateCell(21).SetCellValue("Link");

            // Add other header cells as needed
            List<string> statusDropdown = new List<string> { "--X--", "Unread", "Read", "Continue" };
            // Add data rows
            for (int i = 0; i < data.Count; i++)
            {
                IRow dataRow = sheet.CreateRow(i + 1);
                dataRow.CreateCell(0).SetCellValue(data[i].BookImg);
                dataRow.CreateCell(1).SetCellValue(data[i].Title);
                dataRow.CreateCell(2).SetCellValue(data[i].Summary);
                dataRow.CreateCell(3).SetCellValue(data[i].Author);
                dataRow.CreateCell(4).SetCellValue(data[i].AuthorDescription);
                dataRow.CreateCell(5).SetCellValue(data[i].Translator);
                dataRow.CreateCell(6).SetCellValue((double)data[i].Price);
                dataRow.CreateCell(7).SetCellValue((double)data[i].SellPrice);
                dataRow.CreateCell(8).SetCellValue((double)data[i].Discount);
                dataRow.CreateCell(9).SetCellValue((double)data[i].DiscountRate);
                dataRow.CreateCell(10).SetCellValue(data[i].NumberofPages);
                dataRow.CreateCell(11).SetCellValue(data[i].Publisher);
                dataRow.CreateCell(12).SetCellValue(data[i].Category);
                dataRow.CreateCell(13).SetCellValue(data[i].ISBN);
                dataRow.CreateCell(14).SetCellValue(data[i].Edition);
                dataRow.CreateCell(15).SetCellValue(data[i].Country);
                dataRow.CreateCell(16).SetCellValue(data[i].Language);
                dataRow.CreateCell(17).SetCellValue(data[i].CategoryLink);
                dataRow.CreateCell(18).SetCellValue(statusDropdown[1]);
                dataRow.CreateCell(19).SetCellValue(3);
                dataRow.CreateCell(20).SetCellValue(data[i].Tag);
                dataRow.CreateCell(21).SetCellValue(data[i].Link);
                // Add other data cells as needed
            }

            //Style
            ICellStyle cellStyle = workbook.CreateCellStyle();
            cellStyle.Alignment = HorizontalAlignment.Center; // Adjust alignment as needed

            //Total 
            IRow footerRow = sheet.CreateRow(data.Count + 1);
            footerRow.CreateCell(6).SetCellFormula($"SUM(G2:G{data.Count + 1})");
            footerRow.CreateCell(7).SetCellFormula($"SUM(H2:H{data.Count + 1})");
            footerRow.CreateCell(8).SetCellFormula($"SUM(I2:I{data.Count + 1})");
            footerRow.CreateCell(9).SetCellFormula($"SUM(J2:J{data.Count + 1})");
            var cell1 = footerRow.CreateCell(0);
            cell1.SetCellValue("Total");
            cell1.CellStyle = cellStyle;
            CellRangeAddress mergedRegion = new CellRangeAddress(data.Count + 1, data.Count + 1, 0, 5);

            // Apply data validation to the cell (A2 in this case)
            CellRangeAddressList statusRang = new CellRangeAddressList(1, 1000, 18, 18);
            XSSFDataValidationHelper dvHelper = new XSSFDataValidationHelper((XSSFSheet)sheet);
            XSSFDataValidationConstraint dvConstraint = (XSSFDataValidationConstraint)dvHelper.CreateExplicitListConstraint(statusDropdown.ToArray());
            XSSFDataValidation validation = (XSSFDataValidation)dvHelper.CreateValidation(dvConstraint, statusRang);

            // Set other data validation options if needed
            validation.ShowErrorBox = true;
            validation.CreateErrorBox("Invalid Value", "Please select a value from the dropdown list.");

            sheet.AddValidationData(validation);


            ISheet summary = workbook.CreateSheet("Summary");

            ////Total 
            //IRow summaryRow = summary.CreateRow(0);
            //summaryRow.CreateCell(4).SetCellFormula($"CountIf(Sheet1!K1:Sheet1!K{data.Count + 1},\"unread\")");
            ////footerRow.CreateCell(7).SetCellFormula($"SUM(H2:H{data.Count + 1})");
            ////footerRow.CreateCell(8).SetCellFormula($"SUM(I2:I{data.Count + 1})");
            ////footerRow.CreateCell(9).SetCellFormula($"SUM(J2:J{data.Count + 1})");
            //var summaryCell = summaryRow.CreateCell(3);
            //summaryCell.SetCellValue("Unread");
            //summaryCell.CellStyle = cellStyle;
            //// Create a cell style and set the text alignment


            // Apply the cell style to the cell
            sheet.AddMergedRegion(mergedRegion);
            XSSFFormulaEvaluator.EvaluateAllFormulaCells(workbook);


            // Save the new workbook to a file
            string filePath = @$"bookmark_{DateTime.Now.ToString("yyyy_MM_dd_HHmmss")}.xlsx";
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
        public decimal Price { get; set; }
        public decimal SellPrice { get; set; }
        public decimal Discount { get; set; }
        public decimal DiscountRate { get; set; }
        public int NumberofPages { get; set; }
        public string Publisher { get; set; }
        public string Category { get; set; }
        public string ISBN { get; set; }
        public string Edition { get; set; }
        public string Country { get; set; }
        public string Language { get; set; }
        public string CategoryLink { get; set; }
        public string Status { get; set; }
        public string Rating { get; set; }
        public string Tag { get; set; }
        public string Link { get; set; }
    }
}
