using HtmlAgilityPack;
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
        
        private static List<string> statusDropdown = new List<string> { "Unread", "Read", "Continue" };
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
        private static ICellStyle HeaderStyle(IWorkbook workbook)
        {
            // Create a cell style for the header row
            ICellStyle headerStyle = workbook.CreateCellStyle();
            IFont headerFont = workbook.CreateFont();
            headerFont.Boldweight = (short)FontBoldWeight.Bold;
            headerStyle.SetFont(headerFont);
            headerStyle.Alignment = HorizontalAlignment.Center;
            headerStyle.VerticalAlignment = VerticalAlignment.Center;
            headerStyle.FillForegroundColor = IndexedColors.RoyalBlue.Index;//IndexedColors.Grey25Percent.Index;
            headerStyle.FillPattern = FillPattern.SolidForeground;
            return headerStyle;
        }

        static void CreateCell(IRow row, int column, string value, ICellStyle style)
        {
            ICell cell = row.CreateCell(column);
            cell.SetCellValue(value);
            cell.CellStyle = style;
        }
        public static string WriteExcel(List<Book> data)
        {
            IWorkbook workbook = new XSSFWorkbook();
            var firstSheet = "Details";
            ISheet sheet = workbook.CreateSheet(firstSheet);

            #region Add header row
            ICellStyle headerStyle = HeaderStyle(workbook);
            IRow headerRow = sheet.CreateRow(0);
            CreateCell(headerRow, 0, "Image", headerStyle);
            CreateCell(headerRow, 1, "Title", headerStyle);
            CreateCell(headerRow, 2, "Summary", headerStyle);
            CreateCell(headerRow, 3, "Author", headerStyle);
            CreateCell(headerRow, 4, "Author Description", headerStyle);
            CreateCell(headerRow, 5, "Translator", headerStyle);
            CreateCell(headerRow, 6, "Price", headerStyle);
            CreateCell(headerRow, 7, "Sell Price", headerStyle);
            CreateCell(headerRow, 8, "Discount", headerStyle);
            CreateCell(headerRow, 9, "Discount Rate", headerStyle);
            CreateCell(headerRow, 10, "Number of Pages", headerStyle);
            CreateCell(headerRow, 11, "Publisher", headerStyle);
            CreateCell(headerRow, 12, "Category", headerStyle);
            CreateCell(headerRow, 13, "ISBN", headerStyle);
            CreateCell(headerRow, 14, "Edition", headerStyle);
            CreateCell(headerRow, 15, "Country", headerStyle);
            CreateCell(headerRow, 16, "Language", headerStyle);
            CreateCell(headerRow, 17, "Category Link", headerStyle);
            CreateCell(headerRow, 18, "Status", headerStyle);
            CreateCell(headerRow, 19, "Rating", headerStyle);
            CreateCell(headerRow, 20, "Tag", headerStyle);
            CreateCell(headerRow, 21, "Link", headerStyle);
            #endregion
            sheet.SetAutoFilter(new CellRangeAddress(0, 0, 0, headerRow.LastCellNum - 1));
            sheet.SetColumnWidth(1, 7000);
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
                dataRow.CreateCell(18).SetCellValue(statusDropdown[0]);
                dataRow.CreateCell(19).SetCellValue(3);
                dataRow.CreateCell(20).SetCellValue(data[i].Tag);
                dataRow.CreateCell(21).SetCellValue(data[i].Link);
                // Add other data cells as needed
            }

            //Style
            ICellStyle cellStyle = workbook.CreateCellStyle();
            cellStyle.Alignment = HorizontalAlignment.Center; // Adjust alignment as needed

            #region Total count at the end
            IRow footerRow = sheet.CreateRow(data.Count + 1);
            var cell1 = footerRow.CreateCell(0);
            cell1.SetCellValue("Total");
            cell1.CellStyle = cellStyle;
            CellRangeAddress mergedRegion = new CellRangeAddress(data.Count + 1, data.Count + 1, 0, 5);
            sheet.AddMergedRegion(mergedRegion);


            footerRow.CreateCell(6).SetCellFormula($"SUM(G2:G{data.Count + 1})");
            footerRow.CreateCell(7).SetCellFormula($"SUM(H2:H{data.Count + 1})");
            footerRow.CreateCell(8).SetCellFormula($"SUM(I2:I{data.Count + 1})");
            footerRow.CreateCell(9).SetCellFormula($"SUM(J2:J{data.Count + 1})");
            footerRow.CreateCell(10).SetCellFormula($"SUM(K2:K{data.Count + 1})");
            #endregion

            #region status dropdown
            // Apply data validation to the cell (A2 in this case)
            CellRangeAddressList statusRang = new CellRangeAddressList(1, 1000, 18, 18);
            XSSFDataValidationHelper dvHelper = new XSSFDataValidationHelper((XSSFSheet)sheet);
            XSSFDataValidationConstraint dvConstraint = (XSSFDataValidationConstraint)dvHelper.CreateExplicitListConstraint(statusDropdown.ToArray());
            XSSFDataValidation validation = (XSSFDataValidation)dvHelper.CreateValidation(dvConstraint, statusRang);

            // Set other data validation options if needed
            validation.ShowErrorBox = true;
            validation.CreateErrorBox("Invalid Value", "Please select a value from the dropdown list.");
            sheet.AddValidationData(validation);
            #endregion

            #region add new sheet
            var summarySheet = "Summary";
            ISheet summary = workbook.CreateSheet(summarySheet);
            IRow status1 = summary.CreateRow(5);
            status1.CreateCell(3).SetCellValue(@$"{statusDropdown[0]}");
            status1.CreateCell(4).SetCellFormula($"CountIf({firstSheet}!S1:{firstSheet}!S{data.Count + 1},\"{statusDropdown[0]}\")");

            IRow status2 = summary.CreateRow(6);
            status2.CreateCell(3).SetCellValue(@$"{statusDropdown[1]}");
            status2.CreateCell(4).SetCellFormula($"CountIf({firstSheet}!S1:{firstSheet}!S{data.Count + 1},\"{statusDropdown[1]}\")");

            IRow status3 = summary.CreateRow(7);
            status3.CreateCell(3).SetCellValue(@$"{statusDropdown[2]}");
            status3.CreateCell(4).SetCellFormula($"CountIf({firstSheet}!S1:{firstSheet}!S{data.Count + 1},\"{statusDropdown[2]}\")");
            #endregion

            XSSFFormulaEvaluator.EvaluateAllFormulaCells(workbook);
            // Save the new workbook to a file
            string filePath = @$"bookmark.xlsx";
            using (var fileStream = System.IO.File.OpenWrite(filePath))
            {
                workbook.Write(fileStream);
            }
            return filePath;
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
