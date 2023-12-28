using HtmlAgilityPack;
using System.Dynamic;
using System.Formats.Asn1;
using System.Globalization;
using System.Text;
using static System.Reflection.Metadata.BlobBuilder;

namespace BookMarkApp.Models
{
    internal static class BookInfo
    {
        public static async Task<Book> GetBookDetails(string url)
        {
            try
            {

            var rokomari = url.Contains("rokomari");
            if (!Uri.IsWellFormedUriString(url, UriKind.Absolute) || !rokomari)
            {
                return new Book() { Title = "Invalid", Link = url};
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
            book.CategoryLink = string.Concat(url.Split("com")?[0], "com",category?[0].Attributes["href"].Value);

            book.BookImg = doc.DocumentNode.SelectNodes("//div//div//div//div//div[@class='image-container']//img")?[0].Attributes["src"].Value;

            book.Summary = doc.DocumentNode.SelectNodes("//div[@id='js--summary-description']")?[0].InnerText;

            //var authorName = doc.DocumentNode.SelectNodes("//div[@class='col author_des']//p")?[0].InnerText;

            book.AuthorDescription = doc.DocumentNode.SelectNodes("//div[@class='col author_des']//div")?[0].InnerText;

            return book;

            }
            catch (Exception ex)
            {

                throw;
            }
        }

        public static byte[] makeExcel(List<Book> books)
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
                var values = typeof(T).GetProperties()
                    .Select(prop => QuoteIfNeeded(prop.GetValue(item)?.ToString() ?? ""));
                csvString.WriteLine(string.Join(",", values));
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
