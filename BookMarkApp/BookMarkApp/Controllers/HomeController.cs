using BookMarkApp.Models;
using HtmlAgilityPack;
using Microsoft.AspNetCore.Mvc;
using System.Diagnostics;
using System.Net;
using System.Text;
using System.Xml.Linq;

namespace BookMarkApp.Controllers
{
    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> _logger;

        public HomeController(ILogger<HomeController> logger)
        {
            _logger = logger;
        }

        public async Task<IActionResult> Index()
        {
            var url = "https://www.rokomari.com/book/15991/bangali-musulmaner-mon";
            //using (var client = new HttpClient())
            //{
            //    // We'll use the GetAsync method to send  
            //    // a GET request to the specified URL 
            //    var response = await client.GetAsync(url);

            //    // If the response is successful, we'll 
            //    // interpret the response as XML 
            //    if (response.IsSuccessStatusCode)
            //    {
            //        var xml = await response.Content.ReadAsStringAsync();

            //        // We can then use the LINQ to XML API to query the XML 
            //        var doc = XDocument.Parse(xml);

            //        // Let's query the XML to get all of the <title> elements 
            //        var titles = from el in doc.Descendants("title")
            //                     select el.Value;

            //        // And finally, we'll print out the titles 
            //        foreach (var title in titles)
            //        {
            //            Console.WriteLine(title);
            //        }
            //    }
            //}
           // string url = "http://website.com";
            var Webget = new HtmlWeb();
            var doc = Webget.Load(url);
            //foreach (HtmlNode node in doc.DocumentNode.SelectNodes("//table"))
            //{
            //    var j = node;
            //   // names.Add(node.ChildNodes[0].InnerHtml);
            //}
            //foreach (HtmlNode node in doc.DocumentNode.SelectNodes("//li[@class='tel']//a"))
            //{
            //    var node1 = node;
            //   // phones.Add(node.ChildNodes[0].InnerHtml);
            //}
            var cols = doc.DocumentNode.SelectNodes("//table[@class='table table-bordered']//tr//td");
            for (int i = 0; i < cols.Count; i = i + 2)
            {
                string name = cols[i].InnerText.Trim();
                string va = cols[i + 1].InnerText;
            }
            var price = doc.DocumentNode.SelectNodes("//div[@class='details-book-info__content-book-price']//strike[@class='original-price']"); //price
            for (int i = 0; i < price.Count; i = i + 2)
            {
                string name2 = price[i].InnerText.Trim();
                //string va = price[i + 1].InnerText;
            }
            var sellPrice = doc.DocumentNode.SelectNodes("//div[@class='details-book-info__content-book-price']//span[@class='sell-price']"); //price
            string se = sellPrice[0].InnerText.Trim();

            var discount = doc.DocumentNode.SelectNodes("//div[@class='details-book-info__content-book-price']//span[@class='js--save-message']"); //price
            string d = discount[0].InnerText.Trim();

            var category = doc.DocumentNode.SelectNodes("//div[@class='details-book-info__content-category d-flex align-items-center']//a"); //price
            string c = category[0].InnerText.Trim();
            string href = category[0].Attributes["href"].Value;

            var bookImg = doc.DocumentNode.SelectNodes("//div//div//div//div//div[@class='image-container']//img"); //price
            var img = bookImg[0].Attributes["src"].Value;

            var bookSummary = doc.DocumentNode.SelectNodes("//div[@id='js--summary-description']"); //price
            var summary = bookSummary[0].InnerText;

            var authorName = doc.DocumentNode.SelectNodes("//div[@class='col author_des']//p"); //price
            var aname = authorName[0].InnerText;

            var author = doc.DocumentNode.SelectNodes("//div[@class='col author_des']//div"); //price
            var authorDesc = author[0].InnerText;

            return View();
        }

        public IActionResult Privacy()
        {
            return View();
        }

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }
    }
}