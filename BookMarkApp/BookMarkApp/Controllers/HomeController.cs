using BookMarkApp.Models;
using HtmlAgilityPack;
using Microsoft.AspNetCore.Mvc;
using System.Diagnostics;
using System.IO;
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

            return View();
        }
        [HttpPost]
        public async Task<IActionResult> Index(string urls)
        {
            try
            {
                var books = new List<Book>();
                var urlArray = urls.Trim().Split(',');
                if (urlArray.Length ==1)
                {
                    urlArray = urls.Split("\r\n");
                    urlArray = urlArray.Where(f=>f.Length > 1).ToArray();
                }
                // var url = "https://www.rokomari.com/book/15991/bangali-musulmaner-mon";
                foreach (var url in urlArray)
                {
                    var urlstring = url.Replace("\r\n", "");
                    var book = await BookInfo.GetBookDetails(urlstring.Trim());
                    books.Add(book);
                }

                var bytes = BookInfo.makeExcel(books);

                var content = new System.IO.MemoryStream(bytes);
                var contentType = "APPLICATION/octet-stream";
                var fileName = "something.csv";
                return File(content, contentType, fileName); //email to the copy
            }
            catch (Exception ex)
            {
                return Content(ex.Message);
            }
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