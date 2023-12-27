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
            var books = new List<Book>();
            var url = "https://www.rokomari.com/book/15991/bangali-musulmaner-mon";
            var book = await BookInfo.GetBookDetails(url);
            books.Add(book);
            var bytes = BookInfo.makeExcel(books);


            var content = new System.IO.MemoryStream(bytes);
            var contentType = "APPLICATION/octet-stream";
            var fileName = "something.csv";
            return File(content, contentType, fileName);
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