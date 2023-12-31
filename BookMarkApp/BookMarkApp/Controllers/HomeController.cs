using BookMarkApp.Models;
using HtmlAgilityPack;
using Microsoft.AspNetCore.Mvc;
using NPOI.POIFS.Crypt.Dsig;
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
        public async Task<IActionResult> Index(string urls, IFormFile file)
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

                foreach (var url in urlArray)
                {
                    var urlstring = url.Replace("\r\n", "");
                    var book = await BookInfo.GetBookDetails(urlstring.Trim());
                    if(book != null)
                    {
                        books.Add(book);
                    }
                }
                if(file != null)
                {
                    var existingBooks = BookInfo.ReadExcel(file);
                    if (books.Any() && existingBooks.Any())
                    {
                        //var firstNotSecond = books.Except(existingBooks).ToList();
                        var secondNotFirst = existingBooks.Except(books).ToList();
                        if (secondNotFirst.Any()) { books.AddRange(secondNotFirst); }
                    }
                    else
                    {
                        books = existingBooks;
                    }
                }

                //var bytes = BookInfo.makeExcel(books);
                string filePath = BookInfo.WriteExcel(books);


                // Offer the new Excel file for download
                return DownloadExcel(filePath);
            }
            catch (Exception ex)
            {
                return Content(ex.Message);
            }
        }
        private IActionResult DownloadExcel(string filePath)
        {
            byte[] fileBytes = System.IO.File.ReadAllBytes(filePath);
            string fileName = @$"bookmark{DateTime.Now.ToShortDateString()}.xlsx";

            return File(fileBytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);
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