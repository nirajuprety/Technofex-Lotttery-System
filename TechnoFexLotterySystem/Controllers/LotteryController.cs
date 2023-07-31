using Microsoft.AspNetCore.Hosting.Server;
using Microsoft.AspNetCore.Mvc;

namespace TechnoFexLotterySystem.Controllers
{
    public class LotteryController : Controller
    {
        [HttpPost]
        public ActionResult UploadExcel(IFormFile excelFile)
        {
            if (excelFile != null && excelFile.Length > 0)
            {
                // Process the uploaded file here (e.g., save it, read data, etc.)
                // For example, you can save the file to a specific folder on the server:
                string uploadFolderPath = Path.Combine(Directory.GetCurrentDirectory(), "UploadedExcelFiles");
                if (!Directory.Exists(uploadFolderPath))
                {
                    Directory.CreateDirectory(uploadFolderPath);
                }

                string filePath = Path.Combine(uploadFolderPath, excelFile.FileName);

                using (var fileStream = new FileStream(filePath, FileMode.Create))
                {
                    excelFile.CopyTo(fileStream);
                }

                // Add additional logic to handle the uploaded file as needed.

                ViewBag.Message = "Excel file uploaded successfully!";
            }
            else
            {
                ViewBag.Message = "Please choose an Excel file to upload.";
            }

            return View("UploadExcel");
        }

        public ActionResult UploadExcel()
        {
            return View();
        }
        [HttpGet]
        public IActionResult StartProcessing()
        {
            return View();
        }
        // Function to handle after start button
        [HttpPost] // This action will be triggered when the "Start" button is clicked
        public IActionResult StartProcessing(IFormFile excelFile)
        {
            if (excelFile != null && excelFile.Length > 0)
            {
                // Process the uploaded file here (e.g., save it, read data, etc.)
                // For example, you can save the file to a specific folder on the server:
                string uploadFolderPath = Path.Combine(Directory.GetCurrentDirectory(), "UploadedExcelFiles");
                if (!Directory.Exists(uploadFolderPath))
                {
                    Directory.CreateDirectory(uploadFolderPath);
                }

                string filePath = Path.Combine(uploadFolderPath, excelFile.FileName);

                using (var fileStream = new FileStream(filePath, FileMode.Create))
                {
                    excelFile.CopyTo(fileStream);
                }

                // Add additional logic to handle the uploaded file as needed.

                ViewBag.Message = "Excel file uploaded and processing started successfully!";
            }
            else
            {
                ViewBag.Message = "Please choose an Excel file to upload.";
            }

            return View("UploadExcel"); // Redirect back to the CSHTML page
        }
    }
}