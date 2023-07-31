﻿using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Http;
using System.IO;
using System.Linq;
using System;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using TechnoFexLotterySystem.Models;

public class LotterySystemController : Controller
{
    [HttpGet]
    public IActionResult UploadExcel()
    {
        return View();
    }

    [HttpPost]
    public IActionResult StartProcessing(IFormFile excelFile)
    {
        if (excelFile != null && excelFile.Length > 0)
        {
            try
            {
                // Save the uploaded file temporarily on the server
                string tempFilePath = Path.GetTempFileName();
                using (var fileStream = new FileStream(tempFilePath, FileMode.Create))
                {
                    excelFile.CopyTo(fileStream);
                }

                // Perform lottery selection
                LotteryWinner winner = GetRandomWinner(tempFilePath);

                // Delete the temporary file after processing
                System.IO.File.Delete(tempFilePath);

                // Pass the winner's name and number to the DisplayWinner view
                return RedirectToAction("DisplayWinner", new { winnerName = winner.Name, winnerNumber = winner.Number, totalAmount = winner.Amount });
            }
            catch (Exception ex)
            {
                ViewBag.Message = "Error occurred during the lottery selection: " + ex.Message;
            }
        }
        else
        {
            ViewBag.Message = "Please choose an Excel file to upload.";
        }

        return View("UploadExcel");
    }

    private LotteryWinner GetRandomWinner(string filePath)
    {
        LotteryWinner winner = new LotteryWinner();

        // Read the Excel file and select a random row as the winner
        using (SpreadsheetDocument document = SpreadsheetDocument.Open(filePath, false))
        {
            WorkbookPart workbookPart = document.WorkbookPart;
            Sheet sheet = workbookPart.Workbook.Descendants<Sheet>().FirstOrDefault();
            WorksheetPart worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id);
            SheetData sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();

            // Get the total number of rows (excluding header)
            int rowCount = sheetData.Elements<Row>().Count() - 1;

            if (rowCount > 0)
            {
                // Generate a random row index (from 1 to rowCount)
                Random random = new Random();
                int randomRowIndex = random.Next(1, rowCount + 1);

                // Find the selected row and extract the data (assuming column B contains names and column C contains numbers)
                Row selectedRow = sheetData.Elements<Row>().Skip(randomRowIndex).First();
                Cell nameCell = selectedRow.Elements<Cell>().ElementAtOrDefault(1); // Index 1 corresponds to column B
                Cell numberCell = selectedRow.Elements<Cell>().ElementAtOrDefault(2); // Index 2 corresponds to column C
                Cell amountCell = selectedRow.Elements<Cell>().ElementAtOrDefault(3); // Index 2 corresponds to column C
                winner.Name = GetCellValue(workbookPart, nameCell);
                winner.Number = GetCellValue(workbookPart, numberCell);
                winner.Amount= GetCellValue(workbookPart, amountCell);

            }
            else
            {
                winner.Name = "No participants found in the Excel file.";
                winner.Number = string.Empty;
            }
        }

        return winner;
    }

    private string GetCellValue(WorkbookPart workbookPart, Cell cell)
    {
        string value = cell?.InnerText;

        if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
        {
            int ssid = int.Parse(value);
            SharedStringItem ssi = workbookPart.SharedStringTablePart.SharedStringTable.Elements<SharedStringItem>().ElementAt(ssid);
            if (ssi.Text != null)
            {
                value = ssi.Text.Text;
            }
            else if (ssi.InnerText != null)
            {
                value = ssi.InnerText;
            }
            else if (ssi.InnerXml != null)
            {
                value = ssi.InnerXml;
            }
        }

        return value;
    }
    public IActionResult DisplayWinner(string winnerName,string winnerNumber, string totalAmount)
    {
        ViewBag.WinnerName = winnerName;
        ViewBag.WinnerNumber = winnerNumber;
        ViewBag.TotalAmount = totalAmount;
        return View();
    }
}
