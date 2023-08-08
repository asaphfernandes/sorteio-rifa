using System.Diagnostics;
using App.Models;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.AspNetCore.Mvc;

namespace App.Controllers
{
    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> _logger;

        public HomeController(ILogger<HomeController> logger)
        {
            _logger = logger;
        }

        private List<string> _numbers = new List<string>();
        private List<RifaViewModel> _rifas = new List<RifaViewModel>();

        public IActionResult Index()
        {
            return View();
        }

        public IActionResult Privacy()
        {
            return View();
        }

        [HttpPost, Route("upload")]
        public async Task<IActionResult> UploadFile(IFormFile file)
        {
            var memoryStream = new MemoryStream();
            await file.CopyToAsync(memoryStream);

            LoadDia12(memoryStream);

            // return file path
            return Ok(file.Name);
        }

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }

        private void LoadDia12(Stream file)
        {
            var sheetName = "Dia 12";
            // Open the spreadsheet document for read-only access.
            using SpreadsheetDocument document = SpreadsheetDocument.Open(file, false);

            // Retrieve a reference to the workbook part.
            WorkbookPart wbPart = document.WorkbookPart;

            // Find the sheet with the supplied name, and then use that 
            // Sheet object to retrieve a reference to the first worksheet.
            Sheet theSheet = wbPart.Workbook.Descendants<Sheet>()
              .Where(s => s.Name == sheetName).FirstOrDefault();

            if (theSheet == null)
                return;

            // Retrieve a reference to the worksheet part.
            var wsPart = (WorksheetPart)wbPart.GetPartById(theSheet.Id);

            var theRows = wsPart.Worksheet.Descendants<Row>().ToList();

            int? nomePos = null;
            int? quantidadePos = null;
            int? telefonePos = null;

            foreach (var row in theRows)
            {
                var cs = row.Descendants<Cell>().ToList();
                if (!nomePos.HasValue)
                {
                    var aux = 0;
                    foreach (var c in cs)
                    {
                        var v = GetValue(wbPart, c);
                        if (v == "Pessoas")
                            nomePos = aux;
                        if (v == "Quantidade")
                            quantidadePos = aux;
                        if (v == "Telefone")
                            telefonePos = aux;
                        aux++;
                    }
                }
                else
                {
                    var rifa = new RifaViewModel
                    {
                        Nome = GetValue(wbPart, cs[nomePos.Value]),
                        Quantidade = Convert.ToInt32(GetValue(wbPart, cs[quantidadePos.Value])),
                        Telefone = GetValue(wbPart, cs[telefonePos.Value]),
                    };

                    _rifas.Add(rifa);
                }
            }
        }

        private static string GetValue(WorkbookPart wbPart, Cell theCell)
        {
            string value = null;
            if (theCell != null)
            {
                value = theCell.InnerText;

                // If the cell represents an integer number, you are done. 
                // For dates, this code returns the serialized value that 
                // represents the date. The code handles strings and 
                // Booleans individually. For shared strings, the code 
                // looks up the corresponding value in the shared string 
                // table. For Booleans, the code converts the value into 
                // the words TRUE or FALSE.
                if (theCell.DataType != null)
                {
                    switch (theCell.DataType.Value)
                    {
                        case CellValues.SharedString:

                            // For shared strings, look up the value in the
                            // shared strings table.
                            var stringTable =
                                wbPart.GetPartsOfType<SharedStringTablePart>()
                                .FirstOrDefault();

                            // If the shared string table is missing, something 
                            // is wrong. Return the index that is in
                            // the cell. Otherwise, look up the correct text in 
                            // the table.
                            if (stringTable != null)
                            {
                                value =
                                    stringTable.SharedStringTable
                                    .ElementAt(int.Parse(value)).InnerText;
                            }
                            break;

                        case CellValues.Date:

                            // For shared strings, look up the value in the
                            // shared strings table.
                            var stringTable2 =
                                wbPart.GetPartsOfType<SharedStringTablePart>()
                                .FirstOrDefault();

                            // If the shared string table is missing, something 
                            // is wrong. Return the index that is in
                            // the cell. Otherwise, look up the correct text in 
                            // the table.
                            if (stringTable2 != null)
                            {
                                value =
                                    stringTable2.SharedStringTable
                                    .ElementAt(int.Parse(value)).InnerText;
                            }
                            break;

                        case CellValues.Boolean:
                            value = value switch
                            {
                                "0" => "FALSE",
                                _ => "TRUE",
                            };
                            break;
                    }
                }
            }

            return value;
        }
    }

    public class RifaViewModel
    {
        public string Nome { get; set; }
        public int Quantidade { get; set; }
        public string Telefone { get; set; }
    }
}