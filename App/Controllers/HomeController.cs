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
        private List<ParticipanteViewModel> _participantes = new List<ParticipanteViewModel>();
        private List<RifaViewModel> _rifas = new List<RifaViewModel>();
        private List<RifaViewModel> _premios1;
        private List<RifaViewModel> _premios2;
        private List<RifaViewModel> _premios3;

        private ParticipanteViewModel _ganhador1;
        private ParticipanteViewModel _ganhador2;
        private ParticipanteViewModel _ganhador3;

        private int _premio = 0;

        public IActionResult Index()
        {
            return View();
        }

        public IActionResult Privacy()
        {
            return View();
        }

        [HttpPost, Route("upload")]
        public async Task<IActionResult> Result(IFormFile file)
        {
            var memoryStream = new MemoryStream();
            await file.CopyToAsync(memoryStream);

            LoadParticipantes(memoryStream);

            var quantidade = _participantes.Sum(s => s.Quantidade);
            _premio = quantidade < 125 ? 2 : 3;
            
            LoadPremios();
            Ganhadores();

            var response = new ResponseViewModel
            {
                Premio = _premio,
                Quantidade = quantidade,
                Ganhador1 = _ganhador1,
                Ganhador2 = _ganhador2,
                Ganhador3 = _ganhador3,
            };

            // return file path
            return View(response);
        }

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }

        private void LoadParticipantes(Stream file)
        {
            // Open the spreadsheet document for read-only access.
            using SpreadsheetDocument document = SpreadsheetDocument.Open(file, false);

            // Retrieve a reference to the workbook part.
            WorkbookPart wbPart = document.WorkbookPart;

            // Find the sheet with the supplied name, and then use that 
            // Sheet object to retrieve a reference to the first worksheet.
            Sheet theSheet = wbPart.Workbook.Descendants<Sheet>().FirstOrDefault();

            if (theSheet == null)
                return;

            // Retrieve a reference to the worksheet part.
            var wsPart = (WorksheetPart)wbPart.GetPartById(theSheet.Id);

            var theRows = wsPart.Worksheet.Descendants<Row>().ToList();

            int? nomePos = null;
            int? quantidadePos = null;
            int? telefonePos = null;

            var quantidade = 0;
            foreach (var row in theRows)
            {
                var cs = row.Descendants<Cell>().ToList();
                if (!nomePos.HasValue)
                {
                    var aux = 0;
                    foreach (var c in cs)
                    {
                        var v = GetValue(wbPart, c);
                        if (v == "Nome")
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
                    var rifa = new ParticipanteViewModel
                    {
                        Nome = GetValue(wbPart, cs[nomePos.Value]),
                        Quantidade = Convert.ToInt32(GetValue(wbPart, cs[quantidadePos.Value])),
                        Telefone = GetValue(wbPart, cs[telefonePos.Value]),
                    };

                    quantidade += rifa.Quantidade;

                    _participantes.Add(rifa);
                }
            }
        }

        private void LoadRifas()
        {
            _rifas = new List<RifaViewModel>();
            foreach (var participante in _participantes)
            {
                for (int j = 0; j < participante.Quantidade; j++)
                {
                    _rifas.Add(new RifaViewModel
                    {
                        Telefone = participante.Telefone,
                        Numero = GetNumber(),
                    });
                }
            }
        }

        private void LoadPremios()
        {
            _premios1 = new List<RifaViewModel>();
            _premios2 = new List<RifaViewModel>();
            _premios3 = new List<RifaViewModel>();

            for (int i = 0; i < 10; i++)
            {
                LoadRifas();


                foreach (var rifa in _rifas)
                {
                    if (rifa.Numero % _premio == 0)
                        _premios1.Add(rifa);
                    if (rifa.Numero % _premio == 1)
                        _premios2.Add(rifa);
                    if (rifa.Numero % _premio == 2)
                        _premios3.Add(rifa);
                }

                var count12 = _premios1.Count(w => !_premios2.Any(a => w.Telefone == a.Telefone));
                var count21 = _premios2.Count(w => !_premios1.Any(a => w.Telefone == a.Telefone));

                var count13 = _premios1.Count(w => !_premios3.Any(a => w.Telefone == a.Telefone));
                var count23 = _premios2.Count(w => !_premios3.Any(a => w.Telefone == a.Telefone));
                var count31 = _premios3.Count(w => !_premios1.Any(a => w.Telefone == a.Telefone));
                var count32 = _premios3.Count(w => !_premios2.Any(a => w.Telefone == a.Telefone));

                if (_premio == 2 && count12 > 0 && count21 > 0)
                    break;

                if (_premio == 3 && count12 > 0 && count21 > 0 && count13 > 0 && count23 > 0 && count31 > 0 && count32 > 0)
                    break;

                _premios1 = new List<RifaViewModel>();
                _premios2 = new List<RifaViewModel>();
                _premios3 = new List<RifaViewModel>();
            }

            if (!_premios1.Any())
                throw new Exception("Não foi possível gerar os prêmios");
        }

        private void Ganhadores()
        {
            foreach (var premio1 in _premios1)
                premio1.Numero = GetNumber();

            var ganhador1 = _premios1.OrderBy(o => o.Numero).First();
            _ganhador1 = _participantes.First(w => w.Telefone == ganhador1.Telefone);

            _premios2 = _premios2.Where(w => w.Telefone != ganhador1.Telefone).ToList();
            _premios3 = _premios3.Where(w => w.Telefone != ganhador1.Telefone).ToList();

            foreach (var premio2 in _premios2)
                premio2.Numero = GetNumber();

            var ganhador2 = _premios2.OrderBy(o => o.Numero).First();
            _ganhador2 = _participantes.First(w => w.Telefone == ganhador2.Telefone);

            _premios3 = _premios3.Where(w => w.Telefone != ganhador2.Telefone).ToList();

            foreach (var premio3 in _premios3)
                premio3.Numero = GetNumber();

            if (_premio == 3)
            {
                var ganhador3 = _premios3.OrderBy(o => o.Numero).First();
                _ganhador3 = _participantes.First(w => w.Telefone == ganhador3.Telefone);
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

        public int GetNumber()
        {
            var r = new Random();
            var n = r.Next(ushort.MaxValue);
            return n;
        }
    }
}