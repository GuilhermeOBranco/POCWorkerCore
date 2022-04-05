using System;
using System.Threading;
using System.Threading.Tasks;
using ClosedXML.Excel;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;

namespace POCWorkerCore
{
    public class WorkerService : BackgroundService
    {
        private readonly ILogger<WorkerService> _logger;

        public WorkerService(ILogger<WorkerService> logger)
        {
            _logger = logger;
        }
        protected override async Task ExecuteAsync(CancellationToken stoppingToken)
        {
            while(!stoppingToken.IsCancellationRequested)
            {
                string nome = string.Empty;
                using(var wb = new XLWorkbook())
                {
                    var ws = wb.Worksheets.Add("My first wb with closed xml");
                    ws.Cell(1,1).Value = "Um texto qualquer";
                    ws.Cell("A2").Value = "OUTRO TEXTO";
                    ws.Cell("D1").Value = "OUTRO TEXTO QUALQUER";
                    nome = @$"./documents/MySimpleTest_{DateTime.Now.Minute}_{DateTime.Now.Second}.xlsx";
                    wb.SaveAs(nome);
                }
                
                _logger.LogInformation("Worker running at: {time}", DateTimeOffset.Now);
                
                await Task.Delay(2000);

                using(var wb = new XLWorkbook(nome))
                {
                    var ws = wb.Worksheet(1);
                    Console.WriteLine(ws.Cell(1,1).Value);
                }

                using(var wb = new XLWorkbook(nome))
                {
                    var ws = wb.Worksheet(1);
                    var range = ws.Range("A1:D78");
                    string teste = range.Cell(1,4).Value.ToString();
                    Console.WriteLine(teste);
                }
                
            }
        }
    }
}