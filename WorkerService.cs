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
                using(var wb = new XLWorkbook())
                {
                    var ws = wb.Worksheets.Add("My first wb with closed xml");
                    ws.Cell(1,1).Value = "Um texto qualquer";
                    ws.Cell("A2").Value = "OUTRO TEXTO";
                    wb.SaveAs("./documents/MySimpleTest.xlsx");
                }
                _logger.LogInformation("Worker running at: {time}", DateTimeOffset.Now);
                await Task.Delay(2000);

                using(var wb = new XLWorkbook("./documents/MySimpleTest.xlsx"))
                {
                    var ws = wb.Worksheet("My first wb with closed xml");
                    Console.WriteLine(ws.Cell(1,1).Value);
                }
                
            }
        }
    }
}