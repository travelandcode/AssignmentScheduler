using AssignmentScheduler.Interfaces;
using Microsoft.Extensions.Options;
using System.Globalization;

namespace AssignmentScheduler
{
    public class Worker : BackgroundService
    {
        private readonly IServiceProvider _serviceProvider;
        private Timer _timer;

        public Worker(IServiceProvider serviceProvider)
        {
            _serviceProvider = serviceProvider;
        }

        protected override Task ExecuteAsync(CancellationToken stoppingToken)
        {
            _timer = new Timer(DoJob, null, TimeSpan.Zero, TimeSpan.FromHours(24));
            return Task.CompletedTask;
        }

        private void DoJob(object state)
        {
            SendAssignmentSchedule();
        }

        private async void SendAssignmentSchedule()
        {
            if (DateTime.Now.Day == 21)
            {
                using (var scope = _serviceProvider.CreateScope())
                {
                    var excelGenerator = scope.ServiceProvider.GetRequiredService<IAssignmentExcelGenerator>();
                    var emailService = scope.ServiceProvider.GetRequiredService<IEmailService>();
                    // Calculate the start and end dates for the previous month
                    DateTime now = DateTime.Now;
                    DateTime previousMonth = new DateTime(now.Year, now.Month, 1).AddMonths(-1);
                    string previousMonthFormatted = previousMonth.ToString("MMMM yyyy", CultureInfo.InvariantCulture);

                    //Generate Excel Files using stored procedure
                    List<byte[]> excelAttachments = new List<byte[]> { await excelGenerator.GenerateAssignmentSchedule() };
                    string[] recipientEmails = { "ryonwalkerjnr2001@gmail.com" }; // recipient emails
                    string subject = "Papiin Asainment Skedyuul";

                    // Send email with the attachments
                   

                    // Send email with attachments
                    await emailService.SendEmail(recipientEmails, subject, excelAttachments);
                }
            }
        }

        public override void Dispose()
        {
            _timer?.Dispose();
            base.Dispose();
        }
    }
}
