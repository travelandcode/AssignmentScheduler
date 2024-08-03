using AssignmentScheduler.Interfaces;
using AssignmentScheduler.Repositories;
using Microsoft.Extensions.Options;
using System.Globalization;

namespace AssignmentScheduler
{
    public class Worker : BackgroundService
    {
        private readonly IServiceScopeFactory _serviceScopeFactory;
        private Timer _timer;
        public Worker(IServiceScopeFactory serviceScopeFactory)
        {
            _serviceScopeFactory = serviceScopeFactory;
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
            if (DateTime.Now.Day == 20)
            {
                using (var scope = _serviceScopeFactory.CreateScope())
                {
                    var excelGenerator = scope.ServiceProvider.GetRequiredService<IAssignmentExcelGenerator>();
                    var emailService = scope.ServiceProvider.GetRequiredService<IEmailService>();
                    var s3Service = scope.ServiceProvider.GetRequiredService<IAwsS3>();
                    var monthRepository = scope.ServiceProvider.GetRequiredService<IMonthRepository>();
                    // Calculate the start and end dates for the previous month
                    DateTime now = DateTime.Now; ;
                    DateTime previousMonth = new DateTime(now.Year, now.Month, 1).AddMonths(-1);
                    string previousMonthFormatted = previousMonth.ToString("MMMM yyyy", CultureInfo.InvariantCulture);
                    var nextMonthInPatwa = await monthRepository.GetNextMonth();
                    var prevMonthInPatwa = await monthRepository.GetCurrentMonth();

                   var lastManUsed = await s3Service.GetFile($"{prevMonthInPatwa} {now.Year}");


                    //Generate Excel Files using stored procedure
                    List<byte[]> excelAttachments = new List<byte[]> { await excelGenerator.GenerateAssignmentSchedule(lastManUsed) };
                    string[] recipientEmails = { "ryonwalkerjnr2001@gmail.com" }; // recipient emails
                    string subject = $"Papiin Asainment Skedyuul ({nextMonthInPatwa} {now.Year})";

                    //Upload Excel File to S3 Bucket
                    await s3Service.UploadFileAsync(excelAttachments[0], $"Papiin Asainment Skedyuul ({nextMonthInPatwa} {now.Year})", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");

                    // Send email with attachments
                    await emailService.SendEmail(recipientEmails, subject, excelAttachments,$"{nextMonthInPatwa} {now.Year}");
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
