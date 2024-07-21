using AssignmentScheduler.Interfaces;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Text;
using System.Threading.Tasks;

namespace AssignmentScheduler.Services
{
    public class EmailService : IEmailService
    {
        public EmailService()
        {
        }

        //public async Task SendEmail(string email, string subject, List<byte[]> attachment)
        //{
        //    var message = new MailMessage
        //    {
        //        From = new MailAddress("ryonwalkerjnr2001@gmail.com", "Ryon Walker Jnr"),
        //        IsBodyHtml = true,
        //        Subject = subject,
        //        Body = "Please see attached."
        //    };




        //    using (var smtp = new SmtpClient())
        //    {

        //        await smtp.SendMailAsync(message);
        //    }
        //}

        public async Task SendEmail(string[] emails, string subject, List<byte[]> attachments)
        {
            var message = new MailMessage
            {
                From = new MailAddress("myonwalkerjnr200@gmail.com", "Ryon Walker Jnr"),
                Subject = subject,
                Body = "Please see attached."
            };

            // Add multiple email addresses
            foreach (var email in emails)
            {
                message.To.Add(email);
            }

            // Attach files
            if (attachments != null && attachments.Count > 0)
            {
                if (attachments != null && attachments.Count > 0)
                {
                    foreach (var attachment in attachments)
                    {
                        if (attachment != null && attachment.Length > 0)
                        {
                            var memoryStream = new MemoryStream(attachment);
                            var mailAttachment = new Attachment(memoryStream, "Papiin Asainment Skedyuul.xlsx");
                            message.Attachments.Add(mailAttachment);

                        }
                    }
                }
            }

            using (var smtp = new SmtpClient("smtp.gmail.com", 587)) // Example for Gmail
            {
                smtp.Credentials = new NetworkCredential("ryonwalkerjnr2001@gmail.com", "huqb yrrp zqnn xefq"); // Use a secure method to handle passwords
                smtp.DeliveryMethod = SmtpDeliveryMethod.Network;
                smtp.UseDefaultCredentials = false;
                smtp.EnableSsl = true;
                

                try
                {
                    await smtp.SendMailAsync(message);
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.ToString());
                }
            }
        }
    }
}
