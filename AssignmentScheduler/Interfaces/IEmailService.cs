using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AssignmentScheduler.Interfaces
{
    public interface IEmailService
    {
        Task SendEmail(string[] emails, string subject, List<byte[]> attachments);
    }
}
