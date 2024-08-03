using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AssignmentScheduler.Interfaces
{
    public interface IAwsS3
    {
        Task UploadFileAsync(byte[] fileBytes, string fileName, string contentType);

        Task<String> GetFile(string month);
    }
}
