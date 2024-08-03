using Amazon.S3;
using Amazon.S3.Model;
using Amazon.S3.Transfer;
using AssignmentScheduler.Interfaces;
using ClosedXML.Excel;


namespace AssignmentScheduler.Services
{
    public class AwsS3Service : IAwsS3
    {
        private readonly string bucketName;
        private readonly IAmazonS3 s3Client;
        private readonly IMonthRepository _monthRepository;

        public AwsS3Service(string bucketName, string accessKey, string secretKey, string region)
        {
            this.bucketName = bucketName;
            this.s3Client = new AmazonS3Client(accessKey, secretKey, Amazon.RegionEndpoint.GetBySystemName(region));
        }

        public async Task<String> GetFile(string month)
        {
            try
            {
                // Create a request to get the Excel file from the S3 bucket
                var request = new GetObjectRequest
                {
                    BucketName = bucketName,
                    Key = $"Papiin Asainment Skedyuul ({month})"
                };

                // Retrieve the object from S3
                using (var response = await s3Client.GetObjectAsync(request))
                using (var responseStream = response.ResponseStream)
                using (var memoryStream = new MemoryStream())
                {
                    // Copy the response stream to a memory stream
                    await responseStream.CopyToAsync(memoryStream);

                    // Reset the position of the memory stream to the beginning
                    memoryStream.Position = 0;

                    // Use ClosedXML to read the Excel file
                    using (var workbook = new XLWorkbook(memoryStream))
                    {
                        // Access the first worksheet in the Excel file
                        var worksheet = workbook.Worksheet(1);

                        // Find the last used row in the worksheet
                        var lastRow = worksheet.LastRowUsed();
                        if (lastRow != null)
                        {
                            // Access the last cell in the last row
                            var lastCell = lastRow.LastCellUsed();

                            // Store the value of the last cell
                            var lastCellValue = lastCell?.GetValue<string>();

                            // Return the last cell value
                            return lastCellValue;
                        }
                        else
                        {
                            Console.WriteLine("The worksheet is empty.");
                            return null;
                        }
                    }

                }
            }
            catch (AmazonS3Exception e)
            {
                Console.WriteLine("Error encountered on server. Message: '{0}' when reading object", e.Message);
            }
            catch (Exception e)
            {
                Console.WriteLine("Unknown encountered on server. Message: '{0}' when reading object", e.Message);
            }

            return null;
        }

        public async Task UploadFileAsync(byte[] fileBytes, string fileName, string contentType)
        {
            try
            {
                // Create a TransferUtility instance for high-level operations
                using var fileTransferUtility = new TransferUtility(s3Client);

                using (var memoryStream = new MemoryStream(fileBytes))
                {
                    // Use TransferUtilityUploadRequest to specify upload details
                    var uploadRequest = new TransferUtilityUploadRequest
                    {
                        InputStream = memoryStream,
                        Key = fileName,
                        BucketName = bucketName,
                        ContentType = contentType
                    };

                    // Upload the file to S3
                    await fileTransferUtility.UploadAsync(uploadRequest);

                    Console.WriteLine("File uploaded successfully.");
                }
            }
            catch (AmazonS3Exception e)
            {
                Console.WriteLine("Error encountered on server. Message:'{0}' when writing an object", e.Message);
            }
            catch (Exception e)
            {
                Console.WriteLine("Unknown encountered on server. Message:'{0}' when writing an object", e.Message);
            }
        }

    }
}
