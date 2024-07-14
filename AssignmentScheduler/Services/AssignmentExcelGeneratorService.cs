using AssignmentScheduler.Interfaces;
using AssignmentScheduler.Models;
using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AssignmentScheduler.Services
{
    public class AssignmentExcelGenerator : IAssignmentExcelGenerator
    {
        private readonly IMenRepository _menRepository;
        private readonly IProfileAssignmentRepository _profileAssignmentRepository;
        private readonly IProfileRepository _profileRepository;
        private readonly IAssignmentRepository _assignmentRepository;
        private readonly IMonthRepository _monthRepository;

        public AssignmentExcelGenerator(IMenRepository menRepository, IProfileAssignmentRepository profileAssignmentRepository, IProfileRepository profileRepository, IAssignmentRepository assignmentRepository, IMonthRepository monthRepository)
        {
            _menRepository = menRepository;
            _profileAssignmentRepository = profileAssignmentRepository;
            _profileRepository = profileRepository;
            _assignmentRepository = assignmentRepository;
            _monthRepository = monthRepository;
        }

        public async Task<byte[]> GenerateAssignmentSchedule()
        {
            var assignments = await _assignmentRepository.GetAllAssignments();
            var men = await _menRepository.GetAllMen();
            var profileAssignments = await _profileAssignmentRepository.GetProfileAssignments();

            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add();

               //Create Labels
                var labelStyle = workbook.Style;
                labelStyle.Font.Bold = true;
                labelStyle.Font.FontName = "Calibri";
                labelStyle.Font.FontSize = 11;
                labelStyle.Fill.BackgroundColor = XLColor.Yellow;

                var dateCell = worksheet.Cell(1, 1).SetValue("Diet");
                dateCell.Style = labelStyle;

                for (int i = 0; i < assignments.Count; i++)
                {
                    var assignmentCell = worksheet.Cell(1, i + 2).SetValue(assignments[i].Name);
                    assignmentCell.Style = labelStyle;
                }

                //Assign men to appropriate assignments
                var month = await _monthRepository.GetNextMonth();

                // Get the first day of the next month
                DateTime firstDayOfNextMonth = new DateTime(DateTime.Now.AddMonths(1).Year, DateTime.Now.AddMonths(1).Month, 1);
                int row = 2;

                var assignmentQueue = new Queue<Men>(men);

                // Iterate through the days of the next month
                for (int day = 1; day <= DateTime.DaysInMonth(firstDayOfNextMonth.Year, firstDayOfNextMonth.Month); day++)
                {
                    DateTime currentDate = new DateTime(firstDayOfNextMonth.Year, firstDayOfNextMonth.Month, day);
                    worksheet.Cell(row, 1).SetValue(month + " " + currentDate.ToString("d"));

                    if (currentDate.DayOfWeek == DayOfWeek.Monday)
                    {
                        // Merge cells for "Wiik Die Miitn"
                        worksheet.Range(row, 2, row, 3).Merge().Value = "Wiik Die Miitn";
                        worksheet.Cell(row, 2).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

                        // Assign the rest of the tasks
                        AssignTasks(worksheet, row, assignmentQueue, profileAssignments, assignments, startColumn: 4, endColumn: assignments.Count + 1);
                    }
                    else if (currentDate.DayOfWeek == DayOfWeek.Sunday)
                    {
                        // Assign tasks for Sunday
                        AssignTasks(worksheet, row, assignmentQueue, profileAssignments, assignments, startColumn: 2, endColumn: assignments.Count + 1);
                    }

                    row++;
                }

                using (var stream = new MemoryStream())
                {
                    workbook.SaveAs(stream);
                    return stream.ToArray();
                }
            }
        }

        private void AssignTasks (IXLWorksheet worksheet, int row, Queue<Men> assignmentQueue, List<ProfileAssignment> profileAssignments, List<Assignment> assignments, int startColumn, int endColumn)
        {
            for (int col = startColumn; col <= endColumn; col++)
            {
                var assignment = assignments[col - 2];
                var profileIds = profileAssignments.Where(pa => pa.AssignmentId == assignment.Id).Select(pa => pa.ProfileId).ToHashSet();

                while (assignmentQueue.Count > 0)
                {
                    var man = assignmentQueue.Dequeue();
                    if(profileIds.Contains(man.ProfileId))
                    {
                        worksheet.Cell(row, col).Value = man.FullName;
                        assignmentQueue.Enqueue(man);
                        break; 
                    }
                    assignmentQueue.Enqueue(man);
                }

                if(!RequiresTwoMen(assignment.Name))
                {
                    worksheet.Cell(row, col).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    worksheet.Range(row, col, row, col + 1).Merge();
                    col++;
                }
            }
        }

        private bool RequiresTwoMen(string assignmentName)
        {
            return new HashSet<string> { "Chierman","Wachtowa Riida", "Platfaam"}.Contains(assignmentName);
        }
    }
}
