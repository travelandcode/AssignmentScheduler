using AssignmentScheduler.Interfaces;
using AssignmentScheduler.Models;
using ClosedXML.Excel;

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

        public async Task<byte[]> GenerateAssignmentSchedule(string lastAssignedName)
        {
            var assignments = await _assignmentRepository.GetAllAssignments();
            var month = await _monthRepository.GetNextMonth();
            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add();

                // Create Labels
                var labelStyle = workbook.Style;
                labelStyle.Font.Bold = true;
                labelStyle.Font.FontName = "Calibri";
                labelStyle.Font.FontSize = 11;
                labelStyle.Fill.BackgroundColor = XLColor.Yellow;

                var dateCell = worksheet.Cell(1, 1).SetValue("Diet");
                dateCell.Style = labelStyle;

                // Apply a border around the date cell
                dateCell.Style.Border.OutsideBorder = XLBorderStyleValues.Thick; // Use Thick for bold border
                dateCell.Style.Border.OutsideBorderColor = XLColor.Black;

                for (int i = 0; i < assignments.Count; i++)
                {
                    var assignmentCell = worksheet.Cell(1, i + 2).SetValue(assignments[i].name);
                    assignmentCell.Style = labelStyle;

                    // Apply a border around the header cells
                    assignmentCell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    assignmentCell.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                    assignmentCell.Style.Border.OutsideBorder = XLBorderStyleValues.Thick; // Use Thick for bold border
                    assignmentCell.Style.Border.OutsideBorderColor = XLColor.Black;
                }

                // Assign men to appropriate assignments
                DateTime firstDayOfNextMonth = new DateTime(DateTime.Now.Year, DateTime.Now.AddMonths(1).Month, 1);
                int row = 2;
                

                var profileAssignments = await _profileAssignmentRepository.GetProfileAssignments();

                // Find the last assigned name in the queue
                var menList = await _menRepository.GetAllMen();

                // Shuffle the list of men
                var rng = new Random();
                menList = menList.OrderBy(m => rng.Next()).ToList();

                // Find the last assigned name in the shuffled list
                var lastAssignedIndex = menList.FindIndex(m => m.fullname == lastAssignedName);

                // Reorder the list to start after the last assigned name
                if (lastAssignedIndex != -1)
                {
                    menList = menList.Skip(lastAssignedIndex + 5).Concat(menList.Take(lastAssignedIndex + 5)).ToList();
                }

                var assignmentQueue = new Queue<Men>(menList);

                // Track which men have been assigned to special tasks
                var manSpecialAssignments = new Dictionary<string, HashSet<string>>();

                for (int day = 1; day <= DateTime.DaysInMonth(firstDayOfNextMonth.Year, firstDayOfNextMonth.Month); day++)
                {
                    DateTime currentDate = new DateTime(firstDayOfNextMonth.Year, firstDayOfNextMonth.Month, day);

                    // Only process Sundays and Mondays
                    if (currentDate.DayOfWeek == DayOfWeek.Sunday || currentDate.DayOfWeek == DayOfWeek.Monday)
                    {
                        worksheet.Range(row, 1, row + 1, 1).Merge().SetValue(month + " " + currentDate.Day.ToString());
                        worksheet.Cell(row, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                        worksheet.Cell(row, 1).Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                        worksheet.Cell(row, 1).Style.Font.Bold = true; // Make the text bold

                        // Apply a border around the merged cell
                        worksheet.Range(row, 1, row + 1, 1).Style.Border.OutsideBorder = XLBorderStyleValues.Thick; // Use Thick for bold border
                        worksheet.Range(row, 1, row + 1, 1).Style.Border.OutsideBorderColor = XLColor.Black;


                        if (currentDate.DayOfWeek == DayOfWeek.Monday)
                        {
                            // Merge cells for "Wiik Die Miitn"
                            worksheet.Range(row, 2, row + 1, 3).Merge().Value = "Wiik Die Miitn";
                            worksheet.Cell(row, 2).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                            worksheet.Cell(row, 2).Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                            worksheet.Cell(row, 2).Style.Font.Bold = true; // Make the text bold

                            // Apply a gray fill (hex #D9D9D9) to the Monday row
                            worksheet.Range(row, 1, row + 1, assignments.Count + 1).Style.Fill.BackgroundColor = XLColor.FromHtml("#D9D9D9");


                            // Apply a single border around the merged cells
                            worksheet.Range(row, 2, row + 1, 3).Style.Border.OutsideBorder = XLBorderStyleValues.Thick; // Use Thick for bold border
                            worksheet.Range(row, 2, row + 1, 3).Style.Border.OutsideBorderColor = XLColor.Black;


                            // Assign the rest of the tasks
                            AssignTasks(worksheet, row, assignmentQueue, profileAssignments, assignments, startColumn: 4, endColumn: assignments.Count + 1, manSpecialAssignments);
                        }
                        else if (currentDate.DayOfWeek == DayOfWeek.Sunday)
                        {
                            // Assign tasks for Sunday
                            AssignTasks(worksheet, row, assignmentQueue, profileAssignments, assignments, startColumn: 2, endColumn: assignments.Count + 1, manSpecialAssignments);
                        }

                        row += 2; // Move to the next pair of rows for the next day
                    }
                }

                worksheet.Column(1).Width = 9; // Diet column

                // Set the width of all other columns to 18 characters
                for (int col = 2; col <= assignments.Count + 1; col++) // +1 for the total columns including Diet
                {
                    worksheet.Column(col).Width = 18; // Other columns
                }

                worksheet.Row(1).InsertRowsAbove(4);

                // Define the number of columns to merge across (adjust as needed)
                int totalColumns = assignments.Count + 1;

                // Add the specified texts in the new rows with appropriate styles
                var header1 = worksheet.Range(1, 1, 1, totalColumns).Merge();
                header1.SetValue("PAPIIN KANGRIGIESHAN A JEUOVA WIKNES");
                header1.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                header1.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                header1.Style.Font.Bold = true;

                var header2 = worksheet.Range(2, 1, 2, totalColumns).Merge();
                header2.SetValue("ATENDANT ASAINMENT SKEDYUUL");
                header2.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                header2.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                header2.Style.Font.Bold = true;

                var header3 = worksheet.Range(3, 1, 3, totalColumns).Merge();
                header3.SetValue("\"Bot Mek Shuor Se Unu Du Evriting Iina Di Prapa Wie Jos Laik Ou Gad Se It Fi Go\" - 1 Kor. 14:40");
                header3.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                header3.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                header3.Style.Font.Bold = true;
                header3.Style.Font.Underline = XLFontUnderlineValues.Single;

                // Create space for the month/year by inserting a row below row 5
                worksheet.Row(6).InsertRowsAbove(1); // Insert new row 6

                // Add the month and year merged cell in the new sixth row
                var monthYearCell = worksheet.Range(6, 1, 6, totalColumns).Merge();
                monthYearCell.SetValue(month + " " + DateTime.Now.Year.ToString());
                monthYearCell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                monthYearCell.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                monthYearCell.Style.Font.Bold = true; // Make the text bold
                monthYearCell.Style.Fill.BackgroundColor = XLColor.FromHtml("#92D050"); // Set background color
                monthYearCell.Style.Font.Underline = XLFontUnderlineValues.Single;

                using (var stream = new MemoryStream())
                {
                    workbook.SaveAs(stream);
                    return stream.ToArray();
                }
            }
        }

        private void AssignTasks(IXLWorksheet worksheet, int row, Queue<Men> assignmentQueue, List<ProfileAssignment> profileAssignments, List<Assignment> assignments, int startColumn, int endColumn, Dictionary<string, HashSet<string>> manSpecialAssignments)
        {
            for (int col = startColumn; col <= endColumn; col++)
            {
                var assignment = assignments[col - 2];
                var profileIds = profileAssignments.Where(pa => pa.assignmentid == assignment.Id).Select(pa => pa.profileid).ToHashSet();

                while (assignmentQueue.Count > 0)
                {
                    var man = assignmentQueue.Dequeue();

                    // Check if the task is special and if the man has been assigned to it
                    if (RequiresSpecialAssignment(assignment.name) && manSpecialAssignments.TryGetValue(man.fullname, out var assignedTasks) && assignedTasks.Contains(assignment.name))
                    {
                        assignmentQueue.Enqueue(man); // Re-add the man to the queue if already assigned to this task
                        continue;
                    }

                    if (profileIds.Contains(man.profileid))
                    {
                        worksheet.Cell(row, col).Value = man.fullname;
                        assignmentQueue.Enqueue(man); // Re-add the man to the queue

                        // Add to the man's special assignments
                        if (RequiresSpecialAssignment(assignment.name))
                        {
                            if (!manSpecialAssignments.ContainsKey(man.fullname))
                            {
                                manSpecialAssignments[man.fullname] = new HashSet<string>();
                            }
                            manSpecialAssignments[man.fullname].Add(assignment.name);
                        }
                        break;
                    }
                    assignmentQueue.Enqueue(man); // Re-add if not assigned
                }

                if (RequiresTwoMen(assignment.name))
                {
                    // Assign the second man
                    while (assignmentQueue.Count > 0)
                    {
                        var secondMan = assignmentQueue.Dequeue();

                        // Check if the task is special and if the man has been assigned to it
                        if (RequiresSpecialAssignment(assignment.name) && manSpecialAssignments.TryGetValue(secondMan.fullname, out var assignedTasks) && assignedTasks.Contains(assignment.name))
                        {
                            assignmentQueue.Enqueue(secondMan); // Re-add the man to the queue if already assigned to this task
                            continue;
                        }

                        if (profileIds.Contains(secondMan.profileid))
                        {
                            worksheet.Cell(row + 1, col).Value = secondMan.fullname; // Assign the second man below
                            worksheet.Cell(row + 1, col).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                            worksheet.Cell(row + 1, col).Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                            assignmentQueue.Enqueue(secondMan); // Re-add the second man to the queue

                            // Add to the man's special assignments
                            if (RequiresSpecialAssignment(assignment.name))
                            {
                                if (!manSpecialAssignments.ContainsKey(secondMan.fullname))
                                {
                                    manSpecialAssignments[secondMan.fullname] = new HashSet<string>();
                                }
                                manSpecialAssignments[secondMan.fullname].Add(assignment.name);
                            }
                            break;
                        }
                        assignmentQueue.Enqueue(secondMan); // Re-add if not assigned
                    }

                    // Merge the cells vertically after assigning the second man
                    worksheet.Cell(row, col).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    worksheet.Cell(row, col).Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;

                    // Apply a single border around the merged cells
                    worksheet.Range(row, col, row + 1, col).Style.Border.OutsideBorder = XLBorderStyleValues.Thick; // Use Thick for bold border
                    worksheet.Range(row, col, row + 1, col).Style.Border.OutsideBorderColor = XLColor.Black;
                }
                else
                {
                    // Merge the cells if only one man is required
                    worksheet.Range(row, col, row + 1, col).Merge();
                    worksheet.Cell(row, col).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    worksheet.Cell(row, col).Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;

                    // Apply a single border around the merged cells
                    worksheet.Range(row, col, row + 1, col).Style.Border.OutsideBorder = XLBorderStyleValues.Thick; // Use Thick for bold border
                    worksheet.Range(row, col, row + 1, col).Style.Border.OutsideBorderColor = XLColor.Black;
                }
            }
        }

        private bool RequiresSpecialAssignment(string assignmentName)
        {
            return new HashSet<string> { "Chierman", "Wachtowa Riida" }.Contains(assignmentName);
        }

        private bool RequiresTwoMen(string assignmentName)
        {
            return !new HashSet<string> { "Chierman","Wachtowa Riida", "Platfaam"}.Contains(assignmentName);
        }
    }
}
