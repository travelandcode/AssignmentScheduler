using AssignmentScheduler.Models;
using AssignmentScheduler.Repositories;
using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AssignmentScheduler.Interfaces
{
    public interface IAssignmentExcelGenerator
    {
        Task<byte[]> GenerateAssignmentSchedule();

        //void AssignTasks(IXLWorksheet worksheet, int row, Queue<Men> assignmentQueue, List<ProfileAssignment> profileAssignments, List<Assignment> assignments, int startColumn, int endColumn);

        //bool RequiresTwoMen(string assignmentName);
    }
}
