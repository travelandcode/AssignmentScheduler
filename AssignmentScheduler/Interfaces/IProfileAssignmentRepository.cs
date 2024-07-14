using AssignmentScheduler.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AssignmentScheduler.Interfaces
{
    public interface IProfileAssignmentRepository
    {
        Task<List<string>> GetUserAssignments(string profileName);

        Task<List<ProfileAssignment>> GetProfileAssignments();
    }
}
