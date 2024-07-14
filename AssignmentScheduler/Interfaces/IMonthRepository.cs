using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AssignmentScheduler.Interfaces
{
    public interface IMonthRepository
    {
        Task<string> GetNextMonth();
    }
}
