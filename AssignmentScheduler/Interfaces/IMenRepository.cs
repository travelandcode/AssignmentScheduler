using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using AssignmentScheduler.Models;

namespace AssignmentScheduler.Interfaces
{
    public interface IMenRepository
    {
        Task<List<Men>> GetAllMen();
    }
}
