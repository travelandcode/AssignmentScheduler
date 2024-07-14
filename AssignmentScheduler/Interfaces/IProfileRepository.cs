using AssignmentScheduler.Models;
using MongoDB.Bson;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AssignmentScheduler.Interfaces
{
    public interface IProfileRepository
    {
        Task<Profile> GetUserProfile(ObjectId id);
    }
}
