using AssignmentScheduler.Interfaces;
using AssignmentScheduler.Models;
using MongoDB.Driver;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AssignmentScheduler.Repositories
{
    public class AssignmentRepository : IAssignmentRepository
    {
        private readonly IMongoCollection<Assignment> _assignmentCollection;

        public AssignmentRepository(IMongoDatabase db)
        {
            _assignmentCollection = db.GetCollection<Assignment>("Assignment");
        }

        public async Task<List<Assignment>> GetAllAssignments()
        {
            var assignments = await _assignmentCollection.Find(_ => true).ToListAsync();
            return assignments;
        }
    }
}
