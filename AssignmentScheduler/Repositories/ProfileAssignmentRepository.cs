using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using MongoDB.Bson;
using MongoDB.Driver;

using AssignmentScheduler.Interfaces;
using AssignmentScheduler.Models;

namespace AssignmentScheduler.Repositories
{
    public class ProfileAssignmentRepository : IProfileAssignmentRepository
    {
        private readonly IMongoCollection<Profile> _profileCollection;
        private readonly IMongoCollection<Assignment> _assignmentCollection;
        private readonly IMongoCollection<ProfileAssignment> _profileAssignmentCollection;
       public ProfileAssignmentRepository(IMongoDatabase db) 
       {
            _profileCollection = db.GetCollection<Profile>("Profile");
            _assignmentCollection = db.GetCollection<Assignment>("Assignment");
            _profileAssignmentCollection = db.GetCollection<ProfileAssignment>("ProfileAssignment");
       }
        public async Task<List<string>> GetUserAssignments (string profileName)
        {
            Profile profile = await _profileCollection.Find(p => p.Name == profileName).FirstOrDefaultAsync();
            if (profile == null) { return null; }

            var profileAssignments = await _profileAssignmentCollection.Find(pa => pa.ProfileId == profile.Id).ToListAsync();
            var assignmentIds = profileAssignments.Select(pa => pa.AssignmentId).ToList();
            var assignments = await _assignmentCollection.Find(a => assignmentIds.Contains(a.Id)).ToListAsync();

            return assignments.Select(assignment => assignment.Name).ToList();
        }

        public async Task<List<ProfileAssignment>> GetProfileAssignments()
        {
            return await _profileAssignmentCollection.Find(_ => true).ToListAsync(); // Retrieves all profile assignments
        }
    }
}
