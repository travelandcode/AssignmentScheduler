using AssignmentScheduler.Interfaces;
using AssignmentScheduler.Models;
using MongoDB.Bson;
using MongoDB.Driver;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AssignmentScheduler.Repositories
{
    public class ProfileRepository : IProfileRepository
    {
        private readonly IMongoCollection<Profile> _profileCollection;

        public ProfileRepository(IMongoDatabase db) 
        {
            _profileCollection = db.GetCollection<Profile>("Profile");
        }

        public async Task<Profile> GetUserProfile(ObjectId id)
        {
            Profile profile = _profileCollection.Find(p => p.Id == id).FirstOrDefault();
            return profile;
        }
    }
}
