using AssignmentScheduler.Interfaces;
using AssignmentScheduler.Models;
using MongoDB.Driver;

namespace AssignmentScheduler.Repositories
{
    public class MenRepository : IMenRepository
    {
        private readonly IMongoCollection<Men> _menCollection;
        public MenRepository(IMongoDatabase db) 
        {
            _menCollection = db.GetCollection<Men>("Men");
        }

        public async Task<List<Men>> GetAllMen()
        {
            var men = await _menCollection.Find(_ => true).SortBy(m => m.fullname).ToListAsync();
            return men;
        }
    }
}
