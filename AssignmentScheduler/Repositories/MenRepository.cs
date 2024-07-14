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
            var men = await _menCollection.Find(FilterDefinition<Men>.Empty).SortBy(m => m.FullName).ToListAsync();
            return men;
        }
    }
}
