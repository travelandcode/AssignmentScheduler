using MongoDB.Driver;
using AssignmentScheduler.Models;
using AssignmentScheduler.Interfaces;

namespace AssignmentScheduler.Repositories
{
    public class MonthRepository: IMonthRepository
    {
        private readonly IMongoCollection<Month> _monthCollection;

        public MonthRepository(IMongoDatabase db)
        {
            _monthCollection = db.GetCollection<Month>("Month");
        }

        public async Task<string> GetNextMonth()
        {
            DateTime currentDate = DateTime.Now;
            DateTime nextMonth = currentDate.AddMonths(1);
            String month = nextMonth.ToString("MMMM");
            var monthFromDB = await _monthCollection.Find(m => m.english == month).ToListAsync();
            var monthInPatwa = monthFromDB.Select(m => m.patwa).FirstOrDefault();

            return monthInPatwa;
        }


    }
}
