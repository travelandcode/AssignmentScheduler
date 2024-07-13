using MongoDB.Bson;
using MongoDB.Bson.Serialization.Attributes;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AssignmentScheduler.Models
{
    public class Men
    {
        [BsonId]
        public ObjectId Id { get; set; }
        public string FullName { get; set; }

        public ObjectId ProfileId {  get; set; }
    }
}
