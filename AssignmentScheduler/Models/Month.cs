using MongoDB.Bson;
using MongoDB.Bson.Serialization.Attributes;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AssignmentScheduler.Models
{
    public class Month
    {
        [BsonId]
        public ObjectId Id { get; set; }
        public string english {  get; set; }
        public string patwa {  get; set; }
    }
}
