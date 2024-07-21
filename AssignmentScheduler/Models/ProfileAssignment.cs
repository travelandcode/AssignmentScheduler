using MongoDB.Bson;
using MongoDB.Bson.Serialization.Attributes;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AssignmentScheduler.Models
{
    public class ProfileAssignment
    {
        [BsonId]
        public ObjectId Id { get; set; }
        public ObjectId profileid { get; set; }
        public ObjectId assignmentid { get; set; }
    }
}
