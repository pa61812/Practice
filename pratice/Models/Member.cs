using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace pratice.Models
{
    public class Member
    {
        public int ID { get; set; }
        public string Account { get; set; }
        public string Password { get; set; }
        public string Name { get; set; }
        public string Phone { get; set; }
        public string Tel { get; set; }
        public string Gender { get; set; }
        public DateTime? Birthday { get; set; }
    }
}