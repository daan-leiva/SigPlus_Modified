using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SigPlus_Modified
{
   struct DataRow
   {
      public string name { get;  set; }
      public string companyName { get; set; }
      public DateTime timeStamp { get; set; }
      public bool citizenOrResident { get; set; }
      public string visitee { get; set; }
      public string otherPerson { get; set; }
      public string chaperone { get; set; }
   }
}