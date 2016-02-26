using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Drawing;

namespace SigPlus_Modified
{
   class viewRow
   {
      public DateTime timeStamp { get; set; }
      public Image name { get; set; }
      public Image company { get; set; }
      public bool citizenOrResident { get; set; }
      public string visitee { get; set; }
      public Image otherPerson { get; set; }
      public string chaperone { get; set; }

      public viewRow(DateTime _timeStamp, Image _name, Image _company, bool _citizenOrResident, string _visitee, Image _otherPerson, string _chaperone)
      {
         timeStamp = _timeStamp;
         name = _name;
         company = _company;
         citizenOrResident = _citizenOrResident;
         visitee = _visitee;
         otherPerson = _otherPerson;
         chaperone = _chaperone;
      }
   }
}
