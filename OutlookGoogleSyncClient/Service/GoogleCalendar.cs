using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OutlookGoogleSyncClient.Service
{
    public class GoogleCalendar
    {
        public string Name { get; set; } = "Undefined";

        public override string ToString()
        {
            return Name;
        }
    }
}
