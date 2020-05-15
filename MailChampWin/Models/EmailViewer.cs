using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MailChampWin.Models
{
    public class EmailViewer
    {
        public string From { get; set; }
        public DateTime On { get; set; }
        public string Subject { get; set; }
        public string Message { get; set; }
        public string To { get; set; }

    }
}
