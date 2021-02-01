using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FutureClient.Models
{
    public class PosDetail
    {
        public string queue_name { get; set; }

        /// <summary>
        /// Desc:
        /// Default:
        /// Nullable:True
        /// </summary>           
        public string code { get; set; }

        public double netPosition { get; set; }
    }
}
