using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MrBotAddIn
{
    public class conexionesOData
    {
        [Category("Connection")]
        [Description("Url of the OData server.")]

        public string Url
        {
            get;
            set;
        }

        [Category("Connection")]
        [Description("Name to identify the connection.")]

        public string Name
        {
            get;
            set;
        }
        
    }
}
