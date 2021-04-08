using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Version3
{
    public class ListItem1
    {
        private string id = string.Empty;
        private string name = string.Empty;
        public ListItem1(string sid, string sname)
        {
            id = sid;
            name = sname;
        }

        public override string ToString()
        {
            return this.name;
        }

        public string ID
        {
            get
            {
                return this.id;
            }
            set
            {
                this.id = value;
            }
        }

        public string Name
        {
            get
            {
                return this.name;
            }
            set
            {
                this.name = value;
            }
        }
    }
}
