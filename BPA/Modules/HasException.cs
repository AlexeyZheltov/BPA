using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BPA.Modules
{
    public class HasExpection : Exception
    {
        public HasExpection(string message) : base(message) { }
    }

}
