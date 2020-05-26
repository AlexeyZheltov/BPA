using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BPA.Model
{
    class Product : TableBase
    {
        public override string TableName => "Товары";
        public override string SheetName => "Товары";

        public override IDictionary<string, string> Filds { get { return _filds; } }
        private readonly Dictionary<string, string> _filds = new Dictionary<string, string>
        {
            { "Id", "Код" },
            { "и т.д", "и т.д" }

        };
    }
}
