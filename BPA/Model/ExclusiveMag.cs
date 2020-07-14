using BPA.Modules;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace BPA.Model
{
    class ExclusiveMag : TableBase
    {
        public override string TableName => "Exclusives";
        public override string SheetName => "Exclusives";
        public static Dictionary<string, int> ColDict { get; set; } = new Dictionary<string, int>();

        public override IDictionary<string, string> Filds
        {
            get
            {
                return _filds;
            }
        }
        private readonly Dictionary<string, string> _filds = new Dictionary<string, string>
        {
            ["Name"] = "Name"
        };

        public ExclusiveMag() { }
        public ExclusiveMag(Excel.ListRow row) => SetProperty(row);

        /// <summary>
        /// Имя экслюзива
        /// </summary>
        public string Name { get; set; }

        public static List<ExclusiveMag> GetAllExclusives()
        {
            List<ExclusiveMag> result = new List<ExclusiveMag>();

            foreach(Excel.ListRow row in new ExclusiveMag().Table.ListRows)
                result.Add(new ExclusiveMag(row));

            return result;
        }

        public class ContainsCompare : IEqualityComparer<ExclusiveMag>
        {
            public bool Equals(ExclusiveMag x, ExclusiveMag y)
            {
                return x.Name.ToLower() == y.Name.ToLower();
            }

            public int GetHashCode(ExclusiveMag obj)
            {
                return obj.Name.ToLower().GetHashCode();
            }
        }
    }
}
