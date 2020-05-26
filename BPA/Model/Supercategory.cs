using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BPA.Model
{
    /// <summary>
    /// Справочник суперкатегорий
    /// </summary>
    class Supercategory : TableBase
    {
        public override string TableName => "Суперкатегории";
        public override string SheetName => "Суперкатегории";

        public override IDictionary<string, string> Filds { get { return _filds; } }
        private readonly Dictionary<string, string> _filds = new Dictionary<string, string>
        {
            { "Id", "Код" },
            { "NameEn", "Суперкатегория (ENG)" },
            { "NameRu", "Суперкатегория (RUS)" }
        };

        /// <summary>
        /// Идентификатор
        /// </summary>
        public int Id { get; set; }
        /// <summary>
        /// Суперкатегория (ENG)
        /// </summary>
        public string NameEn { get; set; }
        /// <summary>
        /// Суперкатегория (RUS)
        /// </summary>
        public string NameRu { get; set; }

    }
}
