using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BPA.Modules
{
    public struct ArticleQuantity
    {   public string Article
        {
            get; set;
        }
        public int Month
        {
            get; set;
        }
        public double Quantity
        {
            get; set;
        }
        public string Campaign
        {
            get; set;
        }
        public double PriceList
        {
            get; set;
        }
        public double Bonus
        {
            get; set;
        }

        public int row_in_file
        {
            get; set;
        } 
    }
}
