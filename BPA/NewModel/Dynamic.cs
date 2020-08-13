using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BPA.NewModel
{
    class Dynamic
    {
        public object Value { get; set; }

        public Dynamic(object value)
        {
            Value = value;
        }

        public static implicit operator int(Dynamic value)
        {
            if (value == null) return default;
            try
            {
                return Convert.ToInt32(value.Value);
            }
            catch
            {
                return default;
            }
        }

        public static implicit operator double(Dynamic value)
        {
            if (value == null) return default;
            try
            {
                return Convert.ToDouble(value.Value);
            }
            catch
            {
                return default;
            }
        }

        public static implicit operator string(Dynamic value)
        {
            if (value == null) return default;
            try
            {
                return Convert.ToString(value.Value);
            }
            catch
            {
                return default;
            }
        }

        public static implicit operator DateTime(Dynamic value)
        {
            if (value == null) return default;
            try
            {
                return Convert.ToDateTime(value.Value);
            }
            catch
            {
                return default;
            }
        }

        public static implicit operator Dynamic(int value) => new Dynamic(value);
        public static implicit operator Dynamic(double value) => new Dynamic(value);
        public static implicit operator Dynamic(string value) => new Dynamic(value);
        public static implicit operator Dynamic(DateTime value) => new Dynamic(value);
    }
}
