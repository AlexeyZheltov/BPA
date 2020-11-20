using System; 
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BPA.Modules
{
    public static class ErrorMessageClass
    {
        public static string GetExceptionInfo(Exception ex)
        {
            StringBuilder builder = new StringBuilder();
            builder.Append($"Application: {ex.Source}\n");
            builder.Append($"Error description: {ex.Message}\n");
            builder.Append($"Calling method: {ex.TargetSite}\n");
            builder.Append($"Call stack: {ex.StackTrace}\n");
            
            return builder.ToString();
        }
    }
}
