using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace BPA.Modules
{
    class StringCalculator
    {
        public delegate double StringCalculatorFunc(params string[] values);
        Dictionary<string, StringCalculatorFunc> expressions = new Dictionary<string, StringCalculatorFunc>();

        public double? Calculate(string key, params string[] values) => expressions.ContainsKey(key) ? (double?)expressions[key](values) : null;

        //public bool Add(string Key, string expression)
        //{
        //    expression = ClearExpression(expression);
        //    LinkedList<string> tokenList = ToLinkedList(expression);

        //    Stack<object> stack = new Stack<object>();

        //}

        private string ClearExpression(string expression)
        {
            char[] allowedChars = "v0123456789.,+-*/%".ToArray();
            return expression.Where(x => allowedChars.Contains(x))
                        .ToString()
                        .Replace(',', '.');
        }

        private LinkedList<string> ToLinkedList(string expression)
        {
            LinkedList<string> list = new LinkedList<string>();
            StringBuilder builder = new StringBuilder();
            bool numBuilding = false;
            char[] allowedChars = "v+-*/%".ToCharArray();
            char[] numeric = "0123456789.".ToCharArray();
            
            foreach(char ch in expression)
            {
                if(allowedChars.Contains(ch) && numBuilding)
                {
                    numBuilding = false;
                    list.AddLast(new LinkedListNode<string>(builder.ToString()));
                    builder.Clear();
                    list.AddLast(new LinkedListNode<string>(ch.ToString()));
                }
                else if (numeric.Contains(ch))
                {
                    numBuilding = true;
                    builder.Append(ch);
                }
                else
                {
                    list.AddLast(new LinkedListNode<string>(ch.ToString()));
                }
            }

            return list;
        }

        //private string TransformPercent(string expression)
        //{

        //}
        //Отсеять лишнии символы
        //Получить список токенов
        //преобразовать к обратной польской записи заодно отбрасывать невенрые токены
        //стек токенов для вычисления, один из токенов это параметр, то есть то, что будет передваться в делегат
        //в делегате как параметр должен выступать parapms
        //так же нужен словарь вычислителей-делегатов
        
    }
}
