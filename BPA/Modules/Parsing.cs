using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BPA.Modules
{
    public class Parsing
    {
        private string ExcelFormula;
        public double Result;
        public Parsing() { }

        public Parsing(string formula)
        {
            ExcelFormula = formula;
            Result = Double.Parse(CalculateStringFormula());
        }
        public struct Bracket
        {
            public int OpenPos
            {
                get; set;
            }
            public int ClosePos
            {
                get; set;
            }
            public string Formula
            {
                get; set;
            }
            public string Result
            {
                get {
                    string result = ChangePercent(this.Formula);

                    if (Double.TryParse(this.Formula, out _))
                        return Result = result;

                    result = DoMultiplication(result);
                    result = DoPlus(result);
                    return Result = result;
                }
                set { }
            }
        }
        public string CalculateStringFormula()
        {
            string tmpFormula = ExcelFormula;

            tmpFormula = tmpFormula.Replace(" ", "");
            if (tmpFormula.Substring(0, 1) == "=")
                tmpFormula = tmpFormula.Substring(1, tmpFormula.Length - 1);

            int bracketOpenPos;
            do
            {
                Bracket bracket = new Bracket();
                bracketOpenPos = tmpFormula.IndexOf("(");
                if (bracketOpenPos < 0) break;

                bracket = GetResultInBrackets(tmpFormula);
                tmpFormula = tmpFormula.Substring(0, bracket.OpenPos) +
                                    bracket.Result +
                                    tmpFormula.Substring(bracket.ClosePos+1);
            }
            while (bracketOpenPos >= 0);

            Bracket bracketRes = new Bracket();
            bracketRes.Formula = tmpFormula;
            return bracketRes.Result;
        }

        private Bracket GetResultInBrackets(string formula)
        {
            string tmpFormula = formula;

            Bracket bracket = new Bracket();

            int bracketClosePos = tmpFormula.IndexOf(")");
            if (bracketClosePos < 0)
            {
                bracket.Formula = tmpFormula;
            }
            else
            {
                int bracketOpenPos = tmpFormula.LastIndexOf("(", bracketClosePos);

                bracket.OpenPos = bracketOpenPos;
                bracket.ClosePos = bracketClosePos;
                bracket.Formula = tmpFormula.Substring(bracketOpenPos + 1, bracketClosePos - bracketOpenPos - 1);
            }
            return bracket;
        }

        private static string ChangePercent(string tmpFormula)
        {
            int percentPos;
            do
            {
                percentPos = tmpFormula.IndexOf("%");
                if (percentPos < 0)
                    break;

                string numString = GetNumLeft(tmpFormula, percentPos);
                int c = percentPos - numString.Length;

                if (double.TryParse(numString, out double num))
                {
                    double numRatio = num / 100;

                    tmpFormula = string.Concat(tmpFormula.Substring(0, c), numRatio.ToString(), tmpFormula.Substring(percentPos + 1));
                }
            } while (percentPos >= 0);

            return tmpFormula;
        }

        private static string DoMultiplication(string formula)
        {
            int signMultPos;
            int signDivPos;
            string tmpFormula = formula;

            do
            {
                signMultPos = tmpFormula.IndexOf("*");
                signDivPos = tmpFormula.IndexOf("/");

                if (signDivPos < 0 && signMultPos < 0)
                    break;

                if ((signMultPos < signDivPos || signDivPos < 0) && signMultPos >= 0)
                //if (signMultPos < signDivPos || signDivPos < 0)
                {
                    tmpFormula = GetMult(tmpFormula, signMultPos);
                } 
                //else if (signDivPos < signMultPos || signMultPos < 0)
                else if ((signDivPos < signMultPos || signMultPos < 0) && signDivPos >= 0)
                {
                    tmpFormula = GetDiv(tmpFormula, signDivPos);
                }
            } 
            while (signMultPos >= 0 && signDivPos >= 0);

            return tmpFormula;
        }

        private static string DoPlus(string formula)
        {
            int signPlusPos;
            int signMinusPos;
            string tmpFormula = formula;

            do
            {
                signPlusPos = tmpFormula.IndexOf("+");
                signMinusPos = tmpFormula.IndexOf("-");

                if (signPlusPos < 0 && signMinusPos < 0)
                    break;

                if ((signPlusPos < signMinusPos || signMinusPos < 0) && signPlusPos >= 0)
                {
                    tmpFormula = GetPlus(tmpFormula, signPlusPos);
                }
                else if ((signMinusPos < signPlusPos || signPlusPos < 0) && signMinusPos >=0)
                {
                    tmpFormula = GetMinus(tmpFormula, signMinusPos);
                }
            }
            while (signPlusPos >= 0 && signMinusPos >= 0);

            return tmpFormula;
        }

        private static string GetMult(string tmpFormula, int singPos)
        {
            string numLeft = GetNumLeft(tmpFormula, singPos);
            string numRight = GetNumRight(tmpFormula, singPos);

            double result = double.Parse(numRight) * double.Parse(numLeft);

            return tmpFormula.Substring(0, singPos - numLeft.Length) +
                                result.ToString() +
                                tmpFormula.Substring(singPos + numRight.Length+1);
        }

        private static string GetDiv(string tmpFormula, int singPos)
        {
            string numLeft = GetNumLeft(tmpFormula, singPos);
            string numRight = GetNumRight(tmpFormula, singPos);

            double result = double.Parse(numLeft) / double.Parse(numRight);

            return tmpFormula.Substring(0, singPos - numLeft.Length) +
                                result.ToString() +
                                tmpFormula.Substring(singPos + numRight.Length+1);

        }

        private static string GetPlus(string tmpFormula, int singPos)
        {
            string numLeft = GetNumLeft(tmpFormula, singPos);
            string numRight = GetNumRight(tmpFormula, singPos);

            double result = double.Parse(numRight) + double.Parse(numLeft);

            return tmpFormula.Substring(0, singPos - numLeft.Length) +
                                result.ToString() +
                                tmpFormula.Substring(singPos + numRight.Length+1);

        }

        private static string GetMinus(string tmpFormula, int singPos)
        {
            string numLeft = GetNumLeft(tmpFormula, singPos);
            string numRight = GetNumRight(tmpFormula, singPos);

            double result = double.Parse(numLeft) - double.Parse(numRight);

            return tmpFormula.Substring(0, singPos - numLeft.Length) +
                                result.ToString() +
                                tmpFormula.Substring(singPos + numRight.Length+1);

        }

        private static string GetNumLeft(string formula, int signPos)
        {
            string numString = "";
            if (signPos <= 0)
                return numString;
            
            for (int c = signPos - 1; c >= 0; c--)
            {
                string sign = formula.Substring(c, 1);

                if (!double.TryParse(sign, out _) && sign != "." && sign != ",")
                    break;

                numString = sign + numString;
            }

            if (signPos - numString.Length - 1 > 0)
            {
                if (numString.Substring(signPos - numString.Length - 1, 1) == "-")
                {
                    numString = "-" + numString;
                }
            }

            return numString;
        }

        private static string GetNumRight(string formula, int signPos)
        {
            string numString = "";
            if (signPos <= 0)
                return numString;

            for (int c = signPos + 1; c< formula.Length; c++)
            {
                string sign = formula.Substring(c, 1);

                if (!double.TryParse(sign, out _) && sign != "." && sign != ",")
                    break;

                numString = numString+sign;
            }
            return numString;
        }
    }
}
