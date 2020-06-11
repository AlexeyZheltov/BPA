using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BPA.Modules
{
    public class Parsing
    {
        public string CalculateStringFormula(string formula)
        {
            string tmpFormula = formula;
            tmpFormula = ChangePercent(tmpFormula);

            int bracketOpenPos = formula.IndexOf("(");
            int bracketClosePos = formula.LastIndexOf(")");
            if (bracketOpenPos >= 0)
            {
                string bracketFormula = tmpFormula.Substring(bracketOpenPos + 1, bracketClosePos - bracketOpenPos-1);
                tmpFormula = CalculateStringFormula(bracketFormula);
            }
            else if (Double.TryParse(tmpFormula, out _))
            {
                return tmpFormula;
            }

            DoMultiplication(tmpFormula);
            DoPlus(tmpFormula);

            return tmpFormula;
        }

        private string ChangePercent(string tmpFormula)
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

        private string DoMultiplication(string formula)
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

                if (signMultPos < signDivPos || signDivPos < 0)
                {
                    tmpFormula = GetMult(tmpFormula, signMultPos);
                } else if (signDivPos < signMultPos || signMultPos < 0)
                {
                    tmpFormula = GetDiv(tmpFormula, signDivPos);
                }
            } 
            while (signMultPos >= 0 && signDivPos >= 0);

            return tmpFormula;
        }

        private string DoPlus(string formula)
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

        private string GetMult(string tmpFormula, int singPos)
        {
            string numLeft = GetNumLeft(tmpFormula, singPos);
            string numRight = GetNumRight(tmpFormula, singPos);

            double result = double.Parse(numRight) * double.Parse(numLeft);

            return tmpFormula.Substring(0, singPos - numLeft.Length) +
                                result.ToString() +
                                tmpFormula.Substring(singPos + numRight.Length+1);
        }


        private string GetDiv(string tmpFormula, int singPos)
        {
            string numLeft = GetNumLeft(tmpFormula, singPos);
            string numRight = GetNumRight(tmpFormula, singPos);

            double result = double.Parse(numRight) / double.Parse(numLeft);

            return tmpFormula.Substring(0, singPos - numLeft.Length) +
                                result.ToString() +
                                tmpFormula.Substring(singPos + numRight.Length+1);

        }

        private string GetPlus(string tmpFormula, int singPos)
        {
            string numLeft = GetNumLeft(tmpFormula, singPos);
            string numRight = GetNumRight(tmpFormula, singPos);

            if (tmpFormula.Substring(singPos - numLeft.Length - 1, 1) == "-")
            {
                numLeft = "-" + numLeft;
            }

            double result = double.Parse(numRight) + double.Parse(numLeft);

            return tmpFormula.Substring(0, singPos - numLeft.Length) +
                                result.ToString() +
                                tmpFormula.Substring(singPos + numRight.Length+1);

        }

        private string GetMinus(string tmpFormula, int singPos)
        {
            string numLeft = GetNumLeft(tmpFormula, singPos);
            string numRight = GetNumRight(tmpFormula, singPos);

            if (tmpFormula.Substring(singPos - numLeft.Length, 1) == "-" || singPos == 0) {
                numLeft = "-" + numLeft;
            }

            double result = double.Parse(numLeft) - double.Parse(numRight);

            return tmpFormula.Substring(0, singPos - numLeft.Length) +
                                result.ToString() +
                                tmpFormula.Substring(singPos + numRight.Length+1);

        }

        private string GetNumLeft(string formula, int signPos)
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
            return numString;
        }

        private string GetNumRight(string formula, int signPos)
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
