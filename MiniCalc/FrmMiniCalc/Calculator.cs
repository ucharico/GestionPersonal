using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using MiniCalc;

namespace MiniCalc
{
    public class ResultOutOfRangeException : ApplicationException
    {
    }
    public class NegativeParameterException : ApplicationException
    {
    }
    public enum CalcOperation
    {
        eAdd = 0,
        eSubtract = 1,
    }
    public class Calculator
    {

        public int Calculate(int a, CalcOperation op, int b)
        {
            int nResult = 0;
            if (CalcOperation.eAdd == op)
            {
                nResult = Add(a, b);
            }
            else if (CalcOperation.eSubtract == op)
            {
                nResult = Subtract(b, a);
            }
            return nResult;
        }

        public int Subtract(int numberToSubtract, int subtractFrom)
        {
            int result;
            if (subtractFrom < 0 || numberToSubtract < 0)
            {
                NegativeParameterException npEx = new NegativeParameterException();
                throw npEx;
            }
            result = subtractFrom - numberToSubtract;
            if (result < 0)
            {
                ResultOutOfRangeException rangeEx = new ResultOutOfRangeException();
                throw rangeEx;
            }
            return result;
            
            //CheckForNegativeNumbers(a, b);
        }
        public int Add(int a, int b)
        {
            /*
            int result;
            if (a < 0 || b < 0)
            {
                NegativeParameterException npEx = new NegativeParameterException();
                throw npEx;
            }
            result = a + b;
            if (result < 0)
            {
                ResultOutOfRangeException rangeEx = new ResultOutOfRangeException();
                throw rangeEx;
            }
            return result;
            */
            int result;
            if (a < 0 || b < 0)
            {
                NegativeParameterException npEx = new NegativeParameterException();
                throw npEx;
            }
            result = a + b;
            if (result < 0)
            {
                ResultOutOfRangeException rangeEx = new ResultOutOfRangeException();
                throw rangeEx;
            }
            return result;
        }

        public void CheckForNegativeNumbers(int a, int b)
        {
            if (a < 0 || b < 0)
            {
                NegativeParameterException npEx = new NegativeParameterException();
                throw npEx;
            }
        }
    }
}
