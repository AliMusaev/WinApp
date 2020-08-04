using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace winApp.Resources
{
    public class Result
    {
        public int Amount;
        public double Price;

        public Result(int inputAmount, double inputPrice)
        {
            Amount = inputAmount;
            Price = inputPrice;
        }
    }
}
