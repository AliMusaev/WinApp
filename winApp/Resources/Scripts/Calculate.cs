using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using Excel = Microsoft.Office.Interop.Excel;

namespace winApp.Resources
{

    class Calculate
    {
        int cost;
        List<Result> result= new List<Result>();
        List<Product> products;
        public Calculate(List<Product> input)
        {
            products = input;
            
        }
        public List<Result> StartCalculating(List<Product> products, double inputCost, out int postCode)
        {
            this.cost = (int)(inputCost * 100);
            if (products.Count == 1)
            {
                result.Add(new Result(1, inputCost));
                postCode = 0;
                return result;
            }
            else
            {
                if (FindDivisors())
                {
                    postCode = 0;
                    return result;
                }
                else
                {
                    postCode = 1;
                    return null;
                }

            }
            
        }








        private int[] PercentageDistribution()
        {
            int[] distributedValues = new int[products.Count];
            int percent = 10000;
            int tempCost = cost;
            for (int i = 0; i < distributedValues.Length; i++)
            {
                if (percent > 0 && tempCost > 0)
                {
                    if (i + 1 == distributedValues.Length)
                    {
                        distributedValues[i] = tempCost;
                        break;
                    }
                    else
                    {
                        int temp = new Random().Next(0, percent);
                        distributedValues[i] = (cost / 10000) * temp;
                        percent -= temp;
                        tempCost -= distributedValues[i];
                    }
                }
                else
                {
                    break;
                }
            }
            return distributedValues;
        }        
        bool FindDivisors()
        {
            int[] calculatedAmount = new int[products.Count];
            int[] calculatedPrices = new int[products.Count];
            int[] distributedValues = new int[products.Count];
            int counter = 0;
            int k = 0;
            
            while(k < calculatedPrices.Length)
            {
                calculatedPrices[k] = CalculatePrice(products, k, distributedValues[k]);
                if (calculatedPrices[k] != 0)
                {
                    calculatedAmount[k] = distributedValues[k] / calculatedPrices[k];
                    k++;
                }
                else
                {
                    distributedValues = PercentageDistribution();
                    k = 0;
                    counter++;
                }
             if (counter >= 10000)
                {
                    return false;
                }
            } 

            for (int i = 0; i < products.Count; i++)
            {
                result.Add(new Result (calculatedAmount[i], (double)(distributedValues[i] / calculatedAmount[i]) / 100));
            }
            return true;
            
        }

        
        private int CalculatePrice(List<Product> products, int i, int value)
        {
            if (value == 0)
            {
                return 0;
            }
            List<int> arr = new List<int>();
            for (int counter = Convert.ToInt32(products[i].minPrice * 100); counter < Convert.ToInt32(products[i].maxPrice * 100); counter++)
            {
                if (value % counter == 0)
                {
                    arr.Add(counter);
                }
            }
            if(arr.Count > 0)
            {
                return  arr[new Random().Next(0, arr.Count)];
            }
            else
            {
                return 0;
            }
        }
    }
}
    
