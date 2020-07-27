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
        string subName;
        public string SubName { get => subName; set => subName = value; }
        int [] calculatedPrices;
        int[] calculatedAmount;
        double [] results;
        int cost;
        double cost1;
        int[] distributedValues;
        List<Product> products;
        public Calculate(List<Product> input)
        {
            products = input;
            calculatedPrices = new int[products.Count];
            calculatedAmount = new int[products.Count];
            distributedValues = new int[products.Count];
            
        }
        public void StartCalculating(List<Product> products, double cost1, MessageWindow message, LoadingWindow loading)
        {
            loading.Show();
            this.cost = (int)(cost1 * 100);
            if (products.Count == 1)
            {
                
                Output outPut = new Output();
                outPut.LoadCalculatedData(products, (double)cost/100);
                loading.Close();
                outPut.OutputExit();
                message.ShowMessage("Завершено успешно!");
            }
            else
            {
                if (FindDivisors())
                {
                    CalculateResult();
                    Output outPut = new Output();
                    outPut.LoadCalculatedData(calculatedAmount, products, results, cost1);
                    
                    outPut.OutputExit();
                    loading.Close();
                    message.ShowMessage("Завершено успешно!");

                }
                else
                {
                    loading.Close();
                    message.ShowMessage("Совпадений не найдено!");
    
                }
            }
            
        }
        


        




        void PercentageDistribution()
        {
                int percent = 10000;
                int tempCost = cost;
                for (int i = 0; i < calculatedPrices.Length; i++)
                {
                    if (percent > 0 && tempCost > 0)
                    {
                        if (i + 1 == calculatedPrices.Length)
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
        }        
        bool FindDivisors()
        {
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
                    PercentageDistribution();
                    k = 0;
                    counter++;
                }
             if (counter >= 20000)
                {
                    return false;
                }
            }
            return true;
        }


        void CalculateResult()
        {
            results = new double[distributedValues.Length];
            for (int i = 0; i < results.Length; i++)
            {
                results[i] = ((double)(distributedValues[i] / calculatedAmount[i]) / 100);
                cost1 = (double)cost / 100;
            }
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
    
