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
        List<int> calculatedPrices;
        List<int[]> priceLists;
        int [] results;
        double summ;
        int cost;
        bool find = false;
        int mult = 1;
        int[] distributedValues;
        List<Product> prod;
        public Calculate()
        {
            priceLists = new List<int[]>();
            calculatedPrices = new List<int>();
        }
        public void StartCalculating(List<Product> products, double cost1)
        {
            
            this.cost = (int)(cost1 * 100);
            if (products.Count == 1)
            {
                Output outPut = new Output();
                outPut.LoadCalculatedData(products, cost);
                outPut.OutputExit();
            }
            else
            {
                prod = products;
                PercentageDistribution();
                ValueFind();

                //Output outPut = new Output();
                //outPut.LoadCalculatedData(results, products, cost);
                //outPut.OutputExit();
            }

        }
        


        //void BulkheadLists(int i, double cost, int[] multiplyCounter)
        //{
        //    cost = Math.Round(cost, 2);
        //   if (priceLists[i].Length > multiplyCounter[i])
        //    {
        //        for (int k = priceLists[i].Length; k > 0 ; k--)
        //        {
        //            SummCalculating(multiplyCounter);
        //            if (summ == cost && !find)
        //            {
        //                for (int j = 0; j < results.Length; j++)
        //                {
        //                    results[j] += multiplyCounter[j];
                            
        //                }
        //                find = true;
        //                return;
        //            }
        //            else if(find)
        //            {
        //                return;
        //            }
        //            if (summ > cost)
        //            {
        //                break;
        //            }
        //            if (calculatedPrices[i] < cost)
        //            {

        //                    if (multiplyCounter[i] + mult < priceLists[i].Length)
        //                    {
        //                        multiplyCounter[i] += mult;
        //                    }
        //            }
        //            if (i+1 < multiplyCounter.Length)
        //            {
        //                BulkheadLists(i + 1, cost, multiplyCounter);
        //            }
        //        }
        //        multiplyCounter[i] = 0;
        //    }
        //    else
        //    {
        //        BulkheadLists(i + 1, cost, multiplyCounter);
        //    }
        //}




        void PercentageDistribution()
        {
            
                int percent = 10000;
                int tempCost = cost;
                CalculatePrice(prod);
                distributedValues = new int[calculatedPrices.Count];
                for (int i = 0; i < calculatedPrices.Count; i++)
                {
                    if (percent > 0 && tempCost > 0)
                    {
                        if (i + 1 == calculatedPrices.Count)
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
            

        private void ValueFind()
        {
            int k = 0;
            results = new int[calculatedPrices.Count];
            while (k < calculatedPrices.Count)
            {

                if (distributedValues[k] > 0)
                {
                    calculatedPrices[k] = CalculatePrice(prod, k, distributedValues[k]);
                    if (calculatedPrices[k] != 0)
                    {
                        k++;
                    }
                    else
                    {
                        PercentageDistribution();
                        k = 0;
                    }
                }
                else
                    k++;
            }
        }

        //private void SummCalculating(int [] multiplyCounter)
        //{
        //    // Обнуление суммы при повторном проходе
        //    summ = 0;
        //    // Сложение стоимостей всех элементов 
        //    for (int i = 0; i < multiplyCounter.Length; i++)
        //    {
        //        if (priceLists[i].Length > multiplyCounter[i])
        //        {
        //            summ += Math.Round((priceLists[i][multiplyCounter[i]]), 2);
        //        }
        //    }
        //}
        private int CalculatePrice(List<Product> products, int i, int value)
        {
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
        private void CalculatePrice(List<Product> products)
        {

            // Инициализация при определение количество элементов в подклассе
            // Расчет стоимости каждого элемента подкласса в пределах ценового диапазона
            Random rand = new Random();
            calculatedPrices.Clear();
            
                for (int i = 0; i < products.Count; i++)
                {
                
                    calculatedPrices.Add(0);
                }

        
            // Обнуление счетчика повтора при перерасчете цен после неудачной итерации
            // Очистка списка расчитанных цен  после неудачной итерации
            //priceLists.Clear();
            //// Расчет таблицы цены - высчитывание множителя на кажду цену
            //for (int i = 0; i < products.Count; i++)
            //{
            //    int k = 1;
            //    List<double> priceMultiList = new List<double>();
            //    priceMultiList.Add(0);
            //    // цикл перемножений
            //    while (calculatedPrices[i] * k < cost)
            //    {
            //        // Возможно надо добавить проверку на превышение стоимости
            //        priceMultiList.Add((Math.Round(calculatedPrices[i] * k,2)));
            //        k += 1;
            //    }
            //    // Запись готового списка цен продукта в общий список
            //    priceLists.Add(priceMultiList.ToArray());
                
                // создание ячеек счетчика согласно количеству продуктов в подкатегории
            }


        }
    }

