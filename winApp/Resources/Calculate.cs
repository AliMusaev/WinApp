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
        List<double> calculatedPrices;
        List<double[]> priceLists;
        int [] results;
        double summ;
        double cost;
        bool find = false;
        int mult = 1;


        public Calculate()
        {
            priceLists = new List<double[]>();
            calculatedPrices = new List<double>();
        }
        public void StartCalculating(List<Product> products, double cost1)
        {
            
            this.cost = cost1;
            
            if (products.Count == 1)
            {
                Output outPut = new Output();
                outPut.LoadCalculatedData(products, cost);
                outPut.OutputExit();
            }
            else
            {
                int a = RepCalc();
                if (cost - (100000 * a) % 2 == 0)
                    mult = 2;
                while (!find)
                {
                    CalculatePrice(products, cost - (100000 * a));
                    results = new int[priceLists.Count];
                    int[] multiplyCounter = new int[priceLists.Count];
                    BulkheadLists(0, 100000, multiplyCounter);
                    if (!find)
                        continue;
                    for (int j = 0; j < results.Length; j++)
                    {
                        results[j] *= a;
                    }
                    find = false;
                    while (!find)
                    {

                        multiplyCounter = new int[priceLists.Count];
                        BulkheadLists(0, cost - (100000 * a), multiplyCounter);
                        if (!find)
                            break;
                    }
                }
                
                Output outPut = new Output();
                outPut.LoadCalculatedData(results, products, calculatedPrices, cost);
                outPut.OutputExit();
            }

        }
        


        void BulkheadLists(int i, double cost, int[] multiplyCounter)
        {
            cost = Math.Round(cost, 2);
           if (priceLists[i].Length > multiplyCounter[i])
            {
                for (int k = priceLists[i].Length; k > 0 ; k--)
                {
                    SummCalculating(multiplyCounter);
                    if (summ == cost && !find)
                    {
                        for (int j = 0; j < results.Length; j++)
                        {
                            results[j] += multiplyCounter[j];
                            
                        }
                        find = true;
                        return;
                    }
                    else if(find)
                    {
                        return;
                    }
                    if (summ > cost)
                    {
                        break;
                    }
                    if (calculatedPrices[i] < cost)
                    {

                            if (multiplyCounter[i] + mult < priceLists[i].Length)
                            {
                                multiplyCounter[i] += mult;
                            }
                    }
                    if (i+1 < multiplyCounter.Length)
                    {
                        BulkheadLists(i + 1, cost, multiplyCounter);
                    }
                }
                multiplyCounter[i] = 0;
            }
            else
            {
                BulkheadLists(i + 1, cost, multiplyCounter);
            }
        }




        int RepCalc()
        {
            double input = cost;
            int counter = 0;
            while (input > 100000)
            {
                counter++;
                input -= 100000;
            }
            return counter;
        }


        private void SummCalculating(int [] multiplyCounter)
        {
            // Обнуление суммы при повторном проходе
            summ = 0;
            // Сложение стоимостей всех элементов 
            for (int i = 0; i < multiplyCounter.Length; i++)
            {
                if (priceLists[i].Length > multiplyCounter[i])
                {
                    summ += Math.Round((priceLists[i][multiplyCounter[i]]), 2);
                }
            }
        }

        private void CalculatePrice(List<Product> products, double cost)
        {

            // Инициализация при определение количество элементов в подклассе
            // Расчет стоимости каждого элемента подкласса в пределах ценового диапазона
            Random rand = new Random();
            calculatedPrices.Clear();
            //if (Convert.ToInt32(cost) == Convert.ToDouble(cost))
            //{
            //    for (int i = 0; i < products.Count; i++)
            //    {
            //        calculatedPrices.Add((double)rand.Next((Int32)products[i].minPrice, (Int32)products[i].maxPrice));
            //    }
            //}
            //else
            //{
                for (int i = 0; i < products.Count; i++)
                {
                    calculatedPrices.Add(Math.Round((rand.NextDouble() * (products[i].maxPrice - products[i].minPrice) + products[i].minPrice), 2));

                }
            //}
            
            // Обнуление счетчика повтора при перерасчете цен после неудачной итерации
            // Очистка списка расчитанных цен  после неудачной итерации
            priceLists.Clear();
            // Расчет таблицы цены - высчитывание множителя на кажду цену
            for (int i = 0; i < products.Count; i++)
            {
                int k = 1;
                List<double> priceMultiList = new List<double>();
                priceMultiList.Add(0);
                // цикл перемножений
                while (calculatedPrices[i] * k < cost)
                {
                    // Возможно надо добавить проверку на превышение стоимости
                    priceMultiList.Add((Math.Round(calculatedPrices[i] * k,2)));
                    k += 1;
                }
                // Запись готового списка цен продукта в общий список
                priceLists.Add(priceMultiList.ToArray());
                // создание ячеек счетчика согласно количеству продуктов в подкатегории
            }



        }
    }
}
