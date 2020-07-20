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
        List<List<double>> priceLists;
        List<int> multiplyCounter;
        List<List<int>> results = new List<List<int>>();
        double summ;
        bool mult = false;

        
        public Calculate()
        {
            priceLists = new List<List<double>>();
            calculatedPrices = new List<double>();
            multiplyCounter = new List<int>();
        }
        public void StartCalculating(List<Product> products, double cost)
        {
            if(products.Count == 1)
            {
                Output outPut = new Output();
                outPut.LoadCalculatedData(products, cost);
                outPut.OutputExit();
            }
            else
            {
                bool find = false;
                while (!find)
                {
                    results.Clear();
                    CalculatePrice(products, cost);
                    BulkheadLists(cost, 0);
                    if (results.Count > 0)
                    {
                        find = true;
                    }
                }
                Output outPut = new Output();
                outPut.LoadCalculatedData(results, products, calculatedPrices, cost);
                outPut.OutputExit();
            }

        }
        


        void BulkheadLists(double cost, int i)
        {
            Random range = new Random();
            if (priceLists[i].Count > multiplyCounter[i])
            {
                for (int k = priceLists[i].Count; k > 0 ; k--)
                {
                    SummCalculating();
                    if ((cost - summ) > 20000)
                    {
                        mult = true;
                    }
                    if (summ == cost)
                    {
                        List<int> blab = new List<int>(multiplyCounter);
                        results.Add(blab);
                        break;
                    }
                    if (summ > cost)
                    {
                        break;
                    }

                    if (calculatedPrices[i] < cost)
                    {

                        if (mult && multiplyCounter[i] + 71 < priceLists[i].Count)
                            multiplyCounter[i] += 70;
                        else if (mult && multiplyCounter[i] + 61 < priceLists[i].Count)
                            multiplyCounter[i] += 60;
                        else if (mult && multiplyCounter[i] + 51 < priceLists[i].Count)
                            multiplyCounter[i] += 50;
                        else if (mult && multiplyCounter[i] + 31 < priceLists[i].Count)
                            multiplyCounter[i] += 30;
                        else if (mult && multiplyCounter[i] + 21 < priceLists[i].Count)
                            multiplyCounter[i] += 20;
                        else if (mult && multiplyCounter[i] + 11 < priceLists[i].Count)
                            multiplyCounter[i] += 10;
                        else if (multiplyCounter[i] + 1 < priceLists[i].Count)
                            multiplyCounter[i]++;
                    }
                    mult = false;
                    if (i+1 < multiplyCounter.Count)
                    {
                        BulkheadLists(cost, i + 1);
                    }
                    
                }
                multiplyCounter[i] = 0;
            }
            else
            {
                BulkheadLists(cost, i + 1);
            }
        }







        private void SummCalculating()
        {
            // Обнуление суммы при повторном проходе
            summ = 0;
            // Сложение стоимостей всех элементов 
            for (int i = 0; i < multiplyCounter.Count; i++)
            {
                if (priceLists[i].Count > multiplyCounter[i])
                {
                    summ += Math.Round(priceLists[i][multiplyCounter[i]], 2);
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
            multiplyCounter.Clear();
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
                    priceMultiList.Add(calculatedPrices[i] * k);
                    k += 1;
                }
                // Запись готового списка цен продукта в общий список
                priceLists.Add(priceMultiList);
                // создание ячеек счетчика согласно количеству продуктов в подкатегории
                multiplyCounter.Add(0);
            }



        }
    }
}
