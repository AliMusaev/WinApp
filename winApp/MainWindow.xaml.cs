using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using winApp.Resources;

namespace winApp
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        Categories excelLoadFile;
        SubCategories subCategories;
        Calculate calculate;
        DocumentInfo documentInfo;
        
        public MainWindow()
        {
            InitializeComponent();
            subCategories = new SubCategories();
            excelLoadFile = new Categories();
            calculate = new Calculate();
            documentInfo = new DocumentInfo();
            currencyName.ItemsSource = documentInfo.CurrencyNameAndCode.Keys;
        }


        private void categoryList_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            List<string> names = new List<string>();
            foreach (var item in subCategories.LoadSubCategories(excelLoadFile.Book, categoryList.SelectedItem.ToString()))
            {
                names.Add(item.Key);
            }
            subCategoryList.ItemsSource = names;
        }

        private void subCategoryList_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            
        }
        private void calculateButton_Click(object sender, RoutedEventArgs e)
        {
            foreach (var item in subCategories.Data)
            {
                if (item.Key == subCategoryList.SelectedItem.ToString())
                {
                    calculate.SubName = item.Key;
                    calculate.StartCalculating(item.Value, Math.Round((Convert.ToDouble(costField.Text)),2));
                }
            }
        }

        // Load general categories from excel file 
        private void Window_Loaded(object sender, RoutedEventArgs e)
        {

            categoryList.ItemsSource = excelLoadFile.LoadCategories();
        }

        private void Window_Closed(object sender, EventArgs e)
        {
            excelLoadFile.Close();
        }

        private void GetDocInfo()
        {
            documentInfo.VendorName = vendorName.Text;
            documentInfo.VendorAdress = vendorAdress.Text;
            documentInfo.VendorITN = Convert.ToInt16(vendorITN.Text);
            documentInfo.VendorRRC = Convert.ToInt16(vendorRRC.Text);
            documentInfo.ShipperNameAndAdress = shipperNameAndAdress.Text;
            documentInfo.ConsigneeNameAndAdress = consigneeNameAndAdress.Text;
            documentInfo.DocNumber = docNumber.Text;
            documentInfo.DocData = docData.Text;
            documentInfo.CustomerName = customerName.Text;
            documentInfo.CustomerAdress = customerAdress.Text;
            documentInfo.CustomerITN = Convert.ToInt16(customerITN.Text);
            documentInfo.CustomerRRC = Convert.ToInt16(customerRRC.Text);
            documentInfo.CurrencyName = currencyName.Text;
            documentInfo.CurrencyCode = Convert.ToInt16(currencyCode.Text);
            documentInfo.GovermentID = govermentId.Text;
        }

        private void currencyName_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            currencyCode.Text = (documentInfo.CurrencyNameAndCode[currencyName.SelectedItem.ToString()]).ToString();
        }
    }
}
