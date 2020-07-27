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
        public MainWindow()
        {
            InitializeComponent();
            subCategories = new SubCategories();
            excelLoadFile = new Categories();
            currencyName.ItemsSource = DocumentInfo.CurrencyNameAndCode.Keys;
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
            LoadingWindow loadingWindow = new LoadingWindow();
            MessageWindow messageWindow = new MessageWindow();
            messageWindow.Owner = this;
            loadingWindow.Owner = this;
            try
            {
                GetDocInfo();
                foreach (var item in subCategories.Data)
                {
                    if (item.Key == subCategoryList.SelectedItem.ToString())
                    {
                        
                        Calculate calculate = new Calculate(item.Value);
                        calculate.SubName = item.Key;
                        try
                        {
                            calculate.StartCalculating(item.Value, Math.Round((Convert.ToDouble(costField.Text)), 2), messageWindow, loadingWindow);

                        }
                        catch (Exception)
                        {
                            messageWindow.ShowMessage("Не введена сумма");
                        }
                    }
                }

            }
            catch (Exception)
            {
                messageWindow.ShowMessage("Не выбран радел");
            }
           
        }

        // Load general categories from excel file 
        private void Window_Loaded(object sender, RoutedEventArgs e)
        {

            categoryList.ItemsSource = excelLoadFile.LoadCategories();
        }

        private void Window_Closed(object sender, EventArgs e)
        {
            //excelLoadFile.Close();
        }

        private void GetDocInfo()
        {
            DocumentInfo.ChekID = checkID.Text;
            DocumentInfo.ChekData = checkData.Text;
            DocumentInfo.ChangeID = changeID.Text;
            DocumentInfo.ChangeData = changeData.Text;
            DocumentInfo.VendorName = vendorName.Text;
            DocumentInfo.VendorAdress = vendorAdress.Text;
            DocumentInfo.VendorITN = vendorITN.Text;
            DocumentInfo.VendorRRC = vendorRRC.Text;
            DocumentInfo.ShipperNameAndAdress = shipperNameAndAdress.Text;
            DocumentInfo.ConsigneeNameAndAdress = consigneeNameAndAdress.Text;
            DocumentInfo.DocNumber = docNumber.Text;
            DocumentInfo.DocData = docData.Text;
            DocumentInfo.CustomerName = customerName.Text;
            DocumentInfo.CustomerAdress = customerAdress.Text;
            DocumentInfo.CustomerITN = customerITN.Text;
            DocumentInfo.CustomerRRC = customerRRC.Text;
            DocumentInfo.CurrencyName = currencyName.Text;
            DocumentInfo.CurrencyCode = currencyCode.Text;
            DocumentInfo.GovermentID = govermentId.Text;
        }

        private void currencyName_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            currencyCode.Text = (DocumentInfo.CurrencyNameAndCode[currencyName.SelectedItem.ToString()]).ToString();
        }


        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            excelLoadFile.Close();

        }

        private void costField_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            if (!char.IsDigit(e.Text, 0) && (e.Text != ","))
            {
                e.Handled = true; 
            }
        }

        private void costField_GotFocus(object sender, RoutedEventArgs e)
        {
            costField.Text = "";
        }
    }
}
