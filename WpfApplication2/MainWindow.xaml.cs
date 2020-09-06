using System;
using System.Collections.Generic;
using System.Linq;
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
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System.Net;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;

namespace WpfApplication2
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private string GET(string sURL)
        {
            WebRequest req = WebRequest.Create(sURL); 
            req.Proxy = WebProxy.GetDefaultProxy(); //На всякий случай задействуем настройки прокси по умолчанию
            WebResponse resp = req.GetResponse();
            Stream stream = resp.GetResponseStream();
            StreamReader sr = new StreamReader(stream);
            string Out = sr.ReadToEnd();
            sr.Close();
            return Out;
        }


        public void GetData()
        {
            
            string url = "https://jsonplaceholder.typicode.com/posts";
            int counter = 1;


            try
            {
                //Получаем json массив
                JArray json2 = JArray.Parse(GET(url));

                //Создаем книгу Excel
                Excel.Application excelapp = new Excel.Application();
                excelapp.Visible = true;                
                Excel.Workbook workbook = excelapp.Workbooks.Add();
                Excel.Worksheet worksheet = workbook.ActiveSheet;

                //Цикл по разбору массива на токены
                foreach (JToken element in json2)
                {
                    try
                    {
                        //Получаем информацию из токена
                        var title = element["title"].ToString();
                        var body = element["body"].ToString();

                        //Наполняем listbox
                        listBox.Items.Add(title);
                        listBox.Items.Add(body);

                        //наполняем эксель
                        worksheet.Cells[counter, 1] = title;
                        worksheet.Cells[counter + 1, 1] = body;
                        worksheet.Cells[counter + 2, 1] = "";

                        //здесь прибавляяем +3 чтобы в экселе появлялась пустая строка для разделения
                        counter = counter + 3;

                    }
                    catch (Exception ex2)
                    {
                        MessageBox.Show(ex2.Message);

                    }

                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);

                return;
            }

        }



        private void button_Click(object sender, RoutedEventArgs e)
        {
            GetData();
        }
    }
}
