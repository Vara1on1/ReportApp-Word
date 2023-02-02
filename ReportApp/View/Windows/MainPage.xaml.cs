using ReportApp.Model;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Forms.DataVisualization.Charting;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Excel = Microsoft.Office.Interop.Excel;

namespace ReportApp.View.Windows
{
    /// <summary>
    /// Логика взаимодействия для MainPage.xaml
    /// </summary>
    public partial class MainPage : Page
    {
        Excel.Application application = new Excel.Application();
        Core db = new Core();
        Users user = new Users();
        List<Payment> arrayPayment = new List<Payment>();
        List<Category> arrayCategory = new List<Category>();
        public MainPage()
        {
            InitializeComponent();
            user = db.context.Users.Where((userDb) => userDb.id_user == App.UserId).FirstOrDefault();
            arrayCategory = db.context.Category.ToList();

            ChartPayments.ChartAreas.Add(new ChartArea("Main"));
            var currentSeries = new Series("Payments")
            {
                IsValueShownAsLabel = true
            };
            ChartPayments.Series.Add(currentSeries);

            ComboUsers.ItemsSource = db.context.Users.ToList();
                 ComboChartTypes.ItemsSource = Enum.GetValues(typeof(SeriesChartType));
        }

        private void Report_Click(object sender, RoutedEventArgs e)
        {

            Excel.Workbook workbook = application.Workbooks.Add();
            Excel.Worksheet worksheet = workbook.Sheets.Add();

                worksheet.Name = user.last_name + " " + user.first_name + " " + user.patronymic_name;
                worksheet.Cells[1][1] = "Дата платежа";
                worksheet.Cells[2][1] = "Название";
                worksheet.Cells[3][1] = "Стоимость";
                worksheet.Cells[4][1] = "Кол-во";
                worksheet.Cells[5][1] = "Сумма";
                int currentRow = 2;

                foreach (Category category in arrayCategory)
                {
                    List<Payment> payments = db.context.Payment.Where(x => x.category_id == category.id_category && x.user_id == user.id_user).ToList();
                    if (payments.Count > 0)
                    {
                        Excel.Range rng = worksheet.get_Range($"A{currentRow}:E{currentRow}");
                        rng.Merge();
                        rng.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                        worksheet.Cells[1][currentRow] = category.name_category;
                        currentRow++;
                        foreach (Payment payment in payments)
                        {
                            worksheet.Cells[1][currentRow] = payment.date_payment.ToString();
                            worksheet.Cells[2][currentRow] = payment.name;
                            worksheet.Cells[3][currentRow] = payment.cost;
                            worksheet.Cells[4][currentRow] = payment.count;
                            worksheet.Cells[5][currentRow] = payment.cost * payment.count;
                            currentRow++;
                        }
                        rng = worksheet.get_Range($"A{currentRow}:D{currentRow}");
                        rng.Merge();
                        rng.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;
                        worksheet.Cells[1][currentRow] = "ИТОГО:";
                        worksheet.Cells[5][currentRow].Formula = $"=SUM(E{currentRow - payments.Count}:E{currentRow - 1})";
                        currentRow++;
                    }
                    else
                    {
                        continue;
                    }

                }
                worksheet.Columns.AutoFit();
                Excel.Range borderRange = worksheet.get_Range($"A{1}:E{currentRow - 1}");
                borderRange.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                borderRange.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                borderRange.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                borderRange.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
                borderRange.Borders[Excel.XlBordersIndex.xlInsideHorizontal].LineStyle = Excel.XlLineStyle.xlContinuous;
                borderRange.Borders[Excel.XlBordersIndex.xlInsideVertical].LineStyle = Excel.XlLineStyle.xlContinuous;
            
            application.Visible = true;
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {

        }

        private void UpdateChart(object sender, SelectionChangedEventArgs e)
        {
           
        }

        private void ComboUsers_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (ComboUsers.SelectedItem is Users currentUser &&
               ComboChartTypes.SelectedItem is SeriesChartType currentType)
            {
                Series currentSeries = ChartPayments.Series.FirstOrDefault();
                currentSeries.ChartType = currentType;
                currentSeries.Points.Clear();

                var categoriesList = db.context.Category.ToList();
                foreach (var category in categoriesList)
                {
                    currentSeries.Points.AddXY(category.name_category, db.context.Payment.ToList().Where(p => p.Users == currentUser && p.Category == category).Sum(p => p.price * p.cost));
                }
            }
        }
    }
}
