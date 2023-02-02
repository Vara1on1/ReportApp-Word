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
using ReportApp.View.Windows;
using ReportApp.Model;
using Word = Microsoft.Office.Interop.Word;


namespace ReportApp.View.Pages
{
    /// <summary>
    /// Логика взаимодействия для DataWordPage.xaml
    /// </summary>
    public partial class DataWordPage : Page
    {
        Core db = new Core();
        List<Users> arrayUsers = new List<Users>();
        public DataWordPage()
        {
            InitializeComponent();
        }

        private void WordButton_Click(object sender, RoutedEventArgs e)
        {
            var allUsers = db.context.Users.ToList();
            var allCategories = db.context.Category.ToList();

            var application = new Word.Application();

            Word.Document document = application.Documents.Add();

            foreach (var user in allUsers)
            {
                Word.Paragraph userParagraph = document.Paragraphs.Add();
                Word.Range userRange = userParagraph.Range;
                userRange.Text = user.first_name;
                userParagraph.set_Style("Заголовок 1");
                userRange.InsertParagraphAfter();

                Word.Paragraph tableParagraph = document.Paragraphs.Add();
                Word.Range tableRange = tableParagraph.Range;
                Word.Table paymentsTable = document.Tables.Add(tableRange, allCategories.Count() + 1, 3);
                paymentsTable.Borders.InsideLineStyle = paymentsTable.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;
                paymentsTable.Range.Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;

                Word.Range cellRange;

                cellRange = paymentsTable.Cell(1, 1).Range;
                cellRange.Text = "Иконка";
                cellRange = paymentsTable.Cell(1, 2).Range;
                cellRange.Text = "Категория";
                cellRange = paymentsTable.Cell(1, 3).Range;
                cellRange.Text = "Сумма расходов";

                paymentsTable.Rows[1].Range.Bold = 1;
                paymentsTable.Rows[1].Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;

                for (int i = 0; i < allCategories.Count; i++)
                {
                    var currentCatgory = allCategories[i];

                    cellRange = paymentsTable.Cell(i + 2, 1).Range;
                    Word.InlineShape imageShape = cellRange.InlineShapes.AddPicture(AppDomain.CurrentDomain.BaseDirectory + "..\\..\\" + currentCatgory);
                    imageShape.Width = imageShape.Height = 40;
                    cellRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;

                    cellRange = paymentsTable.Cell(i + 2, 2).Range;
                    cellRange.Text = currentCatgory.name_category;

                    cellRange = paymentsTable.Cell(i + 2, 3).Range;
                    cellRange.Text = user.Payment.ToList().Where(p => p.Category == currentCatgory).Sum(p => p.count * p.price).ToString() + " руб.";
                }

                Payment maxPayment = user.Payment.OrderByDescending(p => p.price * p.count).FirstOrDefault();
                if (maxPayment != null)
                {
                    Word.Paragraph maxPaymentParagraph = document.Paragraphs.Add();
                    Word.Range maxPaymentRange = maxPaymentParagraph.Range;
                    maxPaymentRange.Text = $"Самый дорогостоящий платеж - {maxPayment.name} за {(maxPayment.price * maxPayment.count).ToString()} руб. от " + maxPayment.date_payment.ToString();
                }
            }
        }
    }
}
