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
using ReportApp.Model;
using ReportApp.View;
using ReportApp.View.Windows;

namespace ReportApp.View.Pages
{
    /// <summary>
    /// Логика взаимодействия для AuthoregPage.xaml
    /// </summary>
    public partial class AuthoregPage : Page
    {
        Core db = new Core();
        List<Users> arrayUsers = new List<Users>();
        public AuthoregPage()
        {
            InitializeComponent();

            arrayUsers = db.context.Users.ToList();

            foreach (var user in arrayUsers)
            {
                LoginComboBox.Items.Add(user.login);
            }
        }

        private void SignInbuttonClick(object sender, RoutedEventArgs e)
        {
            if (LoginComboBox.SelectedValue == null)
            {
                MessageBox.Show("Выберите пользователя.");
                return;
            } 

            var login = LoginComboBox.SelectedValue.ToString();
            var user = db.context.Users.Where((userDb) => userDb.login == login).FirstOrDefault();
            if (user.id_user == 0) {
                MessageBox.Show("Непредвиденная ошибка.");
                return;
            }

            if (user.password == PasswordBox.Password)
            {
                // Авторизация успешна
                App.UserId = user.id_user;
                NavigationService.Navigate(new MainPage());
                return;
            }

            // Авторизация провалена
            MessageBox.Show("Неправильный логин или пароль.");
        }
        private void ExitButtonClick(object sender, RoutedEventArgs e)
        {
            if (NavigationService.CanGoBack)
                NavigationService.GoBack();
        }

        private void WordButton_Click(object sender, RoutedEventArgs e)
        {

        }
    }
}
