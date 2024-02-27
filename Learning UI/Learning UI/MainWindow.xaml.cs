using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace Learning_UI
{
	/// <summary>
	/// Interaction logic for MainWindow.xaml
	/// </summary>
	public partial class MainWindow : Window
	{
		public MainWindow()
		{
			InitializeComponent();
		}

		private void Window_ImageFailed(object sender, ExceptionRoutedEventArgs e)
		{

		}

		private void Button_Click(object sender, RoutedEventArgs e)
		{
			int x;
			x = 10;
			if (x == 0)
			{
				return;
			}
		}

		private void EmailTextBox_GotFocus(object sender, RoutedEventArgs e)
		{
			if (EmailTextBox.Text == "Email")
			{
				EmailTextBox.Text = "";
				EmailTextBox.Foreground = Brushes.Black;
			}
		}

		private void EmailTextBox_LostFocus(object sender, RoutedEventArgs e)
		{
			if (string.IsNullOrWhiteSpace(EmailTextBox.Text))
			{
				EmailTextBox.Text = "Email";
				EmailTextBox.Foreground = Brushes.LightGray;
			}
		}

		private void PasswordTextBox_GotFocus(object sender, RoutedEventArgs e)
		{
			if (PasswordTextBox.Text == "Password")
			{
				PasswordTextBox.Text = "";
				PasswordTextBox.Foreground = Brushes.Black;
			}
		}

		private void PasswordTextBox_LostFocus(object sender, RoutedEventArgs e)
		{
			if (string.IsNullOrWhiteSpace(PasswordTextBox.Text))
			{
				PasswordTextBox.Text = "Password";
				PasswordTextBox.Foreground = Brushes.LightGray;
			}
		}


	}
}