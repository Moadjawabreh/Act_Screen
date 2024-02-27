using System;
using System.Collections.Generic;
using System.IO.Ports;
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
using FirstMvvm.ViewModel;
using FirstMvvm.Model;


namespace FirstMvvm
{
	/// <summary>
	/// Interaction logic for MainWindow.xaml
	/// </summary>
	public partial class MainWindow : Window
	{

		public MainWindow()
		{
			InitializeComponent();
			DataContext = new ViewModelS(Recieve_RichBox);
			Closing += (s, e) => ((ViewModelS)DataContext).ClosePort(); // Close the port when the window is closed
			if (DataContext is ViewModelS viewModel)
			{
				viewModel.MessageSent += ViewModel_MessageSent;
			}
		}

		private void Window_Loaded(object sender, RoutedEventArgs e)
		{

		}

		private void Send_Click(object sender, RoutedEventArgs e)
		{
			if (DataContext is ViewModelS viewModel)
			{
				viewModel.SendMessage(Message_txt.Text);
			}

		}

		private void ViewModel_MessageSent(object sender, string e)
		{
			// Update the RichTextBox or other UI elements
			Recieve_RichBox.AppendText(e + "\n");
			Recieve_RichBox.ScrollToEnd();

		}




		private void Connect_btn_Click(object sender, RoutedEventArgs e)
		{
			try
			{
				if (DataContext is ViewModelS viewModel)
				{
					if (viewModel.IsPortConnected()) // Assuming you have a method to check if the port is connected in your ViewModel
					{
						viewModel.ClosePort();
					}
					else
					{
						viewModel.OpenPort(Convert.ToInt32(BaudRate_text.Text));
					}
				}
			}
			catch (Exception ex)
			{
				MessageBox.Show($"Error: {ex.Message}");

			}


		}

		private void Clear_btn_Click(object sender, RoutedEventArgs e)
		{
			Recieve_RichBox.Document.Blocks.Clear();
		}

		private async void CheckBox_Checked(object sender, RoutedEventArgs e)
		{
			try
			{

				if (DataContext is ViewModelS viewModel)
				{

					int delayMilliseconds = Convert.ToInt32(Delay_Input.Text);

					string messageText = Message_txt.Text; // Capture the text before entering the background thread


					while (Delay_chk.IsChecked == true)
					{
						await Task.Delay(delayMilliseconds / 2);
						// Use Dispatcher.Invoke to update UI from the UI thread
						Dispatcher.Invoke(() =>
						{
							viewModel.SendMessage(messageText);
						});

						await Task.Delay(delayMilliseconds);
					}
				}
			}
			catch (Exception ex)
			{
				MessageBox.Show($"Error: {ex.Message}");
			}
		}


		private void Delay_Input_TextChanged(object sender, TextChangedEventArgs e)
		{

		}

		private void Combox_ports_SelectionChanged(object sender, SelectionChangedEventArgs e)
		{
			if (DataContext is ViewModelS viewModel)
			{
				viewModel.SelectedChanges();

			}

		}
	}
}
