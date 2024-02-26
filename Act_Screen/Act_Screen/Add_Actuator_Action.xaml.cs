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
using System.Windows.Shapes;

namespace Act_Screen
{
    /// <summary>
    /// Interaction logic for Add_Actuator_Action.xaml
    /// </summary>
    public partial class Add_Actuator_Action : Window
    {
		public List<byte> SelectedActions { get; private set; } = new List<byte>();

		public Add_Actuator_Action()
		{
			InitializeComponent();
		}

		private void OKButton_Click(object sender, RoutedEventArgs e)
		{
			if (upDownCloseRadioButton.IsChecked == true)
			{
				SelectedActions.Add(0x00);
				SelectedActions.Add(0x01);
				SelectedActions.Add(0x10);
			}
			else if (leftRightCloseRadioButton.IsChecked == true)
			{
				SelectedActions.Add(0x02);
				SelectedActions.Add(0x03);
				SelectedActions.Add(0x10);
			}
			else if (unlockLockCloseRadioButton.IsChecked == true)
			{
				SelectedActions.Add(0x04);
				SelectedActions.Add(0x05);
				SelectedActions.Add(0x10);
			}
			else
			{
				MessageBox.Show("please select a port!");
				return;
			}

			DialogResult = true;

		}

		private void CancelButton_Click(object sender, RoutedEventArgs e)
		{
			// Close the current window
			this.Close();
		}

	}
}
