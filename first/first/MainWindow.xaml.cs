using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.IO.Ports;
using System.Linq;
using System.Text;
using System.Threading;
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

namespace first
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private SerialPort serialPort;
        bool firstTime;
        public MainWindow()
        {
            InitializeComponent();
            serialPort = new SerialPort();
             firstTime = true;
        }

        private void Window__Loaded(object sender, RoutedEventArgs e)
        {
            string[] ports = SerialPort.GetPortNames();
            Ports_ComoBox.ItemsSource = ports;

        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            if (Ports_ComoBox.SelectedItem == null)
            {
                MessageBox.Show("Please select a serial port.");
                return;
            }

            try
            {
                serialPort.BaudRate = Convert.ToInt32(baudRate.Text);
                serialPort.Parity = Parity.None;
                serialPort.DataBits = 8;
                if (!serialPort.IsOpen)
                {
                    MessageBox.Show("Serial port is not open. Please open the port before performing this operation.");
                    return;
                }
                StringBuilder stringBuilder = new StringBuilder();

                foreach (Paragraph paragraph in Message_Content.Document.Blocks)
                {
                    foreach (Run run in paragraph.Inlines)
                    {
                        stringBuilder.Append(run.Text);
                    }
                }


				string text = stringBuilder.ToString();
                string hexString = text.Replace(" ", ""); 

                for (int i = 0; i < hexString.Length; i += 2)
                {
                    string hexPair = hexString.Substring(i, 2);

                    if (byte.TryParse(hexPair, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out byte result))
                    {
						// Convert the hex pair to a byte and send it to the serial port
						serialPort.Write(new byte[] { result }, 0, 1);
                        if(i+2==hexString.Length) 
                        {
							Recieve_RichBox.AppendText("Sent : ");
							Recieve_RichBox.AppendText($"{stringBuilder}");
							Recieve_RichBox.AppendText("\n");
						}


					}
                    else
                    {
                        MessageBox.Show($"Invalid hex characters at position {i}.");
                        return;
                    }
                }




            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error: {ex.Message}");
            }

        }

        private void Recieve_btn_Click(object sender, RoutedEventArgs e)
        {

            try
            {
            if (Recieve_btn.Content.ToString() == "Connect")
            {
                if (!serialPort.IsOpen)
                {
                        try
                        {
                            OpenPort();
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show($"Error: {ex.Message}");
                        }
                    }
                Recieve_btn.Content = "Disconnect";
                    if (firstTime == true)
                    {
                        serialPort.DataReceived += SerialPort_DataReceived;
                        firstTime = false;
                    }
            }
            else
            {
                ClosePort();
                Recieve_btn.Content = "Connect";
            }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error: {ex.Message}");
            }
        }

        private void SerialPort_DataReceived(object sender, SerialDataReceivedEventArgs e)
        {
           

            try
            {
                // Read incoming data
                int bytesToRead = serialPort.BytesToRead;
                byte[] buffer = new byte[bytesToRead];
                serialPort.Read(buffer, 0, bytesToRead);

                // Convert the received bytes to a hex string
                string hexString = BitConverter.ToString(buffer).Replace("-", "");

				// Invoke UI updates on the UI thread
				new Thread(() =>
				{
					Application.Current.Dispatcher.Invoke(() =>
					{
						// UI-related code executed on the UI thread
						Recieve_RichBox.AppendText("Received: ");
						Recieve_RichBox.AppendText(hexString);
						Recieve_RichBox.AppendText("\n");
						Recieve_RichBox.ScrollToEnd();
					});
				}).Start();


				//Dispatcher.Invoke(() =>
				//{
				//    // Update UI elements or perform any UI-related actions

				//   Recieve_RichBox.ScrollToEnd();


				//    // Assuming 'Recieve_RichBox' is the RichTextBox control you want to update
				//});

			}


            catch (Exception ex)
            {
                MessageBox.Show($"Error: {ex.Message}");
            }



        }

        private void OpenPort()
        {
            try
            {
                serialPort.PortName = Ports_ComoBox.SelectedItem.ToString();
                serialPort.Open();
            }
            catch (UnauthorizedAccessException ex)
            {
                MessageBox.Show($"Error: {ex.Message}. Make sure you have the necessary permissions and no other application is using the port.");
            }
        }

        private void ClosePort()
        {
            if (serialPort.IsOpen)
            {
                serialPort.Close();
            }
        }

      

        private CancellationTokenSource cancellationTokenSource;

        private async void CheckBox_Checked(object sender, RoutedEventArgs e)
        {
            try
            {
                StringBuilder stringBuilder = new StringBuilder();

                foreach (Paragraph paragraph in Message_Content.Document.Blocks)
                {
                    foreach (Run run in paragraph.Inlines)
                    {
                        stringBuilder.Append(run.Text);
                    }
                }

                string text = stringBuilder.ToString();
                string hexString = text.Replace(" ", "");
                List<byte> byteList = new List<byte>();
                // Assuming inputTime is a TextBox, you need to convert it to a numeric value
                if (int.TryParse(Delay_Input.Text, out int delayMilliseconds))
                {
                    // Initialize the cancellation token source
                    cancellationTokenSource = new CancellationTokenSource();
                    for (int i = 0; i < hexString.Length; i += 2)
                    {
                        string hexPair = hexString.Substring(i, 2);

                        if (byte.TryParse(hexPair, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out byte result))
                        {
                            byteList.Add(result);
                            // Convert the hex pair to a byte and send it to the serial port
                            if (i+2 == hexString.Length)
                            {
                                await SendMessageWithDelay(byteList, delayMilliseconds, cancellationTokenSource.Token, stringBuilder);
                            }
                            }
                        else
                        {
                            MessageBox.Show($"Invalid hex characters at position {i}.");
                            return;
                        }
                    }
                    // Send messages with a delay
                }
                else
                {
                    MessageBox.Show("Invalid input for delay time.");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error: {ex.Message}");
            }
        }

        private async Task SendMessageWithDelay(List<byte> message, int delayMilliseconds, CancellationToken cancellationToken,StringBuilder stringBuilder)
        {
            try
            {
				string hexString = stringBuilder.ToString().ToUpper();
				while (!cancellationToken.IsCancellationRequested)
                    {
                        for (int i = 0; i < message.Count; i++)
                        {
                            serialPort.Write(new byte[] { message[i] }, 0, 1);
						if (message.Count == i + 1)
						{
							Recieve_RichBox.AppendText("Sent: ");
							Recieve_RichBox.AppendText($"{hexString:X2} ");
							Recieve_RichBox.AppendText("\n");
						}
						Recieve_RichBox.ScrollToEnd();
					    }
					// Delay before sending the next message
				    	await Task.Delay(delayMilliseconds, cancellationToken);
                    }
              
                }
			catch (Exception ex)
            {
                MessageBox.Show($"Error: {ex.Message}");
            }
        }

        private void CheckBox_Unchecked(object sender, RoutedEventArgs e)
        {
            // Cancel the ongoing task when CheckBox is unchecked
            cancellationTokenSource?.Cancel();
        }


        private void Delay_Input_TextChanged(object sender, TextChangedEventArgs e)
        {

        }

        private void Clear_btn_Click(object sender, RoutedEventArgs e)
        {
            Recieve_RichBox.Document.Blocks.Clear();

        }
    }
}
