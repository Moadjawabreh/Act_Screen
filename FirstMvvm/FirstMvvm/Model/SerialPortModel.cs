using System;
using System.Globalization;
using System.IO.Ports;
using System.Text;
using System.Windows;
using System.Windows.Documents;
using System.Windows.Threading;
using static System.Net.Mime.MediaTypeNames;

namespace FirstMvvm.Model
{
	public class SerialPortModel
	{
		private SerialPort serialPort;
		private readonly Dispatcher dispatcher;

		public SerialPortModel()
		{
			serialPort = new SerialPort();
			this.dispatcher = Dispatcher.CurrentDispatcher;
			// Initialize other serial port settings if needed

			serialPort.DataReceived += SerialPort_DataReceived;
		}

		public event EventHandler<string> MessageReceived;
		private void OnMessageReceived(string message)
		{
			MessageReceived?.Invoke(this, message);
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
				dispatcher.Invoke(() =>
				{
					// Notify subscribers (e.g., the ViewModel) that a message has been received
					OnMessageReceived(hexString);

				});
			}
			catch (Exception ex)
			{
				MessageBox.Show($"Error: {ex.Message}");
			}
		}

		public string[] GetAvailablePorts()
		{
			return SerialPort.GetPortNames();
		}

		public void OpenPort(string portName, int baudRate)
		{
			try
			{
				serialPort.PortName = portName;
				serialPort.BaudRate = baudRate;
				serialPort.Open();

			}
			catch (UnauthorizedAccessException ex)
			{
				MessageBox.Show($"Error: {ex.Message}. Make sure you have the necessary permissions and no other application is using the port.");
			}
		}

		public void ClosePort()
		{
			if (serialPort.IsOpen)
			{
				serialPort.Close();
			}
		}



		public void SendMessage(string message)
		{

			try
			{
				if (serialPort.IsOpen)
				{
					string hexString = message.Replace(" ", "");
					for(int i = 0; i < hexString.Length; i += 2)
					{
						string hexPair = hexString.Substring(i, 2);

						if (byte.TryParse(hexPair, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out byte result))
						{
							// Convert the hex pair to a byte and send it to the serial port

							serialPort.Write(new byte[] { result }, 0, 1);
						
						}
					}
				}
				else
				{
					MessageBox.Show($"Error: Please open the port. ");
				}

			}
			catch (Exception ex)
			{
				MessageBox.Show($"Error: {ex.Message}. ");
			}
		}

	}
}
