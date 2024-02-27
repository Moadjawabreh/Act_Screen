using FirstMvvm.Model;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO.Ports;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Threading;

namespace FirstMvvm.ViewModel
{
	public class ViewModelS : SerialProtViewModel, INotifyPropertyChanged
	{
		private SerialPortModel serialPortModel;
		private DispatcherTimer portCheckTimer;

		public event EventHandler<string> MessageSent;
		public ObservableCollection<string> AvailablePorts { get; set; }

		private string selectedPort;
		private RichTextBox recieveRichBox;  // Reference to the RichTextBox in MainWindow
		private string serialPortBuffer;
		private int portBaudRate;
		private bool CheckIfConnect;

		private string lastConnectedPort;
		private int lastConnectedBaudRate;

		public ViewModelS(RichTextBox recieveRichBox)
		{
			serialPortModel = new SerialPortModel();
			AvailablePorts = new ObservableCollection<string>(serialPortModel.GetAvailablePorts());
			this.recieveRichBox = recieveRichBox;
			serialPortModel.MessageReceived += SerialPort_DataReceived;
			portCheckTimer = new DispatcherTimer();
			portCheckTimer.Interval = TimeSpan.FromSeconds(1);
			portCheckTimer.Tick += PortCheckTimer_Tick;
			portCheckTimer.Start();
			CheckIfConnect = false;
		}
		public string SelectedPort
		{
			get { return selectedPort; }
			set { SetProperty(ref selectedPort, value, nameof(SelectedPort)); }
		}









		private string connectBtnContent = "Connect";

		public string ConnectBtnContent
		{
			get { return connectBtnContent; }
			set
			{
				if (connectBtnContent != value)
				{
					connectBtnContent = value;
					OnPropertyChanged(nameof(ConnectBtnContent));
				}
			}
		}


		private void PortCheckTimer_Tick(object sender, EventArgs e)
		{
			if (!string.IsNullOrEmpty(selectedPort) && !serialPortModel.GetAvailablePorts().Contains(selectedPort) && IsPortConnected())
			{
				SelectedPort = null;
				MessageBox.Show($"The selected port {selectedPort} is disconnected.");
				ConnectBtnContent = "Connect";
			}

			HashSet<string> existingPorts = new HashSet<string>(AvailablePorts);

			// Get the new available ports
			var newPorts = serialPortModel.GetAvailablePorts().Except(existingPorts);

			// Add the new ports to the collection
			foreach (var newPort in newPorts)
			{
				AvailablePorts.Add(newPort);

			}

			// Remove any ports that are no longer available
			var removedPorts = existingPorts.Except(serialPortModel.GetAvailablePorts());
			foreach (var removedPort in removedPorts)
			{
				AvailablePorts.Remove(removedPort);
				
			}
			// If the selected port is null and there is a last connected port, attempt to reconnect
			if (string.IsNullOrEmpty(selectedPort) && !string.IsNullOrEmpty(lastConnectedPort) && AvailablePorts.Contains(lastConnectedPort))
			{
				SelectedPort = lastConnectedPort;
				OpenPort(lastConnectedBaudRate);
			}
		}

		public bool IsPortConnected()
		{
			return CheckIfConnect;
		}

		private void SerialPort_DataReceived(object sender, string message)
		{
			// Handle the received message, update UI, etc.
			// For example, append the message to the RichTextBox directly
			AppendTextToRichTextBox($"Received: {message}\n");
		}




		private void AppendTextToRichTextBox(string text)
		{
			if (recieveRichBox != null)
			{
				// Invoke UI updates on the UI thread
				recieveRichBox.Dispatcher.Invoke(() =>
				{
					// Append text to the RichTextBox
					recieveRichBox.AppendText(text);
					// Optionally, scroll to the end of the RichTextBox
					recieveRichBox.ScrollToEnd();
				});
			}
		}

		public void OpenPort(int BaudRate)
		{

			//serialPortBuffer = SelectedPort;
			//serialPortModel.OpenPort(serialPortBuffer, BaudRate); // Set your desired baud rate
			//CheckIfConnect = true;
			//ConnectBtnContent = "Disconnect"; // Update button content when connected
			if (serialPortBuffer == null)
			{
				serialPortBuffer = selectedPort;
				lastConnectedPort = selectedPort;
				lastConnectedBaudRate = BaudRate;
				serialPortModel.OpenPort(serialPortBuffer, BaudRate); // Set your desired baud rate
				CheckIfConnect = true;
				ConnectBtnContent = "Disconnect";
			}

			else
			{
				serialPortModel.OpenPort(serialPortBuffer, BaudRate); // Set your desired baud rate
				CheckIfConnect = true;
				serialPortBuffer = null; 
				lastConnectedPort = null;
				ConnectBtnContent = "Disconnect";
			}


		}
		public void ReconnectToLastPort()
		{
			if (!string.IsNullOrEmpty(lastConnectedPort))
			{
				OpenPort(lastConnectedBaudRate);
			}
		}
		public void ClosePort()
		{
			CheckIfConnect = false;
			serialPortModel.ClosePort();
			ConnectBtnContent = "Connect"; // Update button content when connected

		}

		private void OnMessageSent(string message)
		{
			MessageSent?.Invoke(this, message);
		}
		public void SendMessage(string message)
		{
			if (IsPortConnected())
			{
				if (message.Length % 2 == 0)
				{
					serialPortModel.SendMessage(message);
					OnMessageSent($"Sent: {message.ToUpper()}");
				}

				// Notify subscribers (e.g., the View) that a message has been sent
				else
				{
					MessageBox.Show($"Invalid hex characters at position .");
					return;
				}
			}
			else
			{
				MessageBox.Show($"Error: Please open the port. ");
			}

		}


		public void SelectedChanges()
		{
			ClosePort();
			serialPortBuffer = null; 
			ConnectBtnContent = "Connect";
		}

	}
}
