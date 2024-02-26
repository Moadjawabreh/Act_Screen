using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Ports;
using System.Security.Principal;
using System.Windows;
using System.Windows.Controls;
using OfficeOpenXml;
using System.Windows.Media;
using Microsoft.Win32;
using System.Windows.Threading;
using System.ComponentModel;
using System.Windows.Documents;
using OfficeOpenXml.Style;


namespace Act_Screen
{
	public partial class MainWindow : Window
	{
		private List<Actuator> actuators = new List<Actuator>();
		private int currentActuatorIndex = 0;
		private SerialPort ports = new SerialPort();
		private int currentRowIndex = 0;
		private bool isStartEnabled = false;
		private byte[] lastMessageSent;
		private List<RowDefinition> manuallyAddedRows = new List<RowDefinition>();
		private int arrowColumn = 0;
		private int arrowRow = 1;
		private System.Windows.Shapes.Path arrow1;
		private int actNameCount = 0;
		private List<ActionWithCheckboxes> actionList = new List<ActionWithCheckboxes>();
		private int AcIndexForGreenBackGround;
		private DispatcherTimer Arrowtimer;
		private byte[] receivedMessage;
		private DispatcherTimer delayTimer;

		public MainWindow()
		{
			InitializeComponent();
			ports.DataReceived += Com_DataReceived;
			AcIndexForGreenBackGround = 4;

			//Arrow Timer
			Arrowtimer = new DispatcherTimer();
			Arrowtimer.Interval = TimeSpan.FromMilliseconds(0.1);
			Arrowtimer.Tick += Timer_Tick;
			Arrowtimer.Start();

			//Delay Timer
			delayTimer = new DispatcherTimer();
			delayTimer.Interval = TimeSpan.FromSeconds(1.5);
			delayTimer.Tick += DelayTimer_Tick;

			Baud_cmbox.ItemsSource = new object[] { 9600 };
			ports_cmbox.ItemsSource = SerialPort.GetPortNames();
		}


		private void Timer_Tick(object sender, EventArgs e)
		{
			UpdateArrowPosition();
		}

		private void Com_DataReceived(object sender, SerialDataReceivedEventArgs e)
		{

			int bufferSize = 64; // Adjust as needed based on the expected message size

			byte[] buffer = new byte[bufferSize];
			int bytesRead = ports.Read(buffer, 0, bufferSize);

			receivedMessage = new byte[bytesRead];
			Array.Copy(buffer, receivedMessage, bytesRead);

			if (lastMessageSent != null && receivedMessage.Length == lastMessageSent.Length)
			{
				bool match = true;
				for (int i = 0; i < lastMessageSent.Length; i++)
				{
					if (lastMessageSent[i] != receivedMessage[i])
					{
						match = false;
						break;
					}
				}
				ActionWithCheckboxes newAction = new ActionWithCheckboxes();
				if (match)
				{
					UpdateActionBackgroundColor(true);

					Application.Current.Dispatcher.Invoke(() =>
						{
							if (currentRowIndex >= 0 && currentRowIndex < tableGrid.Children.Count)
							{
								if (tableGrid.Children[AcIndexForGreenBackGround] is System.Windows.Controls.Border rowBorder)
								{
									if (rowBorder.Child is Grid rowGrid)
									{
										if (rowGrid.Children.Count > 2 && rowGrid.Children[2] is StackPanel checkBoxStackPanel)
										{
											if (checkBoxStackPanel.Children.Count > 0 && checkBoxStackPanel.Children[0] is RadioButton yesCheckBox)
											{
												checkBoxStackPanel.Children[0].IsEnabled = true;

												yesCheckBox.Checked += (sender, e) =>
												{
													newAction.CheckBox1Checked = true;
													newAction.CheckBox2Checked = false;
													actionList.Add(newAction);
													delayTimer.Start();
												};
											}
											if (checkBoxStackPanel.Children.Count > 0 && checkBoxStackPanel.Children[1] is RadioButton noCheckBox)
											{
												checkBoxStackPanel.Children[1].IsEnabled = true;

												noCheckBox.Checked += (sender, e) =>
												{
													newAction.CheckBox1Checked = false; // Assuming checkBox1 is your CheckBox control
													newAction.CheckBox2Checked = true; // Assuming checkBox2 is your CheckBox control
													actionList.Add(newAction);
													delayTimer.Start();
												};
											}
										}
									}
								}
							}
						});

				}

				else
				{

					UpdateActionBackgroundColor(false);

					Application.Current.Dispatcher.Invoke(() =>
					{

						if (tableGrid.Children[AcIndexForGreenBackGround] is System.Windows.Controls.Border rowBorder &&
							rowBorder.Child is Grid rowGrid &&
								rowGrid.Children.Count > 2 &&
									rowGrid.Children[2] is StackPanel checkBoxStackPanel &&
										checkBoxStackPanel.Children.Count > 0 &&
												checkBoxStackPanel.Children[0] is RadioButton yesCheckBox)
					{
						yesCheckBox.IsEnabled = false;
					}
					if (tableGrid.Children[AcIndexForGreenBackGround] is System.Windows.Controls.Border rowBorder2 &&
							rowBorder2.Child is Grid rowGrid2 &&
								rowGrid2.Children.Count > 2 &&
									rowGrid2.Children[2] is StackPanel checkBoxStackPanel2 &&
										checkBoxStackPanel2.Children.Count > 0 &&
												checkBoxStackPanel2.Children[1] is RadioButton NoCheckBox)
					{
						NoCheckBox.IsEnabled = false;
					}
					});

				}
			}
		}

		private void DelayTimer_Tick(object sender, EventArgs e)
		{
			// Stop the timer
			delayTimer.Stop();

			// Call the NextAction method here
			NextAction();
		}

		private void UpdateActionBackgroundColor(bool isMatch)
		{
			// Find the TextBlock representing the current action
			TextBlock actionTextBlock = GetActionTextBlock(currentRowIndex);

			if (actionTextBlock != null)
			{
				// Update the background color

				Application.Current.Dispatcher.Invoke(() =>
				{
					actionTextBlock.Background = isMatch ? Brushes.LightGreen : Brushes.Red;
				});
			}
		}

		private TextBlock GetActionTextBlock(int rowIndex)
		{
			TextBlock? actionTextBlock = null;

			Application.Current.Dispatcher.Invoke(() =>
			{
				if (rowIndex >= 0 && rowIndex < tableGrid.Children.Count)
				{
					if (tableGrid.Children[AcIndexForGreenBackGround] is System.Windows.Controls.Border rowBorder)
					{
						actionTextBlock = (rowBorder.Child as Grid)?.Children
							.OfType<TextBlock>()
							.FirstOrDefault(tb => Grid.GetColumn(tb) == 1);
					}
				}
			});

			return actionTextBlock;
		}

		private void EnableStartButton()
		{
			startBtn.IsEnabled = true;
			isStartEnabled = true;
		}

		private void AddNewActuator_Click(object sender, RoutedEventArgs e)
		{
			Add_Actuator_Action addWindow = new Add_Actuator_Action();
			addWindow.ShowDialog();
			if (addWindow.DialogResult == true)
			{
				actNameCount++;
				Actuator newActuator = new Actuator();
				newActuator.Actions = addWindow.SelectedActions.Select(action => new ActionWithCheckboxes { Action = action }).ToList();
				actuators.Add(newActuator);
				newActuator.Name = "Ac" + actNameCount;
				AddNewRow(newActuator);

				EnableStartButton();
			}
		}

		private void AddNewRow(Actuator actuator)
		{
			// Create new row definition
			RowDefinition newRow = new RowDefinition();
			newRow.Height = GridLength.Auto;
			tableGrid.RowDefinitions.Add(newRow);
			// Create elements for the new row
			TextBlock nameTextBlock = new TextBlock() { Text = $"{actuator.Name}", Margin = new Thickness(5) };
			TextBlock actionTextBlock = new TextBlock()
			{
				Foreground = Brushes.Black,
				Background = Brushes.LightPink,
				HorizontalAlignment = HorizontalAlignment.Center,
				Margin = new Thickness(5),
				FontFamily = new FontFamily("Arial"),
				FontSize = 14 // Set your desired font size here
			};

			RadioButton checkBox = new RadioButton() { IsEnabled = false, HorizontalAlignment = HorizontalAlignment.Center, Content = "YES", Margin = new Thickness(5) };
			RadioButton checkBox2 = new RadioButton() { IsEnabled = false, HorizontalAlignment = HorizontalAlignment.Center, Content = "NO", Margin = new Thickness(5) };

			if (actuator.CurrentActionIndex >= 0 && actuator.CurrentActionIndex < actuator.Actions.Count)
			{
				byte currentAction = actuator.Actions[actuator.CurrentActionIndex].Action;
				string actionText = GetActionText(currentAction);
				actionTextBlock.Text = $"{actionText}";
			}
			else
			{
				actionTextBlock.Text = "Actions";
			}

			// Create a border to contain the elements
			System.Windows.Controls.Border rowBorder = new System.Windows.Controls.Border()
			{
				Background = Brushes.White,
				BorderBrush = Brushes.Gray,
				BorderThickness = new Thickness(1),
				Margin = new Thickness(0, 5, 0, 5)
			};
			manuallyAddedRows.Add(newRow);


			// Create a grid to contain the elements inside the border
			Grid rowGrid = new Grid();
			rowGrid.ColumnDefinitions.Add(new ColumnDefinition());
			rowGrid.ColumnDefinitions.Add(new ColumnDefinition());
			rowGrid.ColumnDefinitions.Add(new ColumnDefinition());

			// Add elements to the grid
			rowGrid.Children.Add(nameTextBlock);
			Grid.SetColumn(nameTextBlock, 0);

			rowGrid.Children.Add(actionTextBlock);
			Grid.SetColumn(actionTextBlock, 1);

			StackPanel checkBoxStackPanel = new StackPanel();
			checkBoxStackPanel.Orientation = Orientation.Horizontal;
			Grid.SetColumn(checkBoxStackPanel, 2);

			checkBoxStackPanel.Children.Add(checkBox);
			checkBoxStackPanel.Children.Add(checkBox2);

			// Add the stack panel to the grid
			rowGrid.Children.Add(checkBoxStackPanel);

			// Set the grid as the content of the border
			rowBorder.Child = rowGrid;

			// Add the border to the table grid
			tableGrid.Children.Add(rowBorder);
			Grid.SetColumnSpan(rowBorder, 3);
			Grid.SetRow(rowBorder, tableGrid.RowDefinitions.Count - 1);

			// Create and position the arrow if this is the first row
			if (tableGrid.RowDefinitions.Count - 1 == 1)
			{
				// Create and position the arrow
				arrow1 = new System.Windows.Shapes.Path()
				{
					Data = Geometry.Parse("M0,0 L5,5 L10,0 Z"), // Triangle shape pointing right
					Fill = Brushes.Red // Change color as needed
				};
				arrow1.Visibility = Visibility.Visible;

				// Position the arrow on the left side of the row
				Grid.SetColumn(arrow1, arrowColumn);
				Grid.SetRow(arrow1, arrowRow);

				// Add the arrow to the table grid
				tableGrid.Children.Add(arrow1);

			}

			else
			{
				System.Windows.Shapes.Path arrow2 = new System.Windows.Shapes.Path()
				{
					Data = Geometry.Parse("M0,0 L5,5 L10,0 Z"), // Triangle shape pointing right
					Fill = Brushes.Red // Change color as needed
				};
				arrow2.Visibility = Visibility.Hidden;

				// Position the arrow on the left side of the row
				Grid.SetColumn(arrow2, 0);
				Grid.SetRow(arrow2, tableGrid.RowDefinitions.Count - 1);

				// Add the arrow to the table grid
				tableGrid.Children.Add(arrow2);

			}

		}

		private string GetActionText(byte action)
		{
			// You need to define how you want to represent the actions
			// Here's an example mapping action bytes to action names
			switch (action)
			{
				case 0x00:
					return "Up";
				case 0x01:
					return "Down";
				case 0x02:
					return "Left";
				case 0x03:
					return "Right";
				case 0x04:
					return "Unlock";
				case 0x05:
					return "Lock";
				case 0x10:
					return "Close";
				default:
					return "Unknown";
			}
		}

		private void OpenPort()
		{
			try
			{
				if (!ports.IsOpen)
				{
					ports.PortName = ports_cmbox.SelectedItem.ToString();
					ports.BaudRate = Convert.ToInt32(Baud_cmbox.SelectedItem);
					ports.Open();
				}
			}
			catch (Exception ex)
			{
				MessageBox.Show($"Error: {ex.Message}");
			}

		}

		private void startBtn_Click(object sender, RoutedEventArgs e)
		{
			if (currentActuatorIndex >= 0 && currentActuatorIndex <= actuators.Count - 1)
			{
				try
				{
					OpenPort();
					Actuator currentActuator = actuators[currentActuatorIndex];
					byte[] message = { (byte)currentActuatorIndex, currentActuator.CurrentAction.Action };
					ports.Write(message, 0, message.Length);
					lastMessageSent = message;
					startBtn.IsEnabled = false;
					Baud_cmbox.IsEnabled = false;
					ports_cmbox.IsEnabled = false;
				}
				catch (Exception ex)
				{
					MessageBox.Show($"Error: {ex.Message}");
				}


				//NextBtn.IsEnabled = true;
				// Disable start button until finished

			}
		}

		public void NextAction()
		{
			if (currentActuatorIndex >= 0 && currentActuatorIndex < actuators.Count)
			{
				Actuator currentActuator = actuators[currentActuatorIndex];

				// Check if the current action is "Close"
				if (currentActuator.CurrentAction.Action == 0x10) // Assuming 0x10 represents "Close"
				{

					// Disable the Next button
					//NextBtn.IsEnabled = false;

					// Move to the next row or actuator if available
					MoveToNextRowOrActuator();
					AcIndexForGreenBackGround += 2;


				}
				else
				{
					// Move to the next action


					NextAction(currentActuator);

					// Update the UI to reflect the changes
					UpdateRowForCurrentActuator();

					// Send message if needed

					byte[] message = { (byte)currentActuatorIndex, currentActuator.CurrentAction.Action };
					lastMessageSent = message;
					ports.Write(message, 0, message.Length);


				}
			}


		}

		private void MoveToNextRowOrActuator()
		{
			// Move to the next row if available
			if (currentRowIndex < tableGrid.RowDefinitions.Count - 1 && currentActuatorIndex <= actuators.Count - 1)
			{
				currentRowIndex++;
				currentActuatorIndex++;
				startBtn.IsEnabled = true;
				Baud_cmbox.IsEnabled = true;
				ports_cmbox.IsEnabled = true;
				// Update the arrow position

				UpdateArrowPosition();
				CheckIfThereIsNextAct();
			}

			// Enable the Next button if the next action is not "Close"
			//if (currentActuatorIndex < actuators.Count && currentRowIndex < tableGrid.RowDefinitions.Count)
			//{
			//	Actuator nextActuator = actuators[currentActuatorIndex];
			//	if (nextActuator.CurrentAction != 0x10) // Assuming 0x10 represents "Close"
			//	{
			//		NextBtn.IsEnabled = true;
			//	}
			//}
		}

		private void CheckIfThereIsNextAct()
		{
			if (currentActuatorIndex == actuators.Count)
			{
				startBtn.IsEnabled = false;
			}
			else
			{
				startBtn.IsEnabled = true;
			}
		}

		private void UpdateArrowPosition()
		{
			// Find the arrow element
			System.Windows.Shapes.Path? arrow = null;
			foreach (var child in tableGrid.Children)
			{
				if (child is System.Windows.Shapes.Path path && Grid.GetRow(path) == currentRowIndex)
				{
					arrow = path;
					break;
				}
			}

			// If arrow is found
			if (arrow != null)
			{
				RemoveArrowFromGrid();
				arrow.Visibility = Visibility.Visible; // Show the arrow
				Grid.SetRow(arrow, currentRowIndex + 1); // Update its position

			}
		}

		private void RemoveArrowFromGrid()
		{
			// Remove the arrow from the grid
			if (arrow1 != null)
			{
				arrow1.Visibility = Visibility.Hidden;
			}
		}

		public void NextAction(Actuator actuator)
		{
			actuator.CurrentActionIndex++;
			if (actuator.CurrentActionIndex >= actuator.Actions.Count)
				actuator.CurrentActionIndex = 0; // Reset to -1 if reached the end
		}

		private void UpdateRowForCurrentActuator()
		{
			foreach (var row in manuallyAddedRows)
			{
				int rowIndex = tableGrid.RowDefinitions.IndexOf(row);
				if (rowIndex != -1)
				{
					tableGrid.RowDefinitions.RemoveAt(rowIndex); // Remove the RowDefinition objects
				}
			}


			for (int i = tableGrid.Children.Count - 1; i >= 4; i--)
			{
				tableGrid.Children.RemoveAt(i);
			}


			// Clear the list of manually added rows
			manuallyAddedRows.Clear();



			// Add rows for each actuator
			foreach (var actuator in actuators)
			{
				AddNewRow(actuator);
			}
		}

		private void finish_Btn_Click(object sender, RoutedEventArgs e)
		{
			try
			{
				// Create a new Excel package
				using (ExcelPackage excelPackage = new ExcelPackage())
				{
					// Add a worksheet to the Excel package
					ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add("ActuatorData");

					// Add headers to the worksheet
					worksheet.Cells["A1"].Value = "Actuator";
					worksheet.Cells["B1"].Value = "Action";
					worksheet.Cells["C1"].Value = "Completed";
					worksheet.Cells["D1"].Value = "Not Completed";

					// Make the cells bold
					using (var range3 = worksheet.Cells["A1:D1"])
					{
						range3.Style.Font.Bold = true;
					}


					int row = 2;

					// Iterate over each actuator
					foreach (var actuator in actuators)
					{
						foreach (var actionWithCheckBoxes in actuator.Actions)
						{
							worksheet.Cells[row, 1].Value = actuator.Name; // Assuming Actuator has a 'Name' property
							worksheet.Cells[row, 2].Value = GetActionText(actionWithCheckBoxes.Action);

							row++;
						}

					}


					// Reset row to 2 for the next loop
					row = 2;

					// Define formatting for the entire range
					var range = worksheet.Cells[row, 3, row + actionList.Count - 1, 4];
					range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
					range.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
					range.Style.Font.Bold = true;
					range.Style.Fill.PatternType = ExcelFillStyle.Solid;
					range.Style.Fill.BackgroundColor.SetColor(System.Drawing.ColorTranslator.FromHtml("#ADD8E6")); // Light Blue color

					// Write data from actionList to Excel worksheet
					foreach (var actionWithCheckboxes in actionList)
					{
						worksheet.Cells[row, 3].Value = actionWithCheckboxes.CheckBox1Checked ? "Yes" : "No";
						worksheet.Cells[row, 4].Value = actionWithCheckboxes.CheckBox2Checked ? "Yes" : "No";
						row++;
					}

					// Auto-fit columns for better visibility
					worksheet.Cells.AutoFitColumns();

					// Open a file save dialog
					SaveFileDialog saveFileDialog = new SaveFileDialog();
					saveFileDialog.Filter = "Excel Files|*.xlsx|All Files|*.*";
					saveFileDialog.DefaultExt = "xlsx";
					saveFileDialog.Title = "Save Excel File";

					// Show the file save dialog
					if (saveFileDialog.ShowDialog() == true)
					{
						// Save the Excel package to the selected file
						string filePath = saveFileDialog.FileName;
						FileInfo excelFile = new FileInfo(filePath);
						excelPackage.SaveAs(excelFile);

						MessageBox.Show("Data saved successfully to " + filePath);
					}
				}
			}
			catch (Exception ex)
			{
				MessageBox.Show("An error occurred while saving data: " + ex.Message);
			}
		}

		private void ports_cmbox_SelectionChanged(object sender, SelectionChangedEventArgs e)
		{
			ports.Close();
		}
	}

	public class ActionWithCheckboxes
	{
		public byte Action { get; set; }
		public bool CheckBox1Checked { get; set; }
		public bool CheckBox2Checked { get; set; }
	}

	public class Actuator
	{
		public string? Name { get; set; }
		public List<ActionWithCheckboxes> Actions { get; set; }
		public int CurrentActionIndex { get; set; } = 0;
		private bool IsCurrent { get; set; } = false;

		public ActionWithCheckboxes? CurrentAction
		{
			get
			{
				if (CurrentActionIndex >= 0 && CurrentActionIndex < Actions.Count)
					return Actions[CurrentActionIndex];
				return null; 
			}
		}
	}


}
