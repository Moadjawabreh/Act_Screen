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
using System.Windows.Shapes;
using System.Windows.Controls.Primitives;


namespace Act_Screen
{
	public partial class MainWindow : Window
	{
		private List<Actuator> actuators = new List<Actuator>();
		private int currentActuatorIndex = 0;
		private SerialPort ports = new SerialPort();
		private int currentRowIndex = 0;
		private bool permitionToEnable = false;
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
			System.Windows.Controls.Border actionAction = GetBorderAction(currentRowIndex);

			// Find the TextBlock representing the current action
			TextBlock actionTextBlock = GetActionTextBlock(currentRowIndex);


			if (actionAction != null)
			{
				// Update the background color

				Application.Current.Dispatcher.Invoke(() =>
				{
					actionAction.Background = isMatch ? Brushes.LightGreen : Brushes.Red;
					actionTextBlock.Background = isMatch ? Brushes.LightGreen : Brushes.Red;
				});
			}


		}

		private System.Windows.Controls.Border GetBorderAction(int rowIndex)
		{
			System.Windows.Controls.Border childBorder = null;

			Application.Current.Dispatcher.Invoke(() =>
			{
				if (rowIndex >= 0 && rowIndex < tableGrid.Children.Count)
				{
					if (tableGrid.Children[AcIndexForGreenBackGround] is System.Windows.Controls.Border border)
					{
						// Check if the border contains at least two children
						if (border.Child is Grid innerGrid && innerGrid.Children.Count >= 2)
						{
							// Access the second child (index 1) of the inner grid
							if (innerGrid.Children[1] is System.Windows.Controls.Border secondBorder)
							{
								childBorder = secondBorder;
							}
						}
					}
				}



			});


			return childBorder;
		}

		private TextBlock GetActionTextBlock(int rowIndex)
		{
			TextBlock actionTextBlock = null;

			Application.Current.Dispatcher.Invoke(() =>
			{
				if (rowIndex >= 0 && rowIndex < tableGrid.Children.Count)
				{
					if (tableGrid.Children[AcIndexForGreenBackGround] is System.Windows.Controls.Border border)
					{
						// Check if the border contains at least two children
						if (border.Child is Grid innerGrid && innerGrid.Children.Count >= 2)
						{
							// Access the second child (index 1) of the inner grid
							if (innerGrid.Children[1] is System.Windows.Controls.Border secondBorder)
							{
								// Find the TextBlock inside the second border
								actionTextBlock = secondBorder.Child as TextBlock;
							}
						}
					}
				}
			});

			return actionTextBlock;
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

				if (!permitionToEnable)
				{
					startBtn.IsEnabled = true;
				}
			}
		}

		private void AddNewRow(Actuator actuator)
		{
			RowDefinition newRow = new RowDefinition();
			newRow.Height = GridLength.Auto;
			tableGrid.RowDefinitions.Add(newRow);

			TextBlock nameTextBlock = new TextBlock() { Text = $"{actuator.Name}", Margin = new Thickness(5) };
			TextBlock actionTextBlock = new TextBlock()
			{
				Foreground = Brushes.Black,
				Background = Brushes.LightPink,
				HorizontalAlignment = HorizontalAlignment.Center,
				VerticalAlignment = VerticalAlignment.Center, 
				Margin = new Thickness(5),
				FontFamily = new FontFamily("Arial"),
				FontSize = 14 
			};

			// Create a border around the TextBlock
			System.Windows.Controls.Border border = new System.Windows.Controls.Border()
			{
				BorderBrush = Brushes.Black, 
				CornerRadius = new CornerRadius(3),
				Padding = new Thickness(5),
				Width = 70,
				Margin = new Thickness(0, 2, 0, 2),
				Background = Brushes.LightPink,


			};
			// Create RadioButton for "YES"
			// Create RadioButton for "YES"
			RadioButton yesRadioButton = new RadioButton()
			{
				IsEnabled = false,
				HorizontalAlignment = HorizontalAlignment.Center,
				Margin = new Thickness(5),
				Template = CreateRadioButtonTemplate(Brushes.Green, Brushes.Transparent),
				Content = "Yes"
			};

			yesRadioButton.ToolTip = "YES";
			yesRadioButton.MouseMove += (sender, e) =>
			{
				if (IsMouseOverCircle(sender as RadioButton, e.GetPosition(sender as IInputElement)))
				{
					ShowToolTip(sender as RadioButton);
				}
				else
				{
					// Hide the tooltip if the mouse is not over the circle
					HideToolTip(sender as RadioButton);
				}
			};

			yesRadioButton.Checked += (sender, e) => UpdateRadioButtonTemplate(sender as RadioButton, Brushes.Green);

			// Create RadioButton for "NO"
			RadioButton noRadioButton = new RadioButton()
			{
				IsEnabled = false,
				HorizontalAlignment = HorizontalAlignment.Center,
				Margin = new Thickness(5),
				Template = CreateRadioButtonTemplate(Brushes.Red, Brushes.Transparent),
				Content = "No"
			};

			noRadioButton.ToolTip = "NO";
			noRadioButton.MouseMove += (sender, e) =>
			{
				if (IsMouseOverCircle(sender as RadioButton, e.GetPosition(sender as IInputElement)))
				{
					ShowToolTip(sender as RadioButton);
				}
				else
				{
					// Hide the tooltip if the mouse is not over the circle
					HideToolTip(sender as RadioButton);
				}
			};
			noRadioButton.Checked += (sender, e) => UpdateRadioButtonTemplate(sender as RadioButton, Brushes.Red);

			void ShowToolTip(RadioButton radioButton)
			{
				if (radioButton != null && !string.IsNullOrEmpty(radioButton.ToolTip as string))
				{
					ToolTip tt = new ToolTip();
					tt.Content = radioButton.ToolTip;
					radioButton.ToolTip = tt;
					tt.IsOpen = true;
				}
			}

			bool IsMouseOverCircle(RadioButton radioButton, Point mousePosition)
			{
				if (radioButton != null)
				{
					// Calculate the center point of the circle
					double centerX = radioButton.ActualWidth / 2;
					double centerY = radioButton.ActualHeight / 2;

					// Calculate the distance between the mouse position and the center of the circle
					double distance = Math.Sqrt(Math.Pow(mousePosition.X - centerX, 2) + Math.Pow(mousePosition.Y - centerY, 2));

					// Check if the distance is less than the radius of the circle
					return distance <= radioButton.ActualWidth / 2;
				}

				return false;
			}

			void HideToolTip(RadioButton radioButton)
			{
				if (radioButton != null)
				{
					ToolTip tt = radioButton.ToolTip as ToolTip;
					if (tt != null)
					{
						tt.IsOpen = false;
					}
				}
			}


			void UpdateRadioButtonTemplate(RadioButton radioButton, Brush brush)
			{
				if (radioButton != null && radioButton.Template != null)
				{
					Grid grid = (Grid)radioButton.Template.FindName("RadioButtonGrid", radioButton);
					if (grid != null)
					{
						Ellipse outerCircle = (Ellipse)grid.FindName("OuterCircle");
						if (outerCircle != null)
						{
							outerCircle.Fill = brush;
						}
					}
				}
			}

			// Function to create custom control template for RadioButton






			ControlTemplate CreateRadioButtonTemplate(Brush brush, Brush backgroundBrush)
			{
				ControlTemplate template = new ControlTemplate(typeof(RadioButton));

				// Create Grid to hold the circuit
				var grid = new FrameworkElementFactory(typeof(Grid));
				grid.Name = "RadioButtonGrid";

				// Create outer circle for the circuit
				var outerCircle = new FrameworkElementFactory(typeof(Ellipse));
				outerCircle.SetValue(Ellipse.WidthProperty, 30.0);
				outerCircle.SetValue(Ellipse.HeightProperty, 30.0);
				outerCircle.SetValue(Shape.StrokeProperty, brush);
				outerCircle.SetValue(Shape.StrokeThicknessProperty, 2.0);
				outerCircle.SetValue(Shape.FillProperty, backgroundBrush);
				outerCircle.Name = "OuterCircle";

				// Add the outer circle to the grid
				grid.AppendChild(outerCircle);

				// Create inner circle for the radio button
				var innerCircle = new FrameworkElementFactory(typeof(Ellipse));
				innerCircle.SetValue(Ellipse.WidthProperty, 16.0);
				innerCircle.SetValue(Ellipse.HeightProperty, 16.0);
				innerCircle.SetValue(Ellipse.FillProperty, brush);
				innerCircle.SetValue(Grid.HorizontalAlignmentProperty, HorizontalAlignment.Center);
				innerCircle.SetValue(Grid.VerticalAlignmentProperty, VerticalAlignment.Center);

				// Add the inner circle to the grid
				grid.AppendChild(innerCircle);

				// Add the grid to the template
				template.VisualTree = grid;
				return template;
			}








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


			border.Child = actionTextBlock; // Set the TextBlock as the child of the border


			rowGrid.Children.Add(border); // Add the border to the grid instead of the TextBlock

			Grid.SetColumn(border, 1); // Set the column for the border


			StackPanel checkBoxStackPanel = new StackPanel();
			checkBoxStackPanel.Orientation = Orientation.Horizontal;
			Grid.SetColumn(checkBoxStackPanel, 2);

			checkBoxStackPanel.Children.Add(yesRadioButton);
			checkBoxStackPanel.Children.Add(noRadioButton);

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
					Data = Geometry.Parse("M0,0 L5,5 L10,0 Z"), 
					Fill = Brushes.Red 
				};
				arrow1.Visibility = Visibility.Visible;

				Grid.SetColumn(arrow1, arrowColumn);
				Grid.SetRow(arrow1, arrowRow);

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
					label_Color.Background = Brushes.Green;
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

					permitionToEnable = true;
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
				permitionToEnable = false;
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
					worksheet.Cells["C1"].Value = "Status";

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
					var range = worksheet.Cells[row, 3, row + actionList.Count - 1, 3]; // Specify only column 3
					range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
					range.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
					range.Style.Font.Bold = true;
					range.Style.Fill.PatternType = ExcelFillStyle.Solid;
					range.Style.Fill.BackgroundColor.SetColor(System.Drawing.ColorTranslator.FromHtml("#ADD8E6")); // Light Blue color


					// Write data from actionList to Excel worksheet
					foreach (var actionWithCheckboxes in actionList)
					{
						if (actionWithCheckboxes.CheckBox1Checked == true)
						{
							worksheet.Cells[row, 3].Value = "YES";
							worksheet.Cells[row, 3].Style.Fill.PatternType = ExcelFillStyle.Solid;
							worksheet.Cells[row, 3].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Green); // Set background color to green
						}
						else
						{
							worksheet.Cells[row, 3].Value = "No";
							worksheet.Cells[row, 3].Style.Fill.PatternType = ExcelFillStyle.Solid;
							worksheet.Cells[row, 3].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Red); // Set background color to red
						}
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
			label_Color.Background = Brushes.Red;

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
