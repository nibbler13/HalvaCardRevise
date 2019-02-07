using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Input;

namespace HalvaCardRevise {
	class MainWindowViewModel : INotifyPropertyChanged {
		public event PropertyChangedEventHandler PropertyChanged;
		private void NotifyPropertyChanged([CallerMemberName] string propertyName = "") {
			PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
		}
		

		private ICommand buttonClick;
		public ICommand ButtonClick {
			get {
				return buttonClick ?? 
					(buttonClick = new CommandHandler((object parameter) => 
					Action(parameter))); }
		}


		private bool CanExecuteRemoveFileSber {
			get { return ListFileSber.Where(x => x.IsSelected).Count() > 0; }
		}
		private ICommand buttonRemoveFileSberClick;
		public ICommand ButtonRemoveFileSberClick {
			get {
				return buttonRemoveFileSberClick ??
					(buttonRemoveFileSberClick = new CommandHandler((object parameter) => 
					RemoveFileFromList(parameter), param => CanExecuteRemoveFileSber));
			}
		}

		private bool CanExecuteRemoveFileVtb {
			get { return ListFileVtb.Where(x => x.IsSelected).Count() > 0; }
		}
		private ICommand buttonRemoveFileVtbClick;
		public ICommand ButtonRemoveFileVtbClick {
			get {
				return buttonRemoveFileVtbClick ??
					(buttonRemoveFileVtbClick = new CommandHandler((object parameter) =>
					RemoveFileFromList(parameter), param => CanExecuteRemoveFileVtb));
			}
		}
		

		private ItemFileInfo halvaCardReportFile;
		public ItemFileInfo HalvaCardReportFile {
			get { return halvaCardReportFile; }
			private set {
				if (value != halvaCardReportFile) {
					halvaCardReportFile = value;
					NotifyPropertyChanged();
				}
			}
		}


		public ObservableCollection<ItemFileInfo> ListFileSber { get; set; }
		public ObservableCollection<ItemFileInfo> ListFileVtb { get; set; }


		private Visibility gridMainVisibility;
		public Visibility GridMainVisibility {
			get { return gridMainVisibility; }
			private set {
				if (value != gridMainVisibility) {
					gridMainVisibility = value;
					NotifyPropertyChanged();
				}
			}
		}

		private Visibility gridProgressVisibility;
		public Visibility GridProgressVisibility {
			get { return gridProgressVisibility; }
			private set {
				if (value != gridProgressVisibility) {
					gridProgressVisibility = value;
					NotifyPropertyChanged();
				}
			}
		}



		private string progressInfo;
		public string ProgressInfo {
			get { return progressInfo; }
			private set {
				if (value != progressInfo) {
					progressInfo = value;
					NotifyPropertyChanged();
				}
			}
		}

		private double progressValue;
		public int ProgressValue {
			get { return (int)progressValue; } 
			private set {
				if (value != progressValue) {
					progressValue = value;
					NotifyPropertyChanged();
				}
			}
		}


		public MainWindowViewModel() {
			ListFileSber = new ObservableCollection<ItemFileInfo>();
			ListFileVtb = new ObservableCollection<ItemFileInfo>();
			GridMainVisibility = Visibility.Visible;
			GridProgressVisibility = Visibility.Hidden;

			if (Debugger.IsAttached) {
				halvaCardReportFile = new ItemFileInfo(
					@"C:\Users\nn-admin\Desktop\Халва\Халва\" + 
					@"ООО _Клиника ЛМС__7704544391_И_10.12.2018-0.csv", 
					ItemFileInfo.FileType.Halva);

				ListFileSber.Add(new ItemFileInfo(
					@"C:\Users\nn-admin\Desktop\Халва\Сбербанк\" + 
					@"3248426 (3248426_СРБ) - Москва сретенка.xlsx", 
					ItemFileInfo.FileType.Sberbank));
				ListFileSber.Add(new ItemFileInfo(
					@"C:\Users\nn-admin\Desktop\Халва\Сбербанк\" +
					@"3277308 d.toropov@bzklinika.ru - Сочи.xlsx",
					ItemFileInfo.FileType.Sberbank));
				ListFileSber.Add(new ItemFileInfo(
					@"C:\Users\nn-admin\Desktop\Халва\Сбербанк\"+
					@"Отчет о возмещении денежных средств предприятию ООО Клиника ЛМС с 01.11.2018г. по 16.01.2019г. -Каменск-Уральский.xlsx",
					ItemFileInfo.FileType.Sberbank));

				ListFileVtb.Add(new ItemFileInfo(
					@"C:\Users\nn-admin\Desktop\Халва\ВТБ\"+
					@"h_01-11-2018-16-01-2019_ООО_Клиника_ЛМС_-ret_innxxx4391 - Казань.xlsx", 
					ItemFileInfo.FileType.Vtb));
				ListFileVtb.Add(new ItemFileInfo(
					@"C:\Users\nn-admin\Desktop\Халва\ВТБ\" +
					@"h_01-11-2018-16-01-2019_ООО_Клиника_ЛМС_-ret_innxxx4391 - Питер.xlsx",
					ItemFileInfo.FileType.Vtb));
				ListFileVtb.Add(new ItemFileInfo(
					@"C:\Users\nn-admin\Desktop\Халва\ВТБ\" +
					@"h_01-11-2018-16-01-2019_ООО_Клиника_ЛМС_-ret_innxxx5391 - Москва Сущевка, фрунзенская.xlsx",
					ItemFileInfo.FileType.Vtb));
				ListFileVtb.Add(new ItemFileInfo(
					@"C:\Users\nn-admin\Desktop\Халва\ВТБ\" + 
					@"h_01-11-2018-16-01-2019_Филиал_ООО_Клиника_ЛМС_в_г_Краснодаре-ret_innxxx4391 - Краснодар.xlsx",
					ItemFileInfo.FileType.Vtb));
				ListFileVtb.Add(new ItemFileInfo(
					@"C:\Users\nn-admin\Desktop\Халва\ВТБ\" + 
					@"h_01-11-2018-16-01-2019_Филиал_ООО_Клиника_ЛМС_в_г_Уфе-ret_innxxx4391 - Уфа.xlsx",
					ItemFileInfo.FileType.Vtb));
			}
		}

		public void Action(object parameter) {
			string param = parameter.ToString();

			if (param.Equals("SelectHalvaFile")) {
				if (SelectScvFile(out string selectedFile))
					HalvaCardReportFile = new ItemFileInfo(selectedFile, ItemFileInfo.FileType.Halva);
			} else if (param.Equals("AddFileSber"))
				SelectExcelFiles(ListFileSber, ItemFileInfo.FileType.Sberbank);
			else if (param.Equals("AddFileVtb"))
				SelectExcelFiles(ListFileVtb, ItemFileInfo.FileType.Vtb);
			else if (param.Equals("ExecuteRevise"))
				ExecuteRevise();
		}

		private void ExecuteRevise() {
			string warningMessage = string.Empty;

			if (HalvaCardReportFile == null)
				warningMessage += "Не выбран файл с отчетом по картам 'Халва'" + 
					Environment.NewLine;

			if (ListFileSber.Count + ListFileVtb.Count == 0)
				warningMessage += "Не выбрано ни одного файла с отчетом по " +
					"терминалам 'Сбербанк' / 'ВТБ'" + Environment.NewLine;

			if (!string.IsNullOrEmpty(warningMessage)) {
				MessageBox.Show(
					Application.Current.MainWindow,
					warningMessage,
					"Выполнение сверки невозможно выполнить",
					MessageBoxButton.OK,
					MessageBoxImage.Warning);
				return;
			}

			ProgressInfo = string.Empty;
			ProgressValue = 0;
			GridMainVisibility = Visibility.Hidden;
			GridProgressVisibility = Visibility.Visible;

			BackgroundWorker backgroundWorker = new BackgroundWorker {
				WorkerReportsProgress = true
			};

			backgroundWorker.DoWork += BackgroundWorker_DoWork;
			backgroundWorker.ProgressChanged += BackgroundWorker_ProgressChanged;
			backgroundWorker.RunWorkerCompleted += BackgroundWorker_RunWorkerCompleted;
			backgroundWorker.RunWorkerAsync();
		}

		private void BackgroundWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e) {
			if (e.Error != null)
				UpdateProgressInfo(e.Error.Message + Environment.NewLine + e.Error.StackTrace);

			MessageBox.Show(
				Application.Current.MainWindow,
				"Выполнение завершено",
				string.Empty,
				MessageBoxButton.OK,
				MessageBoxImage.Information);
		}

		private void BackgroundWorker_ProgressChanged(object sender, ProgressChangedEventArgs e) {
			ProgressValue = e.ProgressPercentage;

			if (!string.IsNullOrEmpty(e.UserState.ToString()))
				UpdateProgressInfo(e.UserState.ToString());
		}

		private void BackgroundWorker_DoWork(object sender, DoWorkEventArgs e) {
			List<ItemFileInfo> terminalFiles = new List<ItemFileInfo>();
			terminalFiles.AddRange(ListFileSber);
			terminalFiles.AddRange(ListFileVtb);
			Revise revise = new Revise(HalvaCardReportFile, terminalFiles, sender as BackgroundWorker);
			revise.DoRevise();
		}

		private void UpdateProgressInfo(string text) {
			ProgressInfo += DateTime.Now.ToLongTimeString() + ": " + 
				text + Environment.NewLine;
		}

		private void SelectExcelFiles(ObservableCollection<ItemFileInfo> collection, ItemFileInfo.FileType type) {
			if (!SelectExcelFiles(out string[] selectedFiles))
				return;

			foreach (string selectedFile in selectedFiles)
				if (collection.Where(x => x.FullPath.Equals(selectedFile)).Count() == 0)
					collection.Add(new ItemFileInfo(selectedFile, type));
		}

		public void RemoveFileFromList(object parameter) {
			string param = parameter.ToString();

			if (param.Equals("RemoveFileSber"))
				RemoveSelectedItemsFromCollection(ListFileSber);
			else if (param.Equals("RemoveFileVtb"))
				RemoveSelectedItemsFromCollection(ListFileVtb);
		}

		private void RemoveSelectedItemsFromCollection(ObservableCollection<ItemFileInfo> collection) {
			List<ItemFileInfo> toRemove = collection.Where(x => x.IsSelected).ToList();
			foreach (ItemFileInfo item in toRemove)
				collection.Remove(item);
		}

		private bool SelectScvFile(out string selectedFile) {
			OpenFileDialog openFileDialog = GetOpenFileDialog("Файл CSV (*.csv)|*.csv", false);
			
			if (openFileDialog.ShowDialog() == true) {
				selectedFile = openFileDialog.FileName;
				return true;
			} else {
				selectedFile = string.Empty;
				return false;
			}

		}

		private bool SelectExcelFiles(out string[] selectedFiles) {
			OpenFileDialog openFileDialog = GetOpenFileDialog("Книга Excel (*.xls*)|*.xls*", true);
			
			if (openFileDialog.ShowDialog() == true) {
				selectedFiles = openFileDialog.FileNames;
				return true;
			} else {
				selectedFiles = new string[0];
				return false;
			}
		}

		private OpenFileDialog GetOpenFileDialog(string filter, bool isMultiselect) {
			OpenFileDialog openFileDialog = new OpenFileDialog {
				Filter = filter,
				CheckFileExists = true,
				CheckPathExists = true,
				Multiselect = isMultiselect,
				RestoreDirectory = true
			};

			return openFileDialog;
		}
	}
}
