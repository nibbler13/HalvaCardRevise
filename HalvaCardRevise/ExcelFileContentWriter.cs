using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace HalvaCardRevise {
	class ExcelFileContentWriter {
		private static readonly string AssemblyDirectory = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + "\\";
		private static Action<double, string> UpdateProgress;
		private static double ProgressCurrent;

		public static string WriteFileInfoToExcel(ItemFileInfo itemFileInfo,
											 Action<double, string> updateProgress,
											 double progressCurrent) {
			UpdateProgress = updateProgress ?? throw new ArgumentNullException("updateProgress");
			ProgressCurrent = progressCurrent;
			IWorkbook workbook = null;
			ISheet sheet = null;
			string resultFile = string.Empty;

			UpdateProgress(ProgressCurrent, "Создание новой книги Excel из шаблона");
			if (!CreateNewIWorkbook(itemFileInfo.FileName, "Template.xlsx",
				out workbook, out sheet, out resultFile, "Данные"))
				return string.Empty;

			int rowNumber = 1;
			int columnNumber = 0;

			UpdateProgress(ProgressCurrent, "Запись данных");
			foreach (FileContent fileContent in itemFileInfo.FileContents) {
				if (string.IsNullOrEmpty(fileContent.TransactionProcessingDate))
					continue;

				IRow row = null;
				try { row = sheet.GetRow(rowNumber); } catch (Exception) { }

				if (row == null)
					row = sheet.CreateRow(rowNumber);

				object[] values = new object[] {
					fileContent.TransactionProcessingDate,
					fileContent.TransactionCommittingDate,
					fileContent.TransactionCommitingTime,
					fileContent.CustomerName,
					fileContent.CustomerINN == string.Empty ?
						string.Empty : "'" + fileContent.CustomerINN,
					fileContent.CustomerStoreName,
					fileContent.CustomerStoreAddress,
					fileContent.TerminalIdentificator == string.Empty ?
						string.Empty : "'" + fileContent.TerminalIdentificator,
					fileContent.CardNumber,
					fileContent.OperationAmount,
					fileContent.TotalReward,
					fileContent.CompanyReward,
					fileContent.City,
					fileContent.AuthorizationCode == string.Empty ?
						string.Empty : "'" + fileContent.AuthorizationCode,
					fileContent.UniqueOperationNumberRNN == string.Empty ?
						string.Empty : "'" + fileContent.UniqueOperationNumberRNN,
					fileContent.CoincidenceType,
					fileContent.CoincidencePercent == 0 ?
						string.Empty : fileContent.CoincidencePercent.ToString(),
					fileContent.CoincidenceSource,
					fileContent.CoincidenceRowNumber == 0 ?
						string.Empty : fileContent.CoincidenceRowNumber.ToString(),
					fileContent.Comment
				};

				foreach (object value in values) {
					ICell cell = null;
					try { cell = row.GetCell(columnNumber); } catch (Exception) { }

					if (cell == null)
						cell = row.CreateCell(columnNumber);

					string valueToWrite = value == null ? string.Empty : value.ToString();

					if (double.TryParse(valueToWrite, out double result)) {
						cell.SetCellValue(result);
					} else if (DateTime.TryParse(valueToWrite, out DateTime date)) {
						cell.SetCellValue(date);
					} else {
						cell.SetCellValue(valueToWrite);
					}

					columnNumber++;
				}

				columnNumber = 0;
				rowNumber++;
			}

			UpdateProgress(ProgressCurrent, "Сохранение книги Excel");
			if (!SaveAndCloseIWorkbook(workbook, resultFile))
				return string.Empty;

			UpdateProgress(progressCurrent, "Выполнение пост-обработки");
			if (!PerformFinalProcessing(resultFile))
				UpdateProgress(ProgressCurrent, "!!!Во время выполнения возникли ошибки");

			return resultFile;
		}

		private static bool CreateNewIWorkbook(string resultFilePrefix, string templateFileName,
			out IWorkbook workbook, out ISheet sheet, out string resultFile, string sheetName) {
			workbook = null;
			sheet = null;
			resultFile = string.Empty;

			try {
				if (!GetTemplateFilePath(ref templateFileName))
					return false;

				string resultPath = GetResultFilePath(resultFilePrefix, templateFileName);

				using (FileStream stream = new FileStream(templateFileName, FileMode.Open, FileAccess.Read))
					workbook = new XSSFWorkbook(stream);

				if (string.IsNullOrEmpty(sheetName))
					sheetName = "Данные";

				sheet = workbook.GetSheet(sheetName);
				resultFile = resultPath;

				return true;
			} catch (Exception e) {
				UpdateProgress(ProgressCurrent, "!!!" + e.Message + Environment.NewLine + e.StackTrace);
				return false;
			}
		}

		private static bool GetTemplateFilePath(ref string templateFileName) {
			templateFileName = Path.Combine(AssemblyDirectory, templateFileName);

			if (!File.Exists(templateFileName)) {
				UpdateProgress(ProgressCurrent, "!!!Не удалось найти файл шаблона: " + templateFileName);
				return false;
			}

			return true;
		}

		private static string GetResultFilePath(string resultFilePrefix, string templateFileName, bool isPlainText = false) {
			string resultPath = Path.Combine(AssemblyDirectory, "Results");
			if (!Directory.Exists(resultPath))
				Directory.CreateDirectory(resultPath);

			foreach (char item in Path.GetInvalidFileNameChars())
				resultFilePrefix = resultFilePrefix.Replace(item, '-');

			string fileEnding = ".xlsx";
			if (isPlainText)
				fileEnding = ".txt";

			string resultFile = Path.Combine(resultPath, resultFilePrefix + " " + DateTime.Now.ToString("yyyyMMdd_HHmmss") + fileEnding);
			if (isPlainText)
				File.Copy(templateFileName, resultFile, true);

			return resultFile;
		}

		private static bool SaveAndCloseIWorkbook(IWorkbook workbook, string resultFile) {
			try {
				using (FileStream stream = new FileStream(resultFile, FileMode.Create, FileAccess.Write))
					workbook.Write(stream);

				workbook.Close();

				return true;
			} catch (Exception e) {
				UpdateProgress(ProgressCurrent, "!!!" + e.Message + Environment.NewLine + e.StackTrace);
				return false;
			}
		}

		private static bool PerformFinalProcessing(string resultFile) {
			UpdateProgress(ProgressCurrent, "Открытие книги с данными");
			if (!OpenWorkbook(resultFile, out Excel.Application xlApp, out Excel.Workbook wb, out Excel.Worksheet ws))
				return false;

			UpdateProgress(ProgressCurrent, "Применение форматирования");
			try {
				ws.Activate();
				ws.Range["A2:T2"].Select();
				xlApp.Selection.Copy();
				ws.Range["A3:T" + ws.UsedRange.Rows.Count].Select();
				xlApp.Selection.PasteSpecial(Excel.XlPasteType.xlPasteFormats);
			} catch (Exception e) {
				UpdateProgress(ProgressCurrent, e.Message + Environment.NewLine + e.StackTrace);
			}

			UpdateProgress(ProgressCurrent, "Добавление сводной таблицы");
			try {
				AddPivotTable(wb, ws, xlApp);
			} catch (Exception e) {
				UpdateProgress(ProgressCurrent, e.Message + Environment.NewLine + e.StackTrace);
			}

			UpdateProgress(ProgressCurrent, "Сохранение книги Excel");
			SaveAndCloseWorkbook(xlApp, wb, ws);

			return true;
		}

		private static void AddPivotTable(Excel.Workbook wb,
									Excel.Worksheet ws,
									Excel.Application xlApp) {
			ws.Cells[1, 1].Select();

			string pivotTableName = @"HalvaPivotTable";
			Excel.Worksheet wsPivote = wb.Sheets["Сводная таблица"];

			Excel.PivotCache pivotCache = wb.PivotCaches().Create(Excel.XlPivotTableSourceType.xlDatabase, ws.UsedRange, 6);
			Excel.PivotTable pivotTable = pivotCache.CreatePivotTable(wsPivote.Cells[1, 1], pivotTableName, true, 6);

			pivotTable = (Excel.PivotTable)wsPivote.PivotTables(pivotTableName);
			pivotTable.HasAutoFormat = false;

			pivotTable.PivotFields("Совпадение").Orientation = Excel.XlPivotFieldOrientation.xlRowField;
			pivotTable.PivotFields("Совпадение").Position = 1;

			pivotTable.PivotFields("Адрес Торговой точки").Orientation = Excel.XlPivotFieldOrientation.xlColumnField;
			pivotTable.PivotFields("Адрес Торговой точки").Position = 1;

			pivotTable.AddDataField(pivotTable.PivotFields("Дата обработки транзакции Банком"),
				"Количество совпадений", Excel.XlConsolidationFunction.xlCount);
			
			wsPivote.Activate();
			wsPivote.Columns["A:H"].Select();
			xlApp.Selection.ColumnWidth = 15;
			xlApp.Selection.WrapText = true;
			xlApp.Selection.VerticalAlignment = Excel.Constants.xlTop;
			wsPivote.Range["A1"].Select();
			pivotTable.DisplayFieldCaptions = false;

			wb.ShowPivotTableFieldList = false;
		}

		private static bool OpenWorkbook(string workbook,
								   out Excel.Application xlApp,
								   out Excel.Workbook wb,
								   out Excel.Worksheet ws,
								   string sheetName = "") {
			wb = null;
			ws = null;

			xlApp = new Excel.Application();

			if (xlApp == null) {
				UpdateProgress(ProgressCurrent, "!!!Не удалось открыть приложение Excel");
				return false;
			}

			xlApp.Visible = false;

			wb = xlApp.Workbooks.Open(workbook);

			if (wb == null) {
				UpdateProgress(ProgressCurrent, "!!!Не удалось открыть книгу " + workbook);
				return false;
			}

			if (string.IsNullOrEmpty(sheetName))
				sheetName = "Данные";

			ws = wb.Sheets[sheetName];

			if (ws == null) {
				UpdateProgress(ProgressCurrent, "!!!Не удалось открыть лист " + sheetName);
				return false;
			}

			return true;
		}

		private static void SaveAndCloseWorkbook(Excel.Application xlApp,
										   Excel.Workbook wb,
										   Excel.Worksheet ws) {
			if (ws != null) {
				Marshal.ReleaseComObject(ws);
				ws = null;
			}

			if (wb != null) {
				wb.Save();
				wb.Close(0);
				Marshal.ReleaseComObject(wb);
				wb = null;
			}

			if (xlApp != null) {
				xlApp.Quit();
				Marshal.ReleaseComObject(xlApp);
				xlApp = null;
			}

			GC.Collect();
			GC.WaitForPendingFinalizers();
		}
	}
}
