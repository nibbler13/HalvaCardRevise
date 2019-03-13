using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HalvaCardRevise {
	class ExcelFileContentReader {
		private enum FileType {
			SberbankOld,
			SberbankNew,
			Vtb,
			Soyuz,
			Unknown
		}
		public static string ReadExcelFileContent(ItemFileInfo itemFileInfo) {
			IWorkbook workbook;
			using (FileStream file = new FileStream(itemFileInfo.FullPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)) {
				workbook = new XSSFWorkbook(file);
			}

			int startRow = 0;
			int cellRNN = 0;
			int cellAuthorizationCode = 0;
			int cellOperationAmount = 0;
			int cellCardNumber = 0;
			int cellCommittingDate = 0;
			int cellCommittingTime = 0;

			FileType type = FileType.Unknown;
			ISheet sheet = workbook.GetSheetAt(0);

			try {
				IRow row = sheet.GetRow(0);
				string firstValue = GetCellValue(row.GetCell(0)).ToLower();
				if (firstValue.StartsWith("номер_терминала"))
					type = FileType.Soyuz;
				else if (firstValue.StartsWith("отчет о возмещении денежных средств предприятию")) {

					string secondValue = GetCellValue(sheet.GetRow(1).GetCell(0)).ToLower();
					if (secondValue.Equals("номер транзакции"))
						type = FileType.SberbankOld;
					else
						type = FileType.SberbankNew;
				}
				else if (firstValue.StartsWith("отчёт по обработанным операциям за период"))
					type = FileType.Vtb;
			} catch (Exception) { }


			switch (type) {
				case FileType.SberbankOld:
					startRow = 2;
					cellRNN = 0;
					cellAuthorizationCode = 21;
					cellOperationAmount = 10;
					cellCardNumber = 20;
					cellCommittingDate = 8;
					cellCommittingTime = 8;
					break;
				case FileType.SberbankNew:
					startRow = 3;
					cellRNN = -1;
					cellAuthorizationCode = 9;
					cellOperationAmount = 5;
					cellCardNumber = 8;
					cellCommittingDate = 2;
					cellCommittingTime = 2;
					break;
				case FileType.Vtb:
					startRow = 12;
					cellRNN = 13;
					cellAuthorizationCode = 6;
					cellOperationAmount = 9;
					cellCardNumber = 1;
					cellCommittingDate = 4;
					cellCommittingTime = 5;
					break;
				case FileType.Soyuz:
					startRow = 1;
					cellRNN = 5;
					cellAuthorizationCode = 6;
					cellOperationAmount = 7;
					cellCardNumber = 4;
					cellCommittingDate = 2;
					cellCommittingTime = 3;
					break;
				default:
					return type.ToString();
			}

			for (int rowCount = startRow; rowCount <= sheet.LastRowNum; rowCount++) {
				try {
					IRow row = sheet.GetRow(rowCount);

					if (row == null)  //null is when the row only contains empty cells 
						continue;

					string rnn = string.Empty;
					
					if (cellRNN != -1) {
						rnn = GetCellValue(row.GetCell(cellRNN));

						if (string.IsNullOrEmpty(rnn) ||
							string.IsNullOrWhiteSpace(rnn) ||
							!long.TryParse(rnn, out _))
							continue;
					}

					string authorizationCode = GetCellValue(row.GetCell(cellAuthorizationCode));
					string operationAmount = GetCellValue(row.GetCell(cellOperationAmount));
					string cardNumber = GetCellValue(row.GetCell(cellCardNumber));
					string committingDate = GetCellValue(row.GetCell(cellCommittingDate));
					string committingTime = GetCellValue(row.GetCell(cellCommittingTime));

					if (type == FileType.SberbankNew)
						committingDate = GetCellValue(row.GetCell(cellCommittingDate), true);

					if (type == FileType.SberbankOld ||
						type == FileType.SberbankNew) {
						string[] splittedDateTime = committingDate.Split(' ');
						if (splittedDateTime.Length == 2) {
							committingDate = splittedDateTime[0];
							committingTime = splittedDateTime[1];

							if (committingTime.Length == 7)
								committingTime = "0" + committingTime;
						}
					}

					//if (type == FileType.Soyuz) {
					//	if (committingDate.Length == 8)
					//		committingDate = 
					//			committingDate.Substring(6, 2) + "." +
					//			committingDate.Substring(4, 2) + "." +
					//			committingDate.Substring(0, 4);

					//	if (committingTime.Length == 6)
					//		committingTime =
					//			committingTime.Substring(0, 2) + ":" +
					//			committingTime.Substring(2, 2) + ":" +
					//			committingTime.Substring(4, 2);
					//}

					FileContent fileContent = new FileContent {
						UniqueOperationNumberRNN = rnn,
						OperationAmount = operationAmount,
						AuthorizationCode = authorizationCode,
						CardNumber = cardNumber,
						TransactionCommittingDate = committingDate,
						TransactionCommitingTime = committingTime,
						CoincidenceRowNumber = rowCount + 1
					};

					itemFileInfo.FileContents.Add(fileContent);
				} catch (Exception e) {
					Console.WriteLine(e.Message + Environment.NewLine + e.StackTrace);
				}
				
			}

			workbook.Close();
			return type.ToString();
		}

		private static string GetCellValue(ICell cell, bool isDate = false) {
			object retValue = string.Empty;

			if (cell == null)
				return retValue.ToString();

			if (isDate)
				return cell.DateCellValue.ToShortDateString() + " " + cell.DateCellValue.ToLongTimeString();

			switch (cell.CellType) {
				case CellType.Unknown:
					retValue = cell.StringCellValue;
					break;
				case CellType.Numeric:
					retValue = cell.NumericCellValue;
					break;
				case CellType.String:
					retValue = cell.StringCellValue;
					break;
				case CellType.Formula:
					retValue = cell.StringCellValue;
					break;
				case CellType.Blank:
					retValue = string.Empty;
					break;
				case CellType.Boolean:
					retValue = cell.BooleanCellValue;
					break;
				case CellType.Error:
					retValue = cell.ErrorCellValue;
					break;
				default:
					break;
			}

			return retValue.ToString();
		}
	}
}
