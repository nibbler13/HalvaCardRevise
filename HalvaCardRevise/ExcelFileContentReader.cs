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
		public static void ReadExcelFileContent(ItemFileInfo itemFileInfo) {
			IWorkbook workbook;
			using (FileStream file = new FileStream(itemFileInfo.FullPath, FileMode.Open, FileAccess.Read)) {
				workbook = new XSSFWorkbook(file);
			}

			int startRow = 0;
			int cellRNN = 0;
			int cellAuthorizationCode = 0;
			int cellOperationAmount = 0;
			int cellCardNumber = 0;
			int cellCommittingDate = 0;
			int cellCommittingTime = 0;

			switch (itemFileInfo.Type) {
				case ItemFileInfo.FileType.Sberbank:
					startRow = 2;
					cellRNN = 0;
					cellAuthorizationCode = 21;
					cellOperationAmount = 10;
					cellCardNumber = 20;
					cellCommittingDate = 8;
					cellCommittingTime = 8;
					break;
				case ItemFileInfo.FileType.Vtb:
					startRow = 12;
					cellRNN = 13;
					cellAuthorizationCode = 6;
					cellOperationAmount = 9;
					cellCardNumber = 1;
					cellCommittingDate = 4;
					cellCommittingTime = 5;
					break;
				default:
					break;
			}

			ISheet sheet = workbook.GetSheetAt(0);
			for (int rowCount = startRow; rowCount <= sheet.LastRowNum; rowCount++) {
				try {
					IRow row = sheet.GetRow(rowCount);

					if (row == null)  //null is when the row only contains empty cells 
						continue;

					string rnn = GetCellValue(row.GetCell(cellRNN));

					if (string.IsNullOrEmpty(rnn) ||
						string.IsNullOrWhiteSpace(rnn) ||
						!long.TryParse(rnn, out _))
						continue;

					string authorizationCode = GetCellValue(row.GetCell(cellAuthorizationCode));
					string operationAmount = GetCellValue(row.GetCell(cellOperationAmount));
					string cardNumber = GetCellValue(row.GetCell(cellCardNumber));
					string committingDate = GetCellValue(row.GetCell(cellCommittingDate));
					string committingTime = GetCellValue(row.GetCell(cellCommittingTime));

					if (itemFileInfo.Type == ItemFileInfo.FileType.Sberbank) {
						string[] splittedDateTime = committingDate.Split(' ');
						if (splittedDateTime.Length == 2) {
							committingDate = splittedDateTime[0];
							committingTime = splittedDateTime[1];
						}
					}

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
		}

		private static string GetCellValue(ICell cell) {
			object retValue = string.Empty;

			if (cell == null)
				return retValue.ToString();

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
