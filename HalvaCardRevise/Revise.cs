using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace HalvaCardRevise {
	class Revise {
		private readonly ItemFileInfo halvaCardInfo;
		private readonly List<ItemFileInfo> terminalFiles;
		private readonly BackgroundWorker backgroundWorker;
		private double currentProgress;

		public Revise(ItemFileInfo halvaCardInfo, List<ItemFileInfo> terminalFiles, BackgroundWorker backgroundWorker) {
			this.halvaCardInfo = halvaCardInfo ?? throw new ArgumentNullException("halvaCardInfo");
			this.terminalFiles = terminalFiles ?? throw new ArgumentNullException("terminalFiles");
			this.backgroundWorker = backgroundWorker ?? throw new ArgumentNullException("backgroundWorker");
			currentProgress = 0;
		}

		public void DoRevise() {
			UpdateProgress(currentProgress, "---Чтение содержимого выбранных файлов");

			double progressPerFile = 50.0d / (terminalFiles.Count + 1);

			UpdateProgress(currentProgress, "Чтениые файла: " + halvaCardInfo.FileName);
			CsvFileContentReader.ReadCsvFileContent(halvaCardInfo);
			if (halvaCardInfo.FileContents.Count == 0) {
				UpdateProgress(currentProgress, "!!!Файл с отчетом по картам 'Халва' не содержит данных");
				return;
			}

			currentProgress += progressPerFile;
			UpdateProgress(currentProgress, "Считано строк: " + halvaCardInfo.FileContents.Count);

			long totalLinesReaded = 0;
			foreach (ItemFileInfo fileInfo in terminalFiles) {
				UpdateProgress(currentProgress, "Чтение файла: " + fileInfo.FileName);
				UpdateProgress(currentProgress, "Тип файла: " + ExcelFileContentReader.ReadExcelFileContent(fileInfo));

				currentProgress += progressPerFile;
				UpdateProgress(currentProgress, "Считано строк: " + fileInfo.FileContents.Count);
				totalLinesReaded += fileInfo.FileContents.Count;
			}

			if (totalLinesReaded == 0) {
				UpdateProgress(currentProgress, "!!!В выбранных файлах с отчетами по терминалам " +
					"не удалось считать ни одной строки");
				return;
			}

			UpdateProgress(currentProgress, "---Сверка считанных данных");
			int fullCoincidence = 6;
			double progressPerContent = 40.0d / halvaCardInfo.FileContents.Count;
			foreach (FileContent halvaContent in halvaCardInfo.FileContents) {
				currentProgress += progressPerContent;
				UpdateProgress(currentProgress);

				int coincidence = 0;
				foreach (ItemFileInfo terminalFile in terminalFiles) {
					if (!string.IsNullOrEmpty(halvaContent.UniqueOperationNumberRNN)) {
						FileContent terminalContent = terminalFile.FileContents.Where(
							x => x.UniqueOperationNumberRNN.Equals(halvaContent.UniqueOperationNumberRNN)).FirstOrDefault();

						if (terminalContent != null) {
							halvaContent.CoincidenceRowNumber = terminalContent.CoincidenceRowNumber;

							if (halvaContent.AuthorizationCode.Equals(terminalContent.AuthorizationCode))
								coincidence++;
							else
								halvaContent.Comment += "Код авторизации; ";

							CheckCoincidence(halvaContent, terminalContent, out int coin, out string comm);
							coincidence += coin;
							halvaContent.Comment += comm;
						}
					}

					if (coincidence == 0) {
						List<FileContent> searchResults = terminalFile.FileContents.Where(
							x => x.AuthorizationCode.Equals(halvaContent.AuthorizationCode)).ToList();

						if (searchResults.Count == 0)
							continue;

						foreach (FileContent terminalContent in searchResults) {
							if (!CheckCoincidence(halvaContent, terminalContent, out int coin, out string comm))
								continue;

							halvaContent.Comment += "Уникальный номер операции (RNN); ";
							halvaContent.CoincidenceRowNumber = terminalContent.CoincidenceRowNumber;

							coincidence += coin;
							halvaContent.Comment += comm;
						}
					}

					if (coincidence == 0)
						continue;

					coincidence++;
					halvaContent.CoincidenceSource = terminalFile.FullPath;
					break;
				}

				halvaContent.CoincidencePercent = (double)coincidence / (double)fullCoincidence;
				if (coincidence == fullCoincidence)
					halvaContent.CoincidenceType = "Полное";
				else if (coincidence == 0)
					halvaContent.CoincidenceType = "Не найдено";
				else
					halvaContent.CoincidenceType = "Частичное";
			}

			UpdateProgress(currentProgress, "---Выгрузка данных в Excel");
			string resultFile = ExcelFileContentWriter.WriteFileInfoToExcel(halvaCardInfo, UpdateProgress, currentProgress);

			if (!string.IsNullOrEmpty(resultFile)) {
				UpdateProgress(currentProgress, "Данные сохранены в файл: " + resultFile);
				Process.Start(resultFile);
			}

			UpdateProgress(100, "===Завершено");
		}

		private void UpdateProgress(double progressValue, string progressInfo = "") {
			backgroundWorker.ReportProgress((int)progressValue, progressInfo);
		}

		private bool CheckCoincidence(FileContent halvaContent, FileContent terminalContent, out int coincidence, out string comment) {
			coincidence = 0;
			comment = string.Empty;
			bool errorDate = false;
			bool errorTime = false;

			if (halvaContent.TransactionCommittingDate.Equals(terminalContent.TransactionCommittingDate))
				coincidence++;
			else {
				errorDate = true;
				comment += "Дата совершения транзакции; ";
			}

			int halvaTimeLength = halvaContent.TransactionCommitingTime.Length;
			int resultTimeLength = terminalContent.TransactionCommitingTime.Length;
			if (halvaTimeLength >= 5 && resultTimeLength >= 5) {
				string halvaTime = halvaContent.TransactionCommitingTime.Substring(0, 5);
				string resultTime = terminalContent.TransactionCommitingTime.Substring(0, 5);
				resultTime = resultTime.Replace('.', ':'); //For Sberbank time format

				if (TimeSpan.TryParse(halvaTime, out TimeSpan halvaTimeSpan) &&
					TimeSpan.TryParse(resultTime, out TimeSpan resultTimeSpan)) {
					if (halvaContent.City.ToLower().EndsWith("уфа"))
						resultTimeSpan -= TimeSpan.FromHours(2);

					if (Math.Abs((resultTimeSpan - halvaTimeSpan).TotalMinutes) <= 2)
						coincidence++;
					else {
						errorTime = true;
						comment += "Время совершения транзакции; ";
					}
				} else {
					if (halvaTime.Equals(resultTime))
						coincidence++;
					else {
						errorTime = true;
						comment += "Время совершения транзакции; ";
					}
				}
			} else {
				if (halvaContent.TransactionCommitingTime.Equals(terminalContent.TransactionCommitingTime))
					coincidence++;
				else {
					errorTime = true;
					comment += "Время совершения транзакции; ";
				}
			}

			if (errorDate && errorTime)
				return false;

			if (double.TryParse(halvaContent.OperationAmount, out double halvaOperationAmountParsed) &&
				double.TryParse(terminalContent.OperationAmount, out double resultOperationAmountParsed)) {
				if (halvaOperationAmountParsed == resultOperationAmountParsed)
					coincidence++;
				else
					comment += "Сумма операции; ";
			} else {
				if (halvaContent.OperationAmount.Equals(terminalContent.OperationAmount))
					coincidence++;
				else
					comment += "Сумма операции; ";
			}

			int halvaCardLength = halvaContent.CardNumber.Length;
			int resultCardLength = terminalContent.CardNumber.Length;
			if (halvaCardLength >= 4 && resultCardLength >= 4) {
				string halvaCardLast4 = halvaContent.CardNumber.Substring(halvaCardLength - 4, 4);
				string resultCardLast4 = terminalContent.CardNumber.Substring(halvaCardLength - 4, 4);

				if (halvaCardLast4.Equals(resultCardLast4))
					coincidence++;
				else
					comment += "Номер карты; ";
			} else {
				if (halvaContent.CardNumber.Equals(terminalContent.CardNumber))
					coincidence++;
				else
					comment += "Номер карты; ";
			}

			return true;
		}
	}
}
