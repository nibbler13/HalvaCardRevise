using CsvHelper.Configuration;
using CsvHelper.Configuration.Attributes;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HalvaCardRevise {
	class FileContent {
		[Name("Дата обработки транзакции Банком")]
		public string TransactionProcessingDate { get; set; }

		[Name("Дата совершения транзации")]
		public string TransactionCommittingDate { get; set; }

		[Name("Время совершения транзакции")]
		public string TransactionCommitingTime { get; set; }

		[Name("Наименование Заказчика (юр.лица)")]
		public string CustomerName { get; set; }

		[Name("ИНН Заказчика")]
		public string CustomerINN { get; set; }

		[Name("Название Торговой точки/розничной сети Заказчика")]
		public string CustomerStoreName { get; set; }

		[Name("Адрес Торговой точки")]
		public string CustomerStoreAddress { get; set; }

		[Name("Идентификатор платежного терминала")]
		public string TerminalIdentificator { get; set; }

		[Name("Номер банковской карты")]
		public string CardNumber { get; set; }

		[Name("Сумма операции, руб.")]
		public string OperationAmount { get; set; }

		[Name("Общий размер вознаграждения, %")]
		public string TotalReward { get; set; }

		[Name("Вознаграждение в адрес Компании, руб.")]
		public string CompanyReward { get; set; }

		[Name("Город")]
		public string City { get; set; }

		[Name("Код авторизации")]
		public string AuthorizationCode { get; set; }

		[Name("Уникальный номер операции (RNN)")]
		public string UniqueOperationNumberRNN { get; set; }

		[Ignore]
		public string CoincidenceType { get; set; }

		[Ignore]
		public double CoincidencePercent { get; set; }

		[Ignore]
		public string CoincidenceSource { get; set; }

		[Ignore]
		public int CoincidenceRowNumber { get; set; }

		[Ignore]
		public string Comment { get; set; }
	}
}
