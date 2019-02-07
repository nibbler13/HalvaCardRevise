using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using CsvHelper;

namespace HalvaCardRevise {
	class CsvFileContentReader {
		public static void ReadCsvFileContent(ItemFileInfo itemFileInfo) {
			using (StreamReader reader = new StreamReader(itemFileInfo.FullPath, Encoding.GetEncoding("windows-1251"))) {
				using (CsvReader csvReader = new CsvReader(reader)) {
					itemFileInfo.FileContents = csvReader.GetRecords<FileContent>().ToList();
				}
			}
		}
	}
}
