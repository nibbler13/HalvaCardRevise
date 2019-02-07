﻿using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HalvaCardRevise {
	class ItemFileInfo {
		public enum FileType {
			Halva,
			Sberbank,
			Vtb
		}

		public string FileName {
			get { return Path.GetFileName(FullPath); }
		}
		public string FullPath { get; private set; }
		public bool IsSelected { get; set; }
		public FileType Type { get; private set; }

		public List<FileContent> FileContents { get; set; }


		public ItemFileInfo(string fullPath, FileType type) {
			FullPath = fullPath;
			Type = type;
			FileContents = new List<FileContent>();
		}
		
		public override string ToString() {
			return FileName;
		}
	}
}
