using System.Collections.Generic;
using System.IO;
using Bing.Excels.Abstractions.Elements;

namespace Bing.Excels.Core.Elements
{
    /// <summary>
    /// 工作簿
    /// </summary>
    public class Workbook:IWorkbook
    {
        /// <summary>
        /// 工作表列表
        /// </summary>
        public List<ISheet> Sheets { get; }


        public ISheet this[int sheetIndex] => throw new System.NotImplementedException();

        public string FileName { get; }
        public int SheetNumber => Sheets.Count;

        public Workbook()
        {
            Sheets = new List<ISheet>();
        }

        public void Load(string path)
        {
            throw new System.NotImplementedException();
        }

        public void Add(ISheet sheet)
        {
            throw new System.NotImplementedException();
        }

        public ISheet CreateSheet()
        {
            throw new System.NotImplementedException();
        }

        public ISheet CreateSheet(string name)
        {
            throw new System.NotImplementedException();
        }

        public ISheet GetSheet(string name)
        {
            throw new System.NotImplementedException();
        }

        public string SaveToFile(string path)
        {
            throw new System.NotImplementedException();
        }

        public void Write(Stream stream)
        {
            throw new System.NotImplementedException();
        }
    }
}
