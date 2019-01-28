using System.Collections.Generic;
using System.Linq;
using Bing.Excels.Abstractions.Elements;
using Bing.Excels.Abstractions.Styles;
using Bing.Excels.Internal;

namespace Bing.Excels.Core.Elements
{
    /// <summary>
    /// 单元行
    /// </summary>
    public class Row:IRow
    {
        /// <summary>
        /// 索引管理器
        /// </summary>
        private readonly IndexManager _indexManager;

        /// <summary>
        /// 单元格列表
        /// </summary>
        public List<ICell> Cells { get; set; }

        /// <summary>
        /// 单元格
        /// </summary>
        /// <param name="columnIndex">列索引</param>
        /// <returns></returns>
        public ICell this[int columnIndex] => Cells[columnIndex];

        /// <summary>
        /// 行索引
        /// </summary>
        public int RowIndex { get; }

        /// <summary>
        /// 列总数
        /// </summary>
        public int ColumnNumber => Cells.Sum(x => x.Range.ColumnSpan);

        /// <summary>
        /// 行高
        /// </summary>
        public int Height { get; set; }

        /// <summary>
        /// 工作表
        /// </summary>
        public ISheet Sheet { get; }

        /// <summary>
        /// 单元行样式
        /// </summary>
        public ICellStyle RowStyle { get; set; }

        /// <summary>
        /// 初始化一个<see cref="Row"/>类型的实例
        /// </summary>
        /// <param name="sheet">工作表</param>
        /// <param name="rowIndex">行索引</param>
        public Row(ISheet sheet, int rowIndex)
        {
            Sheet = sheet;
            RowIndex = rowIndex;
            Cells = new List<ICell>();
            _indexManager = new IndexManager();
        }

        /// <summary>
        /// 添加单元格
        /// </summary>
        /// <param name="value">值</param>
        /// <param name="columnSpan">列跨度</param>
        /// <param name="rowSpan">行跨度</param>
        public void Add(object value, int columnSpan = 1, int rowSpan = 1)
        {
            Add(new Cell(value, columnSpan, rowSpan));
        }

        /// <summary>
        /// 添加单元格
        /// </summary>
        /// <param name="cell">单元格</param>
        public void Add(ICell cell)
        {
            if (cell == null)
            {
                return;
            }

            cell.Row = this;
            SetColumnIndex(cell);
            Cells.Add(cell);
        }

        /// <summary>
        /// 设置列索引
        /// </summary>
        /// <param name="cell">单元格</param>
        private void SetColumnIndex(ICell cell)
        {
            if (cell.ColumnIndex > 0)
            {
                _indexManager.AddIndex(cell.ColumnIndex, cell.Range.RowSpan);
                return;
            }

            cell.ColumnIndex = _indexManager.GetIndex(cell.Range.ColumnSpan);
        }

        /// <summary>
        /// 创建单元格
        /// </summary>
        /// <returns></returns>
        public ICell CreateCell()
        {
            return new Cell("");
        }

        /// <summary>
        /// 清空内容
        /// </summary>
        /// <returns></returns>
        public IRow ClearContent()
        {
            foreach (var cell in Cells)
            {
                cell.SetValue(string.Empty);
            }

            return this;
        }
    }
}
