using System.Collections.Generic;
using System.Linq;
using Bing.Excels.Abstractions.Elements;

namespace Bing.Excels.Core.Elements
{
    /// <summary>
    /// 单元范围
    /// </summary>
    public class Range:IRange
    {
        /// <summary>
        /// 单元行列表
        /// </summary>
        private readonly List<IRow> _rows;

        /// <summary>
        /// 起始行索引
        /// </summary>
        private readonly int _startIndex;

        /// <summary>
        /// 单元行
        /// </summary>
        /// <param name="rowIndex">行索引</param>
        /// <returns></returns>
        public IRow this[int rowIndex] => _rows[rowIndex];

        /// <summary>
        /// 列数
        /// </summary>
        public int ColumnNumber => _rows.Count > 0 ? _rows.Max(x => x.ColumnNumber) : 0;

        /// <summary>
        /// 行数
        /// </summary>
        public int RowNumber => _rows.Count;

        /// <summary>
        /// 工作表
        /// </summary>
        public ISheet Sheet { get; }

        /// <summary>
        /// 初始化一个<see cref="Range"/>类型的实例
        /// </summary>
        /// <param name="sheet">工作表</param>
        /// <param name="startIndex">起始行索引</param>
        public Range(ISheet sheet, int startIndex = 0)
        {
            Sheet = sheet;
            _rows = new List<IRow>();
            _startIndex = startIndex;
        }

        /// <summary>
        /// 获取单元行
        /// </summary>
        /// <param name="rowIndex">行索引，对应Excel表格行号</param>
        /// <returns></returns>
        public IRow GetRow(int rowIndex)
        {
            var realIndex = rowIndex - _startIndex;
            if (realIndex < 0)
            {
                return null;
            }

            if (realIndex > _rows.Count - 1)
            {
                return null;
            }

            return _rows[realIndex];
        }

        /// <summary>
        /// 获取单元行列表
        /// </summary>
        /// <returns></returns>
        public List<IRow> GetRows() => _rows;

        /// <summary>
        /// 添加单元行
        /// </summary>
        /// <param name="rowIndex">行索引</param>
        /// <param name="cells">单元格集合</param>
        public void AddRow(int rowIndex, IEnumerable<ICell> cells)
        {
            if (cells == null)
            {
                return;
            }

            var row = CreateRow(rowIndex);
            foreach (var cell in cells)
            {
                AddCell(row, cell, rowIndex);
            }
        }

        /// <summary>
        /// 创建单元行
        /// </summary>
        /// <param name="index">行索引</param>
        /// <returns></returns>
        private IRow CreateRow(int index)
        {
            var row = GetRow(index);
            if (row != null)
            {
                return row;
            }

            row = new Row(Sheet,index);
            _rows.Add(row);
            return row;
        }

        /// <summary>
        /// 添加单元格
        /// </summary>
        /// <param name="row">单元行</param>
        /// <param name="cell">单元格</param>
        /// <param name="rowIndex">行索引</param>
        private void AddCell(IRow row, ICell cell, int rowIndex)
        {
            row.Add(cell);
            if (cell.RowSpan <= 1)
            {
                return;
            }

            for (int i = 1; i < cell.RowSpan; i++)
            {
                AddPlaceholderCell(cell, rowIndex + i);
            }
        }

        /// <summary>
        /// 为下方受RowSpan影响的单元行添加占位空格
        /// </summary>
        /// <param name="cell">单元格</param>
        /// <param name="rowIndex">行索引</param>
        private void AddPlaceholderCell(ICell cell, int rowIndex)
        {
            var row = CreateRow(rowIndex);
            row.Add(new NullCell() {ColumnIndex = cell.ColumnIndex, ColumnSpan = cell.ColumnSpan});
        }

        /// <summary>
        /// 清空单元行
        /// </summary>
        public void Clear()
        {
            _rows.Clear();
        }
    }
}
