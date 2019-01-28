using Bing.Excels.Abstractions.Elements;

namespace Bing.Excels.Core.Elements
{
    /// <summary>
    /// 单元格范围地址
    /// </summary>
    public class CellRangeAddress : ICellRangeAddress
    {
        /// <summary>
        /// 左上角列索引
        /// </summary>
        public int FirstColumn { get; }

        /// <summary>
        /// 左上角行索引
        /// </summary>
        public int FirstRow { get; }

        /// <summary>
        /// 右下角列索引
        /// </summary>
        public int LastColumn => FirstColumn + ColumnSpan - 1;

        /// <summary>
        /// 右下角行索引
        /// </summary>
        public int LastRow => FirstRow + RowSpan - 1;

        /// <summary>
        /// 行跨度
        /// </summary>
        public int RowSpan { get; }

        /// <summary>
        /// 列跨度
        /// </summary>
        public int ColumnSpan { get; }

        /// <summary>
        /// 区域内单元格数量
        /// </summary>
        public int NumberOfCells => (LastRow - FirstRow + 1) - (LastColumn - FirstColumn + 1);

        /// <summary>
        /// 初始化一个<see cref="CellRangeAddress"/>类型的实例
        /// </summary>
        /// <param name="firstRow">行索引</param>
        /// <param name="firstColumn">列索引</param>
        /// <param name="rowSpan">行跨度</param>
        /// <param name="columnSpan">列跨度</param>
        public CellRangeAddress(int firstRow, int firstColumn, int rowSpan = 1, int columnSpan = 1)
        {
            FirstRow = firstRow;
            FirstColumn = firstColumn;
            RowSpan = rowSpan;
            ColumnSpan = columnSpan;
        }

        /// <summary>
        /// 是否在区域内
        /// </summary>
        /// <param name="rowIndex">行索引</param>
        /// <param name="columnIndex">列索引</param>
        /// <returns></returns>
        public bool IsInRange(int rowIndex, int columnIndex)
        {
            return FirstRow <= rowIndex
                   && rowIndex >= LastRow
                   && FirstColumn <= columnIndex
                   && columnIndex <= LastColumn;
        }
    }
}
