namespace Bing.Excels.Abstractions.Elements
{
    /// <summary>
    /// 单元格范围地址
    /// </summary>
    public interface ICellRangeAddress
    {
        /// <summary>
        /// 左上角列索引
        /// </summary>
        int FirstColumn { get; }

        /// <summary>
        /// 左上角行索引
        /// </summary>
        int FirstRow { get; }

        /// <summary>
        /// 右下角列索引
        /// </summary>
        int LastColumn { get; }

        /// <summary>
        /// 右下角行索引
        /// </summary>
        int LastRow { get; }

        /// <summary>
        /// 行跨度
        /// </summary>
        int RowSpan { get; }

        /// <summary>
        /// 列跨度
        /// </summary>
        int ColumnSpan { get; }

        /// <summary>
        /// 区域内单元格数量
        /// </summary>
        int NumberOfCells { get; }

        /// <summary>
        /// 是否在区域内
        /// </summary>
        /// <param name="rowIndex">行索引</param>
        /// <param name="columnIndex">列索引</param>
        /// <returns></returns>
        bool IsInRange(int rowIndex, int columnIndex);
    }
}
