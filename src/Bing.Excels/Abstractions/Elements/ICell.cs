using Bing.Excels.Abstractions.Styles;
using Bing.Excels.Core.Styles;

namespace Bing.Excels.Abstractions.Elements
{
    /// <summary>
    /// 单元格
    /// </summary>
    public interface ICell
    {
        /// <summary>
        /// 单元格值
        /// </summary>
        object Value { get; set; }

        /// <summary>
        /// 单元格类型
        /// </summary>
        CellType CellType { get; }

        /// <summary>
        /// 行索引
        /// </summary>
        int RowIndex { get; }

        /// <summary>
        /// 列索引
        /// </summary>
        int ColumnIndex { get; set; }

        /// <summary>
        /// 行跨度
        /// </summary>
        int RowSpan { get; set; }

        /// <summary>
        /// 列跨度
        /// </summary>
        int ColumnSpan { get; set; }

        /// <summary>
        /// 单元行
        /// </summary>
        IRow Row { get; set; }

        /// <summary>
        /// 工作表
        /// </summary>
        ISheet Sheet { get; }

        /// <summary>
        /// 单元格样式
        /// </summary>
        ICellStyle CellStyle { get; set; }

        /// <summary>
        /// 单元格范围
        /// </summary>
        ICellRangeAddress Range { get; }

        /// <summary>
        /// 是否合并单元格
        /// </summary>
        bool IsMergedCell { get; }

        /// <summary>
        /// 设置单元格的值
        /// </summary>
        /// <param name="value">值</param>
        void SetValue(object value);
    }
}
