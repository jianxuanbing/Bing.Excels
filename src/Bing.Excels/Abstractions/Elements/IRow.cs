using System.Collections.Generic;
using Bing.Excels.Abstractions.Styles;

namespace Bing.Excels.Abstractions.Elements
{
    /// <summary>
    /// 单元行
    /// </summary>
    public interface IRow
    {
        /// <summary>
        /// 单元格列表
        /// </summary>
        List<ICell> Cells { get; set; }

        /// <summary>
        /// 单元格
        /// </summary>
        /// <param name="columnIndex">列索引</param>
        /// <returns></returns>
        ICell this[int columnIndex] { get; }

        /// <summary>
        /// 行索引
        /// </summary>
        int RowIndex { get; }

        /// <summary>
        /// 列总数
        /// </summary>
        int ColumnNumber { get; }

        /// <summary>
        /// 行高
        /// </summary>
        int Height { get; set; }

        /// <summary>
        /// 工作表
        /// </summary>
        ISheet Sheet { get; }

        /// <summary>
        /// 单元行样式
        /// </summary>
        ICellStyle RowStyle { get; set; }

        /// <summary>
        /// 添加单元格
        /// </summary>
        /// <param name="value">值</param>
        /// <param name="columnSpan">列跨度</param>
        /// <param name="rowSpan">行跨度</param>
        void Add(object value, int columnSpan = 1, int rowSpan = 1);

        /// <summary>
        /// 添加单元格
        /// </summary>
        /// <param name="cell">单元格</param>
        void Add(ICell cell);

        /// <summary>
        /// 创建单元格
        /// </summary>
        /// <returns></returns>
        ICell CreateCell();

        /// <summary>
        /// 清空内容
        /// </summary>
        /// <returns></returns>
        IRow ClearContent();
    }
}
