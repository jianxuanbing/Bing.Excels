using System.Collections.Generic;

namespace Bing.Excels.Abstractions.Elements
{
    /// <summary>
    /// 单元范围
    /// </summary>
    public interface IRange
    {
        /// <summary>
        /// 单元行
        /// </summary>
        /// <param name="rowIndex">行索引</param>
        /// <returns></returns>
        IRow this[int rowIndex] { get; }

        /// <summary>
        /// 列数
        /// </summary>
        int ColumnNumber { get; }

        /// <summary>
        /// 行数
        /// </summary>
        int RowNumber { get; }        

        /// <summary>
        /// 工作表
        /// </summary>
        ISheet Sheet { get; }

        /// <summary>
        /// 获取单元行
        /// </summary>
        /// <param name="rowIndex">行索引，对应Excel表格行号</param>
        /// <returns></returns>
        IRow GetRow(int rowIndex);

        /// <summary>
        /// 获取单元行列表
        /// </summary>
        /// <returns></returns>
        List<IRow> GetRows();

        /// <summary>
        /// 添加单元行
        /// </summary>
        /// <param name="rowIndex">行索引</param>
        /// <param name="cells">单元格集合</param>
        void AddRow(int rowIndex, IEnumerable<ICell> cells);

        /// <summary>
        /// 清空单元行
        /// </summary>
        void Clear();
    }
}
