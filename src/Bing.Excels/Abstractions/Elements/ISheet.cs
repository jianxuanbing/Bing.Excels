using System.Collections.Generic;
using Bing.Excels.Abstractions.Styles;

namespace Bing.Excels.Abstractions.Elements
{
    /// <summary>
    /// 工作表
    /// </summary>
    public interface ISheet
    {
        /// <summary>
        /// 总标题
        /// </summary>
        string Title { get; set; }

        /// <summary>
        /// 工作表名称
        /// </summary>
        string SheetName { get; set; }

        /// <summary>
        /// 工作表索引
        /// </summary>
        int SheetIndex { get; }

        /// <summary>
        /// 工作簿
        /// </summary>
        IWorkbook Workbook { get; }

        /// <summary>
        /// 是否当前选中工作表
        /// </summary>
        bool IsSelected { get; set; }

        /// <summary>
        /// 表头行数
        /// </summary>
        int HeadRowNumber { get; }

        /// <summary>
        /// 正文行数
        /// </summary>
        int BodyRowNumber { get; }

        /// <summary>
        /// 页脚行数
        /// </summary>
        int FootRowNumber { get; }

        /// <summary>
        /// 总行数
        /// </summary>
        int RowNumber { get; }

        /// <summary>
        /// 总列数
        /// </summary>
        int ColumnNumber { get; }

        /// <summary>
        /// 默认列宽度
        /// </summary>
        int DefaultColumnWidth { get; set; }

        /// <summary>
        /// 默认行高度
        /// </summary>
        int DefaultRowHeight { get; set; }

        /// <summary>
        /// 设置是否选择了工作表
        /// </summary>
        /// <param name="value">是否选择工作表</param>
        void SetActive(bool value);

        /// <summary>
        /// 获取表头
        /// </summary>
        /// <returns></returns>
        List<IRow> GetHeader();

        /// <summary>
        /// 获取表格正文
        /// </summary>
        /// <returns></returns>
        List<IRow> GetBody();

        /// <summary>
        /// 获取页脚
        /// </summary>
        /// <returns></returns>
        List<IRow> GetFooter();

        /// <summary>
        /// 添加表头
        /// </summary>
        /// <param name="titles">标题</param>
        void AddHeadRow(params string[] titles);

        /// <summary>
        /// 添加表头
        /// </summary>
        /// <param name="cells">单元格</param>
        void AddHeadRow(params ICell[] cells);

        /// <summary>
        /// 添加正文
        /// </summary>
        /// <param name="cellValues">值</param>
        void AddBodyRow(params object[] cellValues);

        /// <summary>
        /// 添加正文
        /// </summary>
        /// <param name="cells">单元格集合</param>
        void AddBodyRow(IEnumerable<ICell> cells);

        /// <summary>
        /// 添加页脚
        /// </summary>
        /// <param name="cellValues">值</param>
        void AddFootRow(params string[] cellValues);

        /// <summary>
        /// 添加页脚
        /// </summary>
        /// <param name="cells">单元格集合</param>
        void AddFootRow(params ICell[] cells);

        /// <summary>
        /// 设置列宽
        /// </summary>
        /// <param name="columIndex">列索引</param>
        /// <param name="width">宽度</param>
        void SetColumnWidth(int columIndex, int width);

        /// <summary>
        /// 获取列宽
        /// </summary>
        /// <param name="columnIndex">列索引</param>
        /// <returns></returns>
        int GetColumnWidth(int columnIndex);

        /// <summary>
        /// 获取列样式
        /// </summary>
        /// <param name="columnIndex">列索引</param>
        /// <returns></returns>
        ICellStyle GetColumnsStyle(int columnIndex);
    }
}
