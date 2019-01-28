using System;
using Bing.Excels.Abstractions.Mappings;

namespace Bing.Excels.Attributes
{
    /// <summary>
    /// Excel 工作表属性
    /// </summary>
    public class ExcelSheetAttribute : Attribute, ISheetMap
    {
        /// <summary>
        /// 是否自动索引。如果<see cref="IColumnMap.Index"/>未设置，并且<see cref="AutoIndex"/>设置为true，则将尝试通过其标题查找列索引
        /// </summary>
        public bool AutoIndex { get; set; }

        /// <summary>
        /// 标题
        /// </summary>
        public string Title { get; set; }

        /// <summary>
        /// 工作表索引
        /// </summary>
        public int Index { get; set; } = -1;
    }
}
