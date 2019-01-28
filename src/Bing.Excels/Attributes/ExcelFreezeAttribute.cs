using System;
using Bing.Excels.Abstractions.Mappings;

namespace Bing.Excels.Attributes
{
    /// <summary>
    /// Excel 窗格冻结属性
    /// </summary>
    public class ExcelFreezeAttribute:Attribute,IFreezeMap
    {
        /// <summary>
        /// 列号。要拆分的列号，默认:0
        /// </summary>
        public int ColumnSplit { get; set; }

        /// <summary>
        /// 行号。要拆分的行号，默认：1
        /// </summary>
        public int RowSpit { get; set; }

        /// <summary>
        /// 列索引。最左侧的列索引，默认：0
        /// </summary>
        public int LeftMostColumn { get; set; }

        /// <summary>
        /// 行索引。最顶行的行索引，默认：1
        /// </summary>
        public int TopRow { get; set; }
    }
}
