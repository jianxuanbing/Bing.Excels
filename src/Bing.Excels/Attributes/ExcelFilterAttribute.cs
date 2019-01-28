using System;
using Bing.Excels.Abstractions.Mappings;

namespace Bing.Excels.Attributes
{
    /// <summary>
    /// Excel 筛选器属性
    /// </summary>
    public class ExcelFilterAttribute:Attribute,IFilterMap
    {
        /// <summary>
        /// 左上角列索引
        /// </summary>
        public int FirstColumn { get; set; }

        /// <summary>
        /// 左上角行索引
        /// </summary>
        public int FirstRow { get; set; }

        /// <summary>
        /// 右下角列索引
        /// </summary>
        public int LastColumn { get; set; }

        /// <summary>
        /// 右下角行索引。如果<see cref="LastRow"/>为空，则该值按代码动态计算
        /// </summary>
        public int? LastRow { get; set; }
    }
}
