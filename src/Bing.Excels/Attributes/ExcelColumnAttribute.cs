using System;
using Bing.Excels.Abstractions.Mappings;

namespace Bing.Excels.Attributes
{
    /// <summary>
    /// Excel 列属性
    /// </summary>
    public class ExcelColumnAttribute:Attribute,IColumnMap
    {
        /// <summary>
        /// 名称
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// 标题
        /// </summary>
        public string Title { get; set; }

        /// <summary>
        /// 是否启用自动列索引，如果<see cref="Index"/>未设置，并且<see cref="AutoIndex"/>设置为true，则将尝试通过其标题查找列索引
        /// </summary>
        public bool AutoIndex { get; set; }

        /// <summary>
        /// 列索引
        /// </summary>
        public int Index { get; set; } = -1;

        /// <summary>
        /// 是否允许合并相同值的单元格
        /// </summary>
        public bool AllowMerge { get; set; }

        /// <summary>
        /// 是否忽略映射
        /// </summary>
        public bool Ignore { get; set; }

        /// <summary>
        /// 格式化
        /// </summary>
        public string Format { get; set; }

        /// <summary>
        /// 默认值
        /// </summary>
        public object DefaultValue { get; set; }

        /// <summary>
        /// 值
        /// </summary>
        public string Value { get; set; }
        
        /// <summary>
        /// 枚举类型
        /// </summary>
        public Type EnumType { get; set; }

        /// <summary>
        /// 是否忽略导出映射
        /// </summary>
        public bool IgnoreExport { get; set; }

        /// <summary>
        /// 是否忽略导入映射
        /// </summary>
        public bool IgnoreImport { get; set; }

        /// <summary>
        /// 初始化一个<see cref="ExcelColumnAttribute"/>类型的实例
        /// </summary>
        /// <param name="name">名称</param>
        public ExcelColumnAttribute(string name)
        {
            Name = name;
        }
    }
}
