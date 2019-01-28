﻿using System;

namespace Bing.Excels.Abstractions.Mappings
{
    /// <summary>
    /// 列映射
    /// </summary>
    public interface IColumnMap
    {
        /// <summary>
        /// 名称
        /// </summary>
        string Name { get; }

        /// <summary>
        /// 标题
        /// </summary>
        string Title { get; }

        /// <summary>
        /// 是否启用自动列索引，如果<see cref="Index"/>未设置，并且<see cref="AutoIndex"/>设置为true，则将尝试通过其标题查找列索引
        /// </summary>
        bool AutoIndex { get; }

        /// <summary>
        /// 列索引
        /// </summary>
        int Index { get; }

        /// <summary>
        /// 是否允许合并相同值的单元格
        /// </summary>
        bool AllowMerge { get; }

        /// <summary>
        /// 是否忽略映射
        /// </summary>
        bool Ignore { get; }

        /// <summary>
        /// 格式化
        /// </summary>
        string Format { get; }

        /// <summary>
        /// 默认值
        /// </summary>
        object DefaultValue { get; }

        /// <summary>
        /// 值
        /// </summary>
        string Value { get; }

        /// <summary>
        /// 枚举类型
        /// </summary>
        Type EnumType { get; }

        /// <summary>
        /// 是否忽略导出映射
        /// </summary>
        bool IgnoreExport { get; }

        /// <summary>
        /// 是否忽略导入映射
        /// </summary>
        bool IgnoreImport { get; }
    }
}
