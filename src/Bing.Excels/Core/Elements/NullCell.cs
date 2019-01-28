﻿using Bing.Excels.Abstractions.Elements;

namespace Bing.Excels.Core.Elements
{
    /// <summary>
    /// 空单元格
    /// </summary>
    public class NullCell : Cell, ICell
    {
        /// <summary>
        /// 初始化一个<see cref="NullCell"/>类型的实例
        /// </summary>
        public NullCell() : base("", 1, 1)
        {
        }
    }
}
