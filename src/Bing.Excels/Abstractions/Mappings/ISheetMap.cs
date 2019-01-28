namespace Bing.Excels.Abstractions.Mappings
{
    /// <summary>
    /// 工作表映射
    /// </summary>
    public interface ISheetMap
    {
        /// <summary>
        /// 是否自动索引。如果<see cref="IColumnMap.Index"/>未设置，并且<see cref="AutoIndex"/>设置为true，则将尝试通过其标题查找列索引
        /// </summary>
        bool AutoIndex { get; }

        /// <summary>
        /// 标题
        /// </summary>
        string Title { get; }

        /// <summary>
        /// 工作表索引
        /// </summary>
        int Index { get; }
    }
}
