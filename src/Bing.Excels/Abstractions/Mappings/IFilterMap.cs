namespace Bing.Excels.Abstractions.Mappings
{
    /// <summary>
    /// 筛选器映射
    /// </summary>
    public interface IFilterMap
    {
        /// <summary>
        /// 左上角列索引
        /// </summary>
        int FirstColumn { get; }

        /// <summary>
        /// 左上角行索引
        /// </summary>
        int FirstRow { get; }

        /// <summary>
        /// 右下角列索引
        /// </summary>
        int LastColumn { get; }

        /// <summary>
        /// 右下角行索引。如果<see cref="LastRow"/>为空，则该值按代码动态计算
        /// </summary>
        int? LastRow { get; }
    }
}
