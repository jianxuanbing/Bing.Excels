namespace Bing.Excels.Abstractions
{
    /// <summary>
    /// Excel文件信息
    /// </summary>
    public interface IExcelInfo
    {
        /// <summary>
        /// 公司名称
        /// </summary>
        string Company { get; set; }

        /// <summary>
        /// 作者
        /// </summary>
        string Author { get; set; }

        /// <summary>
        /// 主题
        /// </summary>
        string Subject { get; set; }

        /// <summary>
        /// 是否使用创建信息（Excel公司或作者信息）
        /// </summary>
        bool UseCreateInfo { get; set; }

        /// <summary>
        /// 是否使用*.xlsx文件扩展名
        /// </summary>
        bool UseXlsx { get; set; }
    }
}
