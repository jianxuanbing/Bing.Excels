namespace Bing.Excels.Abstractions
{
    /// <summary>
    /// Excel上下文
    /// </summary>
    public interface IExcelContext
    {
        /// <summary>
        /// Excel文件信息
        /// </summary>
        IExcelInfo Info { get; set; }

        /// <summary>
        /// 导出数据数量
        /// </summary>
        long ExportNumber { get; set; }

        /// <summary>
        /// 导出文件存放路径
        /// </summary>
        string ExportPath { get; set; }

        // 导入配置
        // 导出配置
    }
}
