using System.Collections.Generic;
using System.IO;

namespace Bing.Excels.Abstractions.Elements
{
    /// <summary>
    /// 工作簿
    /// </summary>
    public interface IWorkbook
    {
        /// <summary>
        /// 工作表列表
        /// </summary>
        List<ISheet> Sheets { get; }

        /// <summary>
        /// 工作表
        /// </summary>
        /// <param name="sheetIndex">工作表索引</param>
        /// <returns></returns>
        ISheet this[int sheetIndex] { get; }

        /// <summary>
        /// 文件名称
        /// </summary>
        string FileName { get; }

        /// <summary>
        /// 工作表数
        /// </summary>
        int SheetNumber { get; }

        /// <summary>
        /// 加载数据
        /// </summary>
        /// <param name="path">路径</param>
        void Load(string path);

        /// <summary>
        /// 添加工作表
        /// </summary>
        /// <param name="sheet">工作表</param>
        void Add(ISheet sheet);

        /// <summary>
        /// 创建工作表
        /// </summary>
        /// <returns></returns>
        ISheet CreateSheet();

        /// <summary>
        /// 创建工作表
        /// </summary>
        /// <param name="name">工作表名称</param>
        /// <returns></returns>
        ISheet CreateSheet(string name);

        /// <summary>
        /// 根据工作表名称获取工作表
        /// </summary>
        /// <param name="name">工作表名称</param>
        /// <returns></returns>
        ISheet GetSheet(string name);        

        /// <summary>
        /// 保存到文件
        /// </summary>
        /// <param name="path">文件路径</param>
        /// <returns></returns>
        string SaveToFile(string path);

        /// <summary>
        /// 写入到流中
        /// </summary>
        /// <param name="stream">内存流</param>
        void Write(Stream stream);
    }
}
