using System.Collections.Generic;

namespace Bing.Excels.Abstractions.Imports
{
    /// <summary>
    /// 导入器
    /// </summary>
    public interface IImport
    {
        /// <summary>
        /// 获取结果
        /// </summary>
        /// <typeparam name="T">返回类型</typeparam>
        /// <returns></returns>
        List<T> GetResult<T>() where T : class, new();
    }
}
