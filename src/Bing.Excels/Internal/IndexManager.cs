using System.Collections.Generic;
using System.Linq;

namespace Bing.Excels.Internal
{
    /// <summary>
    /// 索引管理器
    /// </summary>
    public class IndexManager
    {
        /// <summary>
        /// 索引范围列表
        /// </summary>
        private readonly List<IndexRange> _list;

        /// <summary>
        /// 初始化一个<see cref="IndexManager"/>类型的实例
        /// </summary>
        /// <param name="endIndex">结束索引</param>
        public IndexManager(int endIndex=10000)
        {
            _list = new List<IndexRange>() {new IndexRange(0, endIndex) };
        }

        /// <summary>
        /// 获取索引
        /// </summary>
        /// <param name="span">跨度</param>
        /// <returns></returns>
        public int GetIndex(int span = 1)
        {
            var range = _list.First();
            var index = range.GetIndex(span);
            if (range.IsEnd)
            {
                _list.Remove(range);
            }

            return index;
        }

        /// <summary>
        /// 添加索引
        /// </summary>
        /// <param name="index">索引</param>
        /// <param name="span">跨度</param>
        public void AddIndex(int index, int span = 1)
        {
            foreach (var range in _list)
            {
                if (range.Contains(index))
                {
                    AddIndex(range,index,span);
                    return;
                }
            }
        }

        /// <summary>
        /// 添加索引
        /// </summary>
        /// <param name="range">索引范围</param>
        /// <param name="index">索引</param>
        /// <param name="span">跨度</param>
        public void AddIndex(IndexRange range, int index, int span)
        {
            var newRange = range.Split(index, span);
            if (newRange == null)
            {
                return;
            }

            _list.Add(newRange);
        }
    }
}
