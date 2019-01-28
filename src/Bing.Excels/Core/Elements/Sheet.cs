using System.Collections.Generic;
using System.Linq;
using Bing.Excels.Abstractions.Elements;
using Bing.Excels.Abstractions.Styles;

namespace Bing.Excels.Core.Elements
{
    /// <summary>
    /// 工作表
    /// </summary>
    public class Sheet:ISheet
    {
        /// <summary>
        /// 是否当前选中工作表
        /// </summary>
        private bool _isSelected;

        /// <summary>
        /// 头部范围
        /// </summary>
        private readonly IRange _header;

        /// <summary>
        /// 正文单元范围
        /// </summary>
        private IRange _body;

        /// <summary>
        /// 底部单元范围
        /// </summary>
        private IRange _footer;

        /// <summary>
        /// 当前行索引
        /// </summary>
        private int _rowIndex;

        /// <summary>
        /// 总标题
        /// </summary>
        public string Title { get; set; }

        /// <summary>
        /// 工作表名称
        /// </summary>
        public string SheetName { get; set; }

        /// <summary>
        /// 工作表索引
        /// </summary>
        public int SheetIndex { get; }

        /// <summary>
        /// 工作簿
        /// </summary>
        public IWorkbook Workbook { get; }

        /// <summary>
        /// 是否当前选中工作表
        /// </summary>
        public bool IsSelected
        {
            get => _isSelected;
            set => this.SetActive(value);
        }

        /// <summary>
        /// 表头行数
        /// </summary>
        public int HeadRowNumber => _header?.RowNumber ?? 0;

        /// <summary>
        /// 正文行数
        /// </summary>
        public int BodyRowNumber => _body?.RowNumber ?? 0;

        /// <summary>
        /// 页脚行数
        /// </summary>
        public int FootRowNumber => _footer?.RowNumber ?? 0;

        /// <summary>
        /// 总行数
        /// </summary>
        public int RowNumber => HeadRowNumber + BodyRowNumber + FootRowNumber;

        /// <summary>
        /// 总列数
        /// </summary>
        public int ColumnNumber => _body?.ColumnNumber ?? _header.ColumnNumber;

        /// <summary>
        /// 默认列宽度
        /// </summary>
        public int DefaultColumnWidth { get; set; }

        /// <summary>
        /// 默认行高度
        /// </summary>
        public int DefaultRowHeight { get; set; }

        /// <summary>
        /// 初始化一个<see cref="Sheet"/>类型的实例
        /// </summary>
        /// <param name="workbook">工作簿</param>
        /// <param name="sheetIndex">工作表索引</param>
        public Sheet(IWorkbook workbook, int sheetIndex)
        {
            Workbook = workbook;
            SheetIndex = sheetIndex;
            _header = new Range(this);
            _rowIndex = 0;
        }

        /// <summary>
        /// 设置是否选择了工作表
        /// </summary>
        /// <param name="value">是否选择工作表</param>
        public void SetActive(bool value)
        {
            this._isSelected = value;
            if (value)
            {
                foreach (var sheet in this.Workbook.Sheets.Where(x => x.SheetIndex != SheetIndex))
                {
                    sheet.SetActive(false);
                }
            }
        }

        /// <summary>
        /// 获取表头
        /// </summary>
        /// <returns></returns>
        public List<IRow> GetHeader()
        {
            return _header.GetRows();
        }

        /// <summary>
        /// 获取表格正文
        /// </summary>
        /// <returns></returns>
        public List<IRow> GetBody()
        {
            return _body == null ? new List<IRow>() : _body.GetRows();
        }

        /// <summary>
        /// 获取页脚
        /// </summary>
        /// <returns></returns>
        public List<IRow> GetFooter()
        {
            return _footer == null ? new List<IRow>() : _footer.GetRows();
        }

        /// <summary>
        /// 添加表头
        /// </summary>
        /// <param name="titles">标题</param>
        public void AddHeadRow(params string[] titles)
        {
            if (titles == null)
            {
                return;
            }

            AddHeadRow(titles.Select(CreateCell).ToArray());
        }

        /// <summary>
        /// 创建单元格
        /// </summary>
        /// <param name="value">值</param>
        /// <returns></returns>
        private static ICell CreateCell(object value)
        {
            return new Cell(value);
        }

        /// <summary>
        /// 添加表头
        /// </summary>
        /// <param name="cells">单元格</param>
        public void AddHeadRow(params ICell[] cells)
        {
            if (cells == null)
            {
                return;
            }
            AddRowToHeader(cells);
            ResetFirstColumnSpan();
        }

        /// <summary>
        /// 添加表头行
        /// </summary>
        /// <param name="cells">单元格集合</param>
        private void AddRowToHeader(IEnumerable<ICell> cells)
        {
            _header.AddRow(_rowIndex,cells);
            _rowIndex++;
        }

        /// <summary>
        /// 重置第一行的列跨度，第一行可能为总标题
        /// </summary>
        private void ResetFirstColumnSpan()
        {
            if (_rowIndex < 2)
            {
                return;
            }

            if (_header.RowNumber == 0)
            {
                return;
            }

            if (_header[0].ColumnNumber > 1)
            {
                return;
            }

            if (_header.RowNumber > 1)
            {
                _header[0][0].ColumnSpan = _header[1].ColumnNumber;
                return;
            }

            if (_body == null || _body.RowNumber == 0)
            {
                return;
            }

            _header[0][0].ColumnSpan = _body[0].ColumnNumber;
        }

        /// <summary>
        /// 添加正文
        /// </summary>
        /// <param name="cellValues">值</param>
        public void AddBodyRow(params object[] cellValues)
        {
            if (cellValues == null)
            {
                return;
            }

            AddBodyRow(cellValues.Select(x => new Cell(x)));
        }

        /// <summary>
        /// 添加正文
        /// </summary>
        /// <param name="cells">单元格集合</param>
        public void AddBodyRow(IEnumerable<ICell> cells)
        {
            if (cells == null)
            {
                return;
            }

            GetBodyRange().AddRow(_rowIndex, cells);
            _rowIndex++;
            ResetFirstColumnSpan();
        }

        /// <summary>
        /// 获取正文单元范围
        /// </summary>
        /// <returns></returns>
        private IRange GetBodyRange()
        {
            if (_body != null)
            {
                return _body;
            }

            _body = new Range(this, _rowIndex);
            return _body;
        }

        /// <summary>
        /// 添加页脚
        /// </summary>
        /// <param name="cellValues">值</param>
        public void AddFootRow(params string[] cellValues)
        {
            if (cellValues == null)
            {
                return;
            }

            AddFootRow(cellValues.Select(CreateCell).ToArray());
        }

        /// <summary>
        /// 添加页脚
        /// </summary>
        /// <param name="cells">单元格集合</param>
        public void AddFootRow(params ICell[] cells)
        {
            if (cells == null)
            {
                return;
            }

            GetFootRange().AddRow(_rowIndex, cells);
            _rowIndex++;
        }

        /// <summary>
        /// 获取页脚单元范围
        /// </summary>
        /// <returns></returns>
        private IRange GetFootRange()
        {
            if (_footer != null)
            {
                return _footer;
            }

            _footer = new Range(this, _rowIndex);
            return _footer;
        }

        /// <summary>
        /// 设置列宽
        /// </summary>
        /// <param name="columIndex">列索引</param>
        /// <param name="width">宽度</param>
        public void SetColumnWidth(int columIndex, int width)
        {

        }

        /// <summary>
        /// 获取列宽
        /// </summary>
        /// <param name="columnIndex">列索引</param>
        /// <returns></returns>
        public int GetColumnWidth(int columnIndex)
        {
            throw new System.NotImplementedException();
        }

        /// <summary>
        /// 获取列样式
        /// </summary>
        /// <param name="columnIndex">列索引</param>
        /// <returns></returns>
        public ICellStyle GetColumnsStyle(int columnIndex)
        {
            throw new System.NotImplementedException();
        }
    }
}
