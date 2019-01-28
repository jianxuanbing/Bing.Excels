using System;
using Bing.Excels.Abstractions.Elements;
using Bing.Excels.Abstractions.Styles;
using Bing.Excels.Core.Styles;

namespace Bing.Excels.Core.Elements
{
    /// <summary>
    /// 单元格
    /// </summary>
    public class Cell : ICell
    {
        /// <summary>
        /// 单元格值
        /// </summary>
        private object _value;

        /// <summary>
        /// 单元格值
        /// </summary>
        public object Value
        {
            get => _value;
            set => this.SetValue(value);
        }

        /// <summary>
        /// 单元格类型
        /// </summary>
        public CellType CellType { get; private set; }

        /// <summary>
        /// 行索引
        /// </summary>
        public int RowIndex => Row.RowIndex;

        /// <summary>
        /// 列索引
        /// </summary>
        private int _columnIndex;

        /// <summary>
        /// 列索引
        /// </summary>
        public int ColumnIndex
        {
            get => _columnIndex;
            set => this.SetColumnIndex(value);
        }

        /// <summary>
        /// 行跨度
        /// </summary>
        public int RowSpan { get; set; }

        /// <summary>
        /// 列跨度
        /// </summary>
        public int ColumnSpan { get; set; }

        /// <summary>
        /// 单元行
        /// </summary>
        public IRow Row { get; set; }

        /// <summary>
        /// 工作表
        /// </summary>
        public ISheet Sheet => Row.Sheet;

        /// <summary>
        /// 单元格样式
        /// </summary>
        public ICellStyle CellStyle { get; set; }

        /// <summary>
        /// 单元格范围
        /// </summary>
        public ICellRangeAddress Range { get; private set; }

        /// <summary>
        /// 是否合并单元格
        /// </summary>
        public bool IsMergedCell => Range.RowSpan > 1 || Range.ColumnSpan > 1;

        /// <summary>
        /// 空单元格
        /// </summary>
        public static ICell Null => new NullCell();

        /// <summary>
        /// 初始化一个<see cref="Cell"/>类型的实例
        /// </summary>
        /// <param name="value">值</param>
        /// <param name="columnSpan">列跨度</param>
        /// <param name="rowSpan">行跨度</param>
        public Cell(object value, int columnSpan = 1, int rowSpan = 1)
        {
            this.SetValue(value);
            ColumnSpan = columnSpan;
            RowSpan = rowSpan;
            Range = new CellRangeAddress(Row.RowIndex, ColumnIndex, RowSpan, ColumnSpan);
        }

        /// <summary>
        /// 设置单元格的值
        /// </summary>
        /// <param name="value">值</param>
        public void SetValue(object value)
        {
            ResolveCellType(value);
            this._value = value;
        }

        /// <summary>
        /// 解析单元格类型
        /// </summary>
        /// <param name="value">值</param>
        private void ResolveCellType(object value)
        {
            if (value == null)
            {
                CellType = CellType.Unknown;
                return;
            }

            TypeCode valueTypeCode = Type.GetTypeCode(value.GetType());
            switch (valueTypeCode)
            {
                case TypeCode.String:
                    CellType = CellType.String;
                    break;
                case TypeCode.Boolean:
                    CellType = CellType.Boolean;
                    break;
                case TypeCode.DateTime:
                    CellType = CellType.Date;
                    break;
                case TypeCode.Int16:
                case TypeCode.Int32:
                case TypeCode.Int64:
                case TypeCode.Byte:
                case TypeCode.Single:
                case TypeCode.Double:
                case TypeCode.Decimal:
                case TypeCode.UInt16:
                case TypeCode.UInt32:
                case TypeCode.UInt64:
                    CellType = CellType.Number;
                    break;
                default:
                    CellType = CellType.Empty;
                    break;
            }
        }

        /// <summary>
        /// 设置列索引
        /// </summary>
        /// <param name="value">列索引</param>
        private void SetColumnIndex(int value)
        {
            _columnIndex = value;
            Range = new CellRangeAddress(Row.RowIndex, ColumnIndex, RowSpan, ColumnSpan);
        }
    }
}
