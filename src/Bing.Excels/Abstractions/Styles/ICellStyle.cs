using Bing.Excels.Core.Styles;

namespace Bing.Excels.Abstractions.Styles
{
    /// <summary>
    /// 单元格样式
    /// </summary>
    public interface ICellStyle
    {
        /// <summary>
        /// 水平对齐
        /// </summary>
        HorizontalAlignment Alignment { get; set; }

        /// <summary>
        /// 垂直对齐
        /// </summary>
        VerticalAlignment VerticalAlignment { get; set; }

        /// <summary>
        /// 背景色
        /// </summary>
        Color BackgroundColor { get; set; }

        /// <summary>
        /// 前景色。需要同时设置 FillPattern.SolidForeground
        /// </summary>
        Color ForegroundColor { get; set; }

        /// <summary>
        /// 填充模式
        /// </summary>
        FillPattern FillPattern { get; set; }

        /// <summary>
        /// 边框色
        /// </summary>
        Color BorderColor { get; set; }

        /// <summary>
        /// 字号
        /// </summary>
        short FontSize { get; set; }

        /// <summary>
        /// 字体颜色
        /// </summary>
        Color FontColor { get; set; }

        /// <summary>
        /// 字体加粗
        /// </summary>
        short FontBoldWeight { get; set; }

        /// <summary>
        /// 内容是否自动换行
        /// </summary>
        bool IsWrap { get; set; }

        /// <summary>
        /// 斜体。将文字变为斜体
        /// </summary>
        bool Italic { get; set; }
    }
}
