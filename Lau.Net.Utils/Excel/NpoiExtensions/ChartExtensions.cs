using NPOI.SS.UserModel.Charts;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NPOI.SS.Util;
using NPOI.OpenXmlFormats.Spreadsheet;

namespace Lau.Net.Utils.Excel.NpoiExtensions
{
    /// <summary>
    /// 图表相关扩展方法
    /// </summary>
    public static class ChartExtensions
    {
        /// <summary>
        /// 创建图表
        /// </summary>
        /// <param name="sheet">ISheet</param>
        /// <param name="startRowIndex">图表从第几行开始</param>
        /// <param name="endRowIndex">图标到第几行结束(不包含该行）</param>
        /// <param name="startColumnIndex">图表从第几列开始</param>
        /// <param name="endColumnIndex">图标到第几列结束(不包含该列）</param>
        /// <returns></returns>
        public static IChart CreateChart(this ISheet sheet, int startRowIndex, int endRowIndex, int startColumnIndex, int endColumnIndex)
        {
            IDrawing drawing = sheet.CreateDrawingPatriarch();
            IClientAnchor anchor = drawing.CreateAnchor(0, 0, 0, 0, startColumnIndex, startRowIndex, endColumnIndex, endRowIndex);
            IChart chart = drawing.CreateChart(anchor);
            return chart;
        }

        /// <summary>
        /// 获取或创建图例
        /// </summary>
        /// <param name="chart"></param>
        /// <param name="position">图例显示位置</param>
        /// <returns></returns>
        public static IChartLegend GetOrCreateLegend(this IChart chart, LegendPosition position)
        {
            IChartLegend legend = chart.GetOrCreateLegend();
            legend.Position = position;
            return legend;
        }

        /// <summary>
        /// 获取Sheet中字符串或数字Chart数据源
        /// </summary>
        /// <typeparam name="T">T必须为string或double类型</typeparam>
        /// <param name="sheet"></param>
        /// <param name="startRowIndex"></param>
        /// <param name="endRowIndex"></param>
        /// <param name="startColumnIndex"></param>
        /// <param name="endColumnIndex"></param>
        /// <returns></returns>
        /// <exception cref="Exception"></exception>
        public static IChartDataSource<T> GetChartDataSource<T>(this ISheet sheet, int startRowIndex, int endRowIndex, int startColumnIndex, int endColumnIndex)
        {
            if (typeof(T) != typeof(string) && typeof(T) != typeof(double))
            {
                throw new Exception("GetChartStringDataSource 泛型约束类T必须为string或double类型");
            }
            if (typeof(T) == typeof(string))
            {
                return (IChartDataSource<T>)DataSources.FromStringCellRange(sheet, new CellRangeAddress(startRowIndex, endRowIndex, startColumnIndex, endColumnIndex));
            }
            else
            {
                return (IChartDataSource<T>)DataSources.FromNumericCellRange(sheet, new CellRangeAddress(startRowIndex, endRowIndex, startColumnIndex, endColumnIndex));
            }
        } 
    }
}
