using Lau.Net.Utils.Excel.NpoiExtensions;
using NPOI.SS.UserModel;
using NPOI.SS.UserModel.Charts;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lau.Net.Utils.Excel
{
    /// <summary>
    /// 创建图表
    /// </summary>
    /// <typeparam name="Tx">X轴数据源的类型，只能为double或string类型</typeparam>
    public class NpoiChart<Tx>:IDisposable where Tx : IComparable, IConvertible
    {
        private ISheet _sheet = null;
        private IChart _chart = null;
        //X轴数据源
        IChartDataSource<Tx> _xAxis = null;
        //底部X轴线
        IChartAxis _bottomAxis = null;
        //左边值Y轴线
        IValueAxis _leftAxis = null;

        IBarChartData<Tx, double> _barChartData = null;

        /// <summary>
        /// 创建Chart
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="startRowIndex">图表从第几行开始</param>
        /// <param name="endRowIndex">图标到第几行结束(不包含该行）</param>
        /// <param name="startColumnIndex">图表从第几列开始</param>
        /// <param name="endColumnIndex">图标到第几列结束(不包含该列）</param>
        /// <param name="legendPosition">图例显示位置</param>
        /// <exception cref="Exception"></exception>
        public NpoiChart(ISheet sheet, int startRowIndex, int endRowIndex, int startColumnIndex, int endColumnIndex, LegendPosition legendPosition = LegendPosition.Bottom)
        {
            this._sheet = sheet;
            if(typeof(Tx) != typeof(string) && typeof(Tx) != typeof(double))
            {
                throw new Exception("NpoiChart 泛型类T必须为string或double类型");
            }
            CreateChart(startRowIndex, endRowIndex, startColumnIndex, endColumnIndex, legendPosition); 
        }

        /// <summary>
        /// 创建Chart
        /// </summary>
        /// <param name="startRowIndex">图表从第几行开始</param>
        /// <param name="endRowIndex">图标到第几行结束(不包含该行）</param>
        /// <param name="startColumnIndex">图表从第几列开始</param>
        /// <param name="endColumnIndex">图标到第几列结束(不包含该列）</param>
        /// <param name="legendPosition">图例显示位置</param>
        /// <returns></returns>
        private void CreateChart(int startRowIndex, int endRowIndex, int startColumnIndex, int endColumnIndex, LegendPosition legendPosition)
        {
            IChart chart = _sheet.CreateChart(startRowIndex, endRowIndex, startColumnIndex, endColumnIndex);
            chart.GetOrCreateLegend(legendPosition);

            //底部X轴线
            IChartAxis bottomAxis = chart.ChartAxisFactory.CreateCategoryAxis(AxisPosition.Bottom);
            bottomAxis.MajorTickMark = AxisTickMark.None;

            //左边值Y轴线
            IValueAxis leftAxis = chart.ChartAxisFactory.CreateValueAxis(AxisPosition.Left);
            leftAxis.Crosses = AxisCrosses.AutoZero;
            leftAxis.SetCrossBetween(AxisCrossBetween.Between);

            _chart = chart;
            _bottomAxis = bottomAxis;
            _leftAxis = leftAxis;
        }

        /// <summary>
        /// 设置X轴数据源
        /// </summary>
        /// <param name="startRowIndex">X轴数据源起始于哪一行</param>
        /// <param name="endRowIndex"></param>
        /// <param name="startColumnIndex"></param>
        /// <param name="endColumnIndex"></param>
        /// <returns></returns>
        public NpoiChart<Tx> SetXAxisData(int startRowIndex, int endRowIndex, int startColumnIndex, int endColumnIndex)
        {
            this._xAxis = _sheet.GetChartDataSource<Tx>(startRowIndex, endRowIndex, startColumnIndex, endColumnIndex);
            return this;
        }

        /// <summary>
        /// 添加Bar类型的serie
        /// </summary>
        /// <param name="serieTitle">serie</param>
        /// <param name="startRowIndex"></param>
        /// <param name="endRowIndex"></param>
        /// <param name="startColumnIndex"></param>
        /// <param name="endColumnIndex"></param>
        /// <returns></returns>
        public NpoiChart<Tx> AddBarSerie(string serieTitle, int startRowIndex, int endRowIndex, int startColumnIndex, int endColumnIndex)
        {
            //Y轴数据源
            IChartDataSource<double> valueAxis = _sheet.GetChartDataSource<double>(startRowIndex, endRowIndex, startColumnIndex, endColumnIndex);
            if(_barChartData == null)
            {   //如果不用同一个barChartData，那多次添加barSerie的话，则只会叠加在同一个bar上
                _barChartData = _chart.ChartDataFactory.CreateBarChartData<Tx, double>();
            }
            //IBarChartData<Tx, double> chartData = _chart.ChartDataFactory.CreateBarChartData<Tx, double>();
            var serie = _barChartData.AddSeries(_xAxis, valueAxis);
            serie.SetTitle(serieTitle);
            return this;
        }

        /// <summary>
        /// 添加Line类型的serie
        /// </summary>
        /// <param name="serieTitle">serie</param>
        /// <param name="startRowIndex"></param>
        /// <param name="endRowIndex"></param>
        /// <param name="startColumnIndex"></param>
        /// <param name="endColumnIndex"></param>
        /// <returns></returns>
        public NpoiChart<Tx> AddLineSerie( string serieTitle, int startRowIndex, int endRowIndex, int startColumnIndex, int endColumnIndex)
        {
            //Y轴数据源
            IChartDataSource<double> valueAxis = _sheet.GetChartDataSource<double>(startRowIndex, endRowIndex, startColumnIndex, endColumnIndex);

            ILineChartData<Tx, double> chartData = _chart.ChartDataFactory.CreateLineChartData<Tx, double>();
            var serie = chartData.AddSeries(_xAxis, valueAxis);
            serie.SetTitle(serieTitle);
            _chart.Plot(chartData, _bottomAxis, _leftAxis);
            return this;
        }

        /// <summary>
        /// 添加Scatter类型的serie
        /// </summary>
        /// <param name="serieTitle">serie</param>
        /// <param name="startRowIndex"></param>
        /// <param name="endRowIndex"></param>
        /// <param name="startColumnIndex"></param>
        /// <param name="endColumnIndex"></param>
        /// <returns></returns>
        public NpoiChart<Tx> AddScatterSerie(string serieTitle, int startRowIndex, int endRowIndex, int startColumnIndex, int endColumnIndex)
        {
            //Y轴数据源
            IChartDataSource<double> valueAxis = _sheet.GetChartDataSource<double>(startRowIndex, endRowIndex, startColumnIndex, endColumnIndex);

            IScatterChartData<Tx, double> chartData = _chart.ChartDataFactory.CreateScatterChartData<Tx, double>();
            var serie = chartData.AddSeries(_xAxis, valueAxis);
            serie.SetTitle(serieTitle);
            _chart.Plot(chartData, _bottomAxis, _leftAxis);
            return this;
        }

        public void Dispose()
        {
            if(_barChartData != null)
            {
                //如果在AddBarSerie调用该方法的话，多次调用AddBarSerie的话数据会叠加在一起
                _chart.Plot(_barChartData, _bottomAxis, _leftAxis);
            }            
        }
    }
    
}
