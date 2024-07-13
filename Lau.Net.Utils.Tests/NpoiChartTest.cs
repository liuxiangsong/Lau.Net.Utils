using Lau.Net.Utils.Excel;
using Lau.Net.Utils.Excel.NpoiExtensions;
using NPOI.SS.UserModel.Charts;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using NPOI.Util;

namespace Lau.Net.Utils.Tests
{
    [TestFixture]
    public class NpoiChartTest
    {
        [Test]
        public void CreateBarChartTest()
        {
            var dt = CreateTable();
            var workbook = NpoiUtil.CreateWorkbook(dt);
            var sheet = workbook.GetSheetAt(0);
            CreateBarLineChart(sheet);
            //CreateBarChart(sheet);
            var filePath = Path.Combine("E:\\", "test", $"{DateTime.Now.ToString("yyyy-MM-dd HHmmss")}.xls");
            workbook.SaveToExcel(filePath);
        }

        [Test]
        public void CreateLineChartTest()
        {
            var dt = CreateTable();
            var workbook = NpoiUtil.CreateWorkbook(dt);
            var sheet = workbook.GetSheetAt(0);
            
            SetCellFormula(sheet);

            //CreateScatterChart(sheet);
            CreateLineChart(sheet);
            var filePath = Path.Combine("E:\\", "test", $"{DateTime.Now.ToString("yyyy-MM-dd HHmmss")}.xls");
            workbook.SaveToExcel(filePath);
        }

        private void SetCellFormula( ISheet sheet)
        {
            var workbook = sheet.Workbook;
            var style = workbook.CreateCellStyle();
            var cell = sheet.GetOrCreateCell(2, 4);
            cell.SetCellFormula("IFERROR(B3/C3,\"未加工\")");
            //style.CloneStyleFrom(cell.CellStyle);
            style.DataFormat = workbook.CreateDataFormat().GetFormat("0.00%");

            cell.CellStyle = style;
        }

        /// <summary>
        /// 创建散点图
        /// </summary>
        /// <param name="sheet"></param> 
        static void CreateScatterChart(ISheet sheet)
        {
            IChart chart = sheet.CreateChart(13, 25, 0, 5);
            chart.GetOrCreateLegend(LegendPosition.TopRight);
            var lineChartdata = chart.ChartDataFactory.CreateScatterChartData<double, double>();

            // Use a category axis for the bottom axis.
            IChartAxis bottomAxis = chart.ChartAxisFactory.CreateCategoryAxis(AxisPosition.Bottom);
            IValueAxis leftAxis = chart.ChartAxisFactory.CreateValueAxis(AxisPosition.Left);
            //leftAxis.Crosses = AxisCrosses.AutoZero;

            int startDataRow = 1;
            int endDataRow = 12;
            IChartDataSource<double> xs = sheet.GetChartDataSource<double>(startDataRow, endDataRow, 0, 0);
            IChartDataSource<double> ys1 = sheet.GetChartDataSource<double>(startDataRow, endDataRow, 1, 1);
            IChartDataSource<double> ys2 = sheet.GetChartDataSource<double>(startDataRow, endDataRow, 2, 2);

            var s1 = lineChartdata.AddSeries(xs, ys1);
            s1.SetTitle("生产总数量");
            var s2 = lineChartdata.AddSeries(xs, ys2);
            s2.SetTitle("生产合格数");

            chart.Plot(lineChartdata, bottomAxis, leftAxis);
        }


        /// <summary>
        /// 创建折线图
        /// </summary>
        /// <param name="sheet"></param> 
        static void CreateLineChart(ISheet sheet)
        {
            int startDataRow = 1;
            int endDataRow = 12;
            using (var npoiChart = new NpoiChart<string>(sheet, 13, 25, 0, 5))
            {
                npoiChart.SetXAxisData(startDataRow, endDataRow, 0, 0)
                //.AddBarSerie("生产总数量", startDataRow, endDataRow, 1, 1)
                //.AddBarSerie("生产合格数", startDataRow, endDataRow, 2, 2)
                //.AddLineSerie("生产总数量", startDataRow, endDataRow, 3, 3)
                .AddScatterSerie("合格率", startDataRow, endDataRow, 4, 4);
            };


            //IChart chart = sheet.CreateChart( 13, 25, 0, 5);
            //chart.GetOrCreateLegend(LegendPosition.TopRight);

            //ILineChartData<double, double> lineChartdata = chart.ChartDataFactory.CreateLineChartData<double, double>();

            //// Use a category axis for the bottom axis.
            //IChartAxis bottomAxis = chart.ChartAxisFactory.CreateCategoryAxis(AxisPosition.Bottom);
            //IValueAxis leftAxis = chart.ChartAxisFactory.CreateValueAxis(AxisPosition.Left);
            ////leftAxis.Crosses = AxisCrosses.AutoZero;

            //IChartDataSource<double> xs = sheet.GetChartDataSource<double>(startDataRow, endDataRow, 0, 0);
            //IChartDataSource<double> ys1 = sheet.GetChartDataSource<double>(startDataRow, endDataRow, 1, 1);
            //IChartDataSource<double> ys2 = sheet.GetChartDataSource<double>(startDataRow, endDataRow, 2, 2);

            //var s1 = lineChartdata.AddSeries(xs, ys1);
            //s1.SetTitle("生产总数量");
            //var s2 = lineChartdata.AddSeries(xs, ys2);
            //s2.SetTitle("生产合格数");

            //chart.Plot(lineChartdata, bottomAxis, leftAxis);
        }

        /// <summary>
        /// 创建柱状图和拆线图
        /// </summary>
        /// <param name="sheet"></param>
        private static void CreateBarLineChart(ISheet sheet)
        {
            IChart chart = sheet.CreateChart(13, 25, 0, 5);
            chart.GetOrCreateLegend(LegendPosition.Bottom);

            //底部X轴线
            IChartAxis bottomAxis = chart.ChartAxisFactory.CreateCategoryAxis(AxisPosition.Bottom);
            bottomAxis.MajorTickMark = AxisTickMark.None;

            //左边值Y轴线
            IValueAxis leftAxis = chart.ChartAxisFactory.CreateValueAxis(AxisPosition.Left);
            leftAxis.Crosses = AxisCrosses.AutoZero;
            leftAxis.SetCrossBetween(AxisCrossBetween.Between);

            int startDataRow = 1;
            int endDataRow = 12;
            //X轴数据源
            IChartDataSource<string> categoryAxis = sheet.GetChartDataSource<string>(startDataRow, endDataRow, 0, 0);
            //Y轴数据源
            IChartDataSource<double> valueAxis = sheet.GetChartDataSource<double>(startDataRow, endDataRow, 1, 1);

            IBarChartData<string, double> barChartData = chart.ChartDataFactory.CreateBarChartData<string, double>();
            var serie = barChartData.AddSeries(categoryAxis, valueAxis);
            serie.SetTitle("生产总数量");
            chart.Plot(barChartData, bottomAxis, leftAxis);

            ILineChartData<string, double> lineChartdata = chart.ChartDataFactory.CreateLineChartData<string, double>();
            IChartDataSource<double> ys1 = sheet.GetChartDataSource<double>(startDataRow, endDataRow, 2, 2);
            var serie2 = lineChartdata.AddSeries(categoryAxis, ys1);
            serie2.SetTitle("生产合格数");
            chart.Plot(lineChartdata, bottomAxis, leftAxis);
        }

        /// <summary>
        /// 创建柱状图
        /// </summary>
        /// <param name="sheet"></param>
        private static void CreateBarChart(ISheet sheet)
        {
            IChart chart = sheet.CreateChart(13, 25, 0, 5);
            chart.GetOrCreateLegend(LegendPosition.Bottom);

            //底部X轴线
            IChartAxis bottomAxis = chart.ChartAxisFactory.CreateCategoryAxis(AxisPosition.Bottom);
            bottomAxis.MajorTickMark = AxisTickMark.None;

            //左边值Y轴线
            IValueAxis leftAxis = chart.ChartAxisFactory.CreateValueAxis(AxisPosition.Left);
            leftAxis.Crosses = AxisCrosses.AutoZero;
            leftAxis.SetCrossBetween(AxisCrossBetween.Between);

            int startDataRow = 1;
            int endDataRow = 12;
            //X轴数据源
            IChartDataSource<string> categoryAxis = sheet.GetChartDataSource<string>(startDataRow, endDataRow, 0, 0);
            //Y轴数据源
            IChartDataSource<double> valueAxis = sheet.GetChartDataSource<double>(startDataRow, endDataRow, 1, 1);
            IChartDataSource<double> value2Axis = sheet.GetChartDataSource<double>(startDataRow, endDataRow, 2, 2);

            //第一个string类型是X轴数据的类型，第二个double是Y轴数据的类型
            IBarChartData<string, double> barChartData = chart.ChartDataFactory.CreateBarChartData<string, double>();
            var serie = barChartData.AddSeries(categoryAxis, valueAxis);
            var serie2 = barChartData.AddSeries(categoryAxis, value2Axis);
            serie.SetTitle("生产总数量");
            serie2.SetTitle("生产合格数");
            chart.Plot(barChartData, bottomAxis, leftAxis);
        }

        private DataTable CreateTable()
        {
            var dt = new DataTable();
            dt.Columns.Add("月份");
            dt.Columns.Add("生产总数量", typeof(int));
            dt.Columns.Add("生产合格数", typeof(int));
            dt.Columns.Add("不良总数量", typeof(int));
            dt.Columns.Add("合格率");
            Random random = new Random();
            for (int i = 0; i < 12; i++)
            {
                var row = dt.NewRow();
                int num = random.Next(0, 10);
                row["月份"] = (i + 1).ToString();
                var totalCount = random.Next(100, 1000);
                var goodCount = random.Next(1, totalCount); ;
                row["生产总数量"] = totalCount;
                row["生产合格数"] = goodCount;
                row["不良总数量"] = totalCount - goodCount;
                row["合格率"] = string.Format("{0:0.00%}", (decimal)goodCount / totalCount);
                dt.Rows.Add(row);
            }
            return dt;
        }
    }
}
