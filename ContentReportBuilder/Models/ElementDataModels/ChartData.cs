
using System.Collections.Generic;

namespace ContentReportBuilder.Models.ElementDataModels
{
    /// <summary>
    /// Chart element model data
    /// </summary>
    public class ChartData
    {
        public string Title { get; set; }
        public List<ChartSeries> Series { get; set; }
    }

    public class ChartSeries
    {
        public ChartSeries()
        {

        }
        public ChartSeries(string title, List<SeriesItem> items)
        {
            Title = title;
            Items = items;
        }

        public ChartSeries(string title, double barValue)
        {
            Title = title;
            BarValue = barValue;
        }
        public List<SeriesItem> Items { get; set; }
        public double BarValue { get; set; }
        public string Title { get; set; }

        public bool IsExploded { get; set; }
    }

    public class SeriesItem
    {
        public SeriesItem()
        {

        }
        public SeriesItem(double value)
        {
            X = value;
        }

        public SeriesItem(double x, double y)
        {
            X = x;
            Y = y;
        }
        public double X { get; set; }
        public double Y { get; set; }
        public double Value => X;
    }
}
