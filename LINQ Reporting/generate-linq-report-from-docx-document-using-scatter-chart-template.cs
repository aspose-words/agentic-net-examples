using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Drawing.Charts;

namespace AsposeWordsScatterChartReport
{
    // Simple POCO representing a data point for the scatter chart.
    public class DataPoint
    {
        public double X { get; set; }
        public double Y { get; set; }

        public DataPoint(double x, double y)
        {
            X = x;
            Y = y;
        }
    }

    class Program
    {
        static void Main()
        {
            // 1. Load the DOCX template that contains a scatter chart placeholder.
            //    The template should have a chart shape with a tag that the ReportingEngine can bind to,
            //    e.g. <<chart [ds.Points]>> or any custom syntax supported by Aspose.Words Reporting Engine.
            Document template = new Document("ScatterChartTemplate.docx");

            // 2. Prepare the LINQ data source.
            //    Here we create a list of DataPoint objects that will be used to populate the chart.
            List<DataPoint> points = new List<DataPoint>
            {
                new DataPoint(1.2, 3.4),
                new DataPoint(2.5, 5.1),
                new DataPoint(3.8, 2.9),
                new DataPoint(4.0, 4.5),
                new DataPoint(5.3, 6.2)
            };

            // The reporting engine expects an object with properties that can be referenced from the template.
            // We wrap the list in an anonymous object (or a named class) and give it a name "ds".
            var dataSource = new { Points = points };

            // 3. Build the report using the ReportingEngine.
            ReportingEngine engine = new ReportingEngine();
            // The third parameter is the name used inside the template to reference the data source.
            engine.BuildReport(template, dataSource, "ds");

            // 4. Save the populated document.
            template.Save("ScatterChartReport.docx");
        }
    }
}
