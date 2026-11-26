using System;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;
using Aspose.Words.Tables;

namespace ChartCloneExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new document and a builder for it.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a chart into the first section.
            Shape originalChartShape = builder.InsertChart(ChartType.Column, 432, 252);
            Chart originalChart = originalChartShape.Chart;

            // Populate the chart with sample data.
            originalChart.Series.Clear();
            originalChart.Series.Add("Series 1", new double[] { 10, 20, 30 });
            originalChart.Series.Add("Series 2", new double[] { 15, 25, 35 });

            // Add a section break to create a second section.
            builder.InsertBreak(BreakType.SectionBreakNewPage);

            // Write a paragraph in the second section where the cloned chart will be inserted.
            builder.Writeln("This paragraph is in the second section. The cloned chart follows:");

            // Clone the original chart shape (deep clone).
            Shape clonedChartShape = (Shape)originalChartShape.Clone(true);

            // Insert the cloned chart shape into the current paragraph.
            builder.InsertNode(clonedChartShape);

            // Save the document.
            doc.Save("ClonedChart.docx");
        }
    }
}
