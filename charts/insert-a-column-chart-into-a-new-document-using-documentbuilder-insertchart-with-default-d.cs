using System;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;

namespace InsertColumnChartExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new empty document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a column chart with default data.
            Shape chartShape = builder.InsertChart(ChartType.Column, 432, 252);
            Chart chart = chartShape.Chart; // Access chart (optional, default data already present).

            // Save the document.
            doc.Save("insert-chart.docx");
        }
    }
}
