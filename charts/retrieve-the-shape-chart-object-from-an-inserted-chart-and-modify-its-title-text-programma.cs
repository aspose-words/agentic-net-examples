using System;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;
using System.Drawing;

namespace AsposeWordsChartTitleExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new empty document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a column chart into the document.
            Shape chartShape = builder.InsertChart(ChartType.Column, 432, 252);
            // Retrieve the Chart object from the inserted shape.
            Chart chart = chartShape.Chart;

            // Modify the chart title.
            ChartTitle title = chart.Title;
            title.Text = "Quarterly Sales";
            title.Font.Size = 14;
            title.Font.Color = Color.DarkBlue;
            title.Show = true; // Ensure the title is visible.

            // Save the document with the updated chart title.
            doc.Save("ChartWithTitle.docx");
        }
    }
}
