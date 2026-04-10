using System;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;

namespace AsposeChartsExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank Word document.
            Document doc = new Document();

            // Create a DocumentBuilder to insert content.
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a column chart with a width of 500 points and a height of 300 points.
            // The chart will contain Aspose.Words' default demo data.
            Shape chartShape = builder.InsertChart(ChartType.Column, 500, 300);

            // The chart can be accessed via chartShape.Chart if further customization is needed.
            // For this example we keep the default data unchanged.

            // Save the document to the working directory.
            doc.Save("ColumnChart.docx");
        }
    }
}
