using System;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;

class ReplaceChartDataSource
{
    static void Main()
    {
        // Path to the input DOCX file that contains the chart.
        string inputPath = @"C:\Docs\InputDocument.docx";

        // Path to the new Excel file that will become the chart's data source.
        string newDataSourcePath = @"C:\Data\NewChartData.xlsx";

        // The exact title text of the chart we want to modify.
        string targetChartTitle = "Sales Overview";

        // Load the existing Word document.
        Document doc = new Document(inputPath);

        // Iterate through all Shape nodes in the document (including those inside headers/footers).
        foreach (Shape shape in doc.GetChildNodes(NodeType.Shape, true))
        {
            // Ensure the shape actually contains a chart.
            Chart chart = shape.Chart;
            if (chart == null)
                continue;

            // Check if the chart's title matches the target title.
            if (chart.Title != null && chart.Title.Text == targetChartTitle)
            {
                // Replace the chart's linked data source with the new Excel file.
                chart.SourceFullName = newDataSourcePath;

                // If the chart is linked, you may need to update it to reflect the new data.
                // This can be done by calling UpdateChartData() if such a method exists,
                // otherwise the new source will be used the next time the document is opened.
                break; // Assuming titles are unique; exit after the first match.
            }
        }

        // Save the modified document to a new file.
        string outputPath = @"C:\Docs\OutputDocument.docx";
        doc.Save(outputPath);
    }
}
