using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Prepare folders for the document and the extracted SVG files.
        string dataFolder = Path.Combine(Directory.GetCurrentDirectory(), "Data");
        Directory.CreateDirectory(dataFolder);
        string outputFolder = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputFolder);

        // -----------------------------------------------------------------
        // 1. Create a Word document that contains a chart.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a column chart with a deterministic size.
        Shape chartShape = builder.InsertChart(ChartType.Column, 400, 300);

        // Add a simple title to the chart (optional, just to have content).
        chartShape.Chart.Title.Text = "Sample Chart";

        // Save the document (optional, demonstrates the source file).
        string docPath = Path.Combine(dataFolder, "ChartDocument.docx");
        doc.Save(docPath);

        // -----------------------------------------------------------------
        // 2. Extract each chart shape and export it as an SVG file.
        // -----------------------------------------------------------------
        int chartIndex = 0;
        foreach (Shape shape in doc.GetChildNodes(NodeType.Shape, true))
        {
            if (shape.HasChart)
            {
                string svgFileName = Path.Combine(outputFolder, $"Chart_{chartIndex}.svg");

                // Configure SVG save options – we do not embed images because the chart is vector.
                SvgSaveOptions svgOptions = new SvgSaveOptions
                {
                    ExportEmbeddedImages = false,
                    ResourcesFolder = outputFolder,
                    ResourcesFolderAlias = outputFolder,
                    ShowPageBorder = false
                };

                // Render the chart shape directly to an SVG file.
                shape.GetShapeRenderer().Save(svgFileName, svgOptions);
                chartIndex++;
            }
        }

        // -----------------------------------------------------------------
        // 3. Validate that at least one SVG file was produced.
        // -----------------------------------------------------------------
        if (chartIndex == 0)
            throw new InvalidOperationException("No chart shapes were found in the document.");

        Console.WriteLine($"Extracted {chartIndex} chart(s) as SVG files to \"{outputFolder}\".");
    }
}
