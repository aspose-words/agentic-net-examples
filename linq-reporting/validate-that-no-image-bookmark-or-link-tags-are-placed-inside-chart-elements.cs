using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;

public class Program
{
    public static void Main()
    {
        // Register code page provider for Aspose.Words.
        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

        const string templatePath = "Template.docx";

        // -------------------- Create template --------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a chart into the document.
        Shape chartShape = builder.InsertChart(ChartType.Column, 400, 300);

        // Intentionally place an invalid image tag inside the chart title to demonstrate validation.
        if (chartShape.HasChart && chartShape.Chart.Title != null)
        {
            chartShape.Chart.Title.Text = "<<image [model.ImagePath]>>";
        }

        // Save the template to disk.
        doc.Save(templatePath);

        // -------------------- Load template and validate --------------------
        Document loadedDoc = new Document(templatePath);
        bool hasInvalidTags = ValidateChartTags(loadedDoc, out List<string> messages);

        // Output validation result.
        if (hasInvalidTags)
        {
            Console.WriteLine("Invalid tags found inside chart elements:");
            foreach (string msg in messages)
            {
                Console.WriteLine(msg);
            }
        }
        else
        {
            Console.WriteLine("No invalid image, bookmark, or link tags inside chart elements.");
        }
    }

    private static bool ValidateChartTags(Document doc, out List<string> messages)
    {
        messages = new List<string>();

        // Find all shape nodes in the document.
        NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);
        foreach (Shape shape in shapeNodes)
        {
            if (shape.HasChart)
            {
                // Check the chart title for invalid tags.
                Chart chart = shape.Chart;
                if (chart.Title != null && ContainsInvalidTag(chart.Title.Text))
                {
                    messages.Add($"Chart '{shape.Name}' contains invalid tag in title: {chart.Title.Text.Trim()}");
                }

                // Examine all paragraphs inside the chart shape.
                NodeCollection paragraphs = shape.GetChildNodes(NodeType.Paragraph, true);
                foreach (Paragraph para in paragraphs)
                {
                    foreach (Run run in para.Runs)
                    {
                        string text = run.Text;
                        if (ContainsInvalidTag(text))
                        {
                            messages.Add($"Chart '{shape.Name}' contains invalid tag: {text.Trim()}");
                        }
                    }
                }
            }
        }

        return messages.Count > 0;
    }

    private static bool ContainsInvalidTag(string text)
    {
        if (string.IsNullOrEmpty(text))
            return false;

        return text.Contains("<<image") ||
               text.Contains("<<bookmark") ||
               text.Contains("<<link");
    }
}
