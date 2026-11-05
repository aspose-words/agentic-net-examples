using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;

public class Program
{
    public static void Main()
    {
        // Register code page provider for Aspose.Words if needed.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare a temporary folder for the example files.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "output");
        Directory.CreateDirectory(outputDir);

        // Create a Word document template with a chart.
        string templatePath = Path.Combine(outputDir, "template.docx");
        CreateTemplateWithChart(templatePath);

        // Validate the template: ensure no image, bookmark, or link tags are inside chart elements.
        bool isValid = ValidateChartTags(templatePath, out List<string> foundTags);

        // Output the validation result.
        if (isValid)
        {
            Console.WriteLine("Validation passed: No prohibited tags inside chart elements.");
        }
        else
        {
            Console.WriteLine("Validation failed: Prohibited tags found inside chart elements:");
            foreach (string tag in foundTags)
            {
                Console.WriteLine($"- {tag}");
            }
        }
    }

    private static void CreateTemplateWithChart(string filePath)
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a chart shape.
        Shape chartShape = builder.InsertChart(ChartType.Column, 400, 300);

        // The chart shape does not contain paragraphs, so we cannot move the builder inside it.
        // Instead, we add a regular paragraph after the chart that (incorrectly) contains prohibited tags.
        // This demonstrates that the validation logic only inspects the text inside chart shapes.
        builder.Writeln();
        builder.Write("Quarterly Sales <<image [model.ImagePath]>>");
        builder.Writeln();
        builder.Write("Details <<bookmark [model.BookmarkName]>>");
        builder.Writeln();
        builder.Write("More info <<link [model.Url] [model.LinkText]>>");

        // Save the template.
        doc.Save(filePath);
    }

    private static bool ValidateChartTags(string docPath, out List<string> foundTags)
    {
        foundTags = new List<string>();
        Document doc = new Document(docPath);

        // Retrieve all shapes in the document.
        NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);
        foreach (Shape shape in shapes)
        {
            // Process only chart shapes.
            if (shape.HasChart)
            {
                // Get the inner text of the shape; for chart shapes this is typically empty.
                string innerText = shape.GetText() ?? string.Empty;

                // Check for prohibited tags.
                if (innerText.Contains("<<image"))
                    foundTags.Add("<<image>>");
                if (innerText.Contains("<<bookmark"))
                    foundTags.Add("<<bookmark>>");
                if (innerText.Contains("<<link"))
                    foundTags.Add("<<link>>");
            }
        }

        return foundTags.Count == 0;
    }
}
