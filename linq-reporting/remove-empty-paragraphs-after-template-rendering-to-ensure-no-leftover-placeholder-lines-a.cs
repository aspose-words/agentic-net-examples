using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;

#nullable enable

public class Program
{
    public static void Main()
    {
        // Paths for the template and the final document.
        string templatePath = "template.docx";
        string outputPath = "output.docx";

        // -----------------------------------------------------------------
        // 1. Create a template document with LINQ Reporting tags.
        // -----------------------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Normal paragraph.
        builder.Writeln("Report Header");
        builder.Writeln();

        // Paragraph with a required field.
        builder.Writeln("Customer: <<[model.CustomerName]>>");

        // Paragraph that may become empty after rendering.
        builder.Writeln("Notes: <<[model.OptionalNotes]>>");

        // Another normal paragraph.
        builder.Writeln("Thank you for your business.");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Load the template for report generation.
        // -----------------------------------------------------------------
        Document doc = new Document(templatePath);

        // -----------------------------------------------------------------
        // 3. Prepare the data model.
        // -----------------------------------------------------------------
        ReportModel model = new()
        {
            CustomerName = "Acme Corp",
            // OptionalNotes is left null to simulate an empty placeholder line.
            OptionalNotes = null
        };

        // -----------------------------------------------------------------
        // 4. Build the report using the LINQ Reporting engine.
        // -----------------------------------------------------------------
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // -----------------------------------------------------------------
        // 5. Remove empty paragraphs that may have been left after rendering.
        // -----------------------------------------------------------------
        foreach (Section section in doc.Sections)
        {
            Body body = section.Body;
            for (int i = body.Paragraphs.Count - 1; i >= 0; i--)
            {
                Paragraph para = body.Paragraphs[i];
                // Get the paragraph text (includes the paragraph break character).
                string text = para.GetText();
                // Remove control characters and check if the remaining text is whitespace.
                if (string.IsNullOrWhiteSpace(text.Replace("\r", string.Empty).Replace("\a", string.Empty)))
                {
                    para.Remove();
                }
            }
        }

        // -----------------------------------------------------------------
        // 6. Save the final document.
        // -----------------------------------------------------------------
        doc.Save(outputPath);
    }
}

// Data model used by the template.
public class ReportModel
{
    public string CustomerName { get; set; } = string.Empty;
    public string? OptionalNotes { get; set; }
}
