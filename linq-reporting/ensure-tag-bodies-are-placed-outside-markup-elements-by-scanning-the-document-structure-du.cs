using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Register code page provider (required for some Aspose.Words features)
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Create sample data model
        ReportModel model = new ReportModel
        {
            BookmarkName = "MyBookmark",
            Title = "Sample Report"
        };

        // Create a template document programmatically
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // Paragraph with a bookmark tag whose body is inside the same run (needs preprocessing)
        builder.Writeln("<<bookmark [model.BookmarkName]>>" + model.Title + "<</bookmark>>");

        // Preprocess the document to ensure tag bodies are placed outside markup elements
        EnsureTagBodiesOutsideMarkup(template);

        // Build the report
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(template, model, "model");

        // Save the generated report
        string outputPath = Path.Combine(Environment.CurrentDirectory, "ReportOutput.docx");
        template.Save(outputPath);
    }

    // Scans the document and splits runs that contain both opening and closing tags,
    // placing the tag bodies in separate runs outside the markup tags.
    private static void EnsureTagBodiesOutsideMarkup(Document doc)
    {
        foreach (Paragraph paragraph in doc.GetChildNodes(NodeType.Paragraph, true))
        {
            // Iterate over runs; use index because we may insert new runs during iteration
            for (int i = 0; i < paragraph.Runs.Count; i++)
            {
                Run run = (Run)paragraph.Runs[i];
                string text = run.Text;

                // Check for a bookmark tag that contains both opening and closing parts in the same run
                if (text.Contains("<<bookmark") && text.Contains("<</bookmark>>"))
                {
                    int openingEnd = text.IndexOf(">>", StringComparison.Ordinal) + 2;
                    string openingTag = text.Substring(0, openingEnd);

                    int closingStart = text.IndexOf("<</bookmark>>", StringComparison.Ordinal);
                    string body = text.Substring(openingEnd, closingStart - openingEnd);
                    string closingTag = text.Substring(closingStart);

                    // Replace current run with the opening tag
                    run.Text = openingTag;

                    // Insert body run
                    Run bodyRun = new Run(doc, body);
                    paragraph.Runs.Insert(i + 1, bodyRun);

                    // Insert closing tag run
                    Run closingRun = new Run(doc, closingTag);
                    paragraph.Runs.Insert(i + 2, closingRun);

                    // Skip over the newly inserted runs
                    i += 2;
                }
            }
        }
    }
}

// Public data model used by the LINQ Reporting engine
public class ReportModel
{
    public string BookmarkName { get; set; } = "DefaultBookmark";
    public string Title { get; set; } = "Default Title";
}
