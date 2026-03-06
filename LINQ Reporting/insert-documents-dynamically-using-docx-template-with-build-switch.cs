using System;
using Aspose.Words;
using Aspose.Words.Reporting;

class Program
{
    static void Main()
    {
        // Load the DOCX template that contains the build switch tags, e.g.
        // <<doc [src.Document]>> and/or <<doc [src.Document] -sourceNumbering>>
        Document template = new Document("Template.docx");

        // Load the document that will be inserted dynamically.
        Document sourceDoc = new Document("Source.docx");

        // Wrap the source document in a simple holder class.
        // The template will reference the holder via the name "src".
        var srcHolder = new DocumentHolder { Document = sourceDoc };

        // Create the reporting engine and populate the template.
        // The array of data sources contains the holder, and the corresponding
        // name array provides the identifier used in the template.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(template, new object[] { srcHolder }, new string[] { "src" });

        // Save the resulting document.
        template.Save("Result.docx");
    }
}

// Simple class exposing a Document property for use in the template.
public class DocumentHolder
{
    public Document Document { get; set; }
}
