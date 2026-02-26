using System;
using Aspose.Words;
using Aspose.Words.Reporting;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a LINQ Reporting placeholder that will be replaced by the TXT data source.
        builder.Writeln("<<[txt.Content]>>");

        // Prepare a simple TXT‑style data source.
        TxtDataSource txtData = new TxtDataSource("Hello from a TXT data source!");

        // Build the report using the ReportingEngine.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, txtData, "txt");

        // Save the generated document.
        doc.Save("ReportFromTxt.docx");
    }
}

// Minimal class representing a TXT data source with a single property.
public class TxtDataSource
{
    public string Content { get; set; }

    public TxtDataSource(string content)
    {
        Content = content;
    }
}
