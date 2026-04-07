using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Chapter
{
    // Sample chapter number
    public int ChapterNumber { get; set; } = 3;
    // Additional property to show in the report
    public string Title { get; set; } = "Sample Chapter Title";
}

public class Program
{
    public static void Main()
    {
        // Create a blank document and a builder to insert the LINQ Reporting tags.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a heading that will display the chapter number in lower‑case Roman numerals.
        // The expression uses the :roman format specifier.
        builder.Writeln("Chapter <<[model.ChapterNumber]:roman>>: <<[model.Title]>>");

        // Prepare the data source.
        Chapter model = new Chapter();

        // Build the report using the ReportingEngine.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the generated document.
        doc.Save("ChapterReport.docx");
    }
}
