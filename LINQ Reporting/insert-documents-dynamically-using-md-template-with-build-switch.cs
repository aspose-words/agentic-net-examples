using System;
using Aspose.Words;
using Aspose.Words.Reporting;

class Program
{
    static void Main()
    {
        // Load the main template that contains the build switch (e.g., <<doc [src.Document]>>).
        Document template = new Document("Template.docx");

        // Load the document that will be inserted dynamically.
        Document sourceDocument = new Document("Source.docx");

        // Create an anonymous object that holds the source document.
        // The property name ("Document") matches the field used in the template.
        var dataSource = new { Document = sourceDocument };

        // Build the report. The third parameter ("src") is the name used in the template
        // to reference the data source object itself (<<doc [src.Document]>>).
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(template, dataSource, "src");

        // Save the populated document.
        template.Save("Result.docx");
    }
}
