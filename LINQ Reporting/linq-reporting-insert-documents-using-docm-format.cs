using System;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the DOCM template that contains LINQ Reporting tags.
        Document template = new Document("Template.docm");

        // Example data source for the reporting engine.
        var reportData = new
        {
            Title = "Quarterly Report",
            GeneratedOn = DateTime.Now,
            Author = "John Doe"
        };

        // Populate the template with data using the ReportingEngine.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(template, reportData, "ds");

        // Load the additional DOCM document that will be inserted.
        Document docToInsert = new Document("Appendix.docm");

        // Insert the additional document at the end of the populated template.
        DocumentBuilder builder = new DocumentBuilder(template);
        builder.MoveToDocumentEnd();
        builder.InsertDocument(docToInsert, ImportFormatMode.KeepSourceFormatting);

        // Save the final document in DOCM format.
        template.Save("Result.docm", SaveFormat.Docm);
    }
}
