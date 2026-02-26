using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

class LinqReportingInsertDocuments
{
    static void Main()
    {
        // Path to the template document that contains the <<doc [src.Document]>> tag.
        string templatePath = @"C:\Docs\Template.docx";

        // Load the template document.
        Document template = new Document(templatePath);

        // Directory that holds the source DOCX files to be inserted.
        string sourceDocsDir = @"C:\Docs\Sources";

        // Use LINQ to create an array of Document objects from all DOCX files in the directory.
        Document[] sourceDocuments = Directory
            .EnumerateFiles(sourceDocsDir, "*.docx")
            .Select(file => new Document(file))
            .ToArray();

        // Create the reporting engine.
        ReportingEngine engine = new ReportingEngine();

        // Build the report, passing the source documents as a data source named "src".
        // The template must reference the documents with the tag <<doc [src.Document]>>.
        engine.BuildReport(template, sourceDocuments, new[] { "src" });

        // Save the populated document.
        string outputPath = @"C:\Docs\Result.docx";
        template.Save(outputPath);
    }
}
