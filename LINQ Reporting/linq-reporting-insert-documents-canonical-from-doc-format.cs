using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

public class DocumentTestClass
{
    public Document Document { get; set; }

    public DocumentTestClass(Document doc)
    {
        Document = doc;
    }
}

public class ReportingInsertExample
{
    public static void Main()
    {
        // Load the template that contains the <<doc [src.Document]>> tag.
        Document template = new Document("Template.docx");

        // Get all DOCX files from the source folder, filter with LINQ (e.g., only files starting with "Part").
        string sourceFolder = "SourceDocs";
        List<DocumentTestClass> sources = Directory.GetFiles(sourceFolder, "*.docx")
            .Where(f => Path.GetFileName(f).StartsWith("Part", StringComparison.OrdinalIgnoreCase))
            .Select(f => new DocumentTestClass(new Document(f)))
            .ToList();

        // Prepare arrays for the ReportingEngine.
        object[] dataSources = sources.Cast<object>().ToArray();
        string[] dataSourceNames = new[] { "src" }; // name used in the template tag

        // Build the report, inserting each document at the tag location.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(template, dataSources, dataSourceNames);

        // Save the resulting document.
        template.Save("Result.docx");
    }
}
