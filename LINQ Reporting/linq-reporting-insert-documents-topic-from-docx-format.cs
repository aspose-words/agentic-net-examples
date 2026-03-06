using System;
using System.IO;
using System.Linq;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

class Program
{
    static void Main()
    {
        // Load the template that contains the <<doc [src.Document]>> tags.
        Document template = new Document("Template.docx");

        // Load all DOCX files from a folder, create Document objects and filter them with LINQ.
        string sourceFolder = "Sources";
        Document[] sourceDocs = Directory.GetFiles(sourceFolder, "*.docx")
            .Select(path => new Document(path))               // create a Document for each file
            .Where(doc => doc.Range.Text.Contains("InsertMe")) // example LINQ filter
            .ToArray();

        // Prepare an array of empty names – the BuildReport overload requires a name for each source.
        string[] sourceNames = Enumerable.Repeat(string.Empty, sourceDocs.Length).ToArray();

        // Use the ReportingEngine to insert the filtered documents into the template.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(template, sourceDocs, sourceNames);

        // Save the final document.
        template.Save("Result.docx");
    }
}
