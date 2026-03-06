using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

class Program
{
    static void Main()
    {
        // Load the template document that contains the <<doc [src.Document]>> tags.
        Document template = new Document("Template.docx");

        // Load the source documents that will be inserted into the template.
        // Here we use LINQ to create an array of wrapper objects that expose a Document property.
        DocumentTestClass[] sources = new[]
        {
            "Doc1.docx",
            "Doc2.docx",
            "Doc3.docx"
        }
        .Select(path => new DocumentTestClass { Document = new Document(path) })
        .ToArray();

        // Build the report. The data source name "src" matches the tag used in the template.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(template, new object[] { sources }, new[] { "src" });

        // Save the populated document.
        template.Save("Result.docx");
    }
}

// Simple wrapper class required by the ReportingEngine to expose the document to insert.
class DocumentTestClass
{
    public Document Document { get; set; }
}
