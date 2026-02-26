using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

class Program
{
    static void Main()
    {
        // Load the DOCM template that contains the <<doc [src.Document] -sourceStyles>> tag.
        Document template = new Document("Template.docm");

        // Prepare the documents that will be inserted dynamically.
        // Each document is wrapped in a simple holder class so that the template can reference it as [src.Document].
        List<DocHolder> sources = new List<DocHolder>
        {
            new DocHolder { Document = new Document("Part1.docx") },
            new DocHolder { Document = new Document("Part2.docx") },
            new DocHolder { Document = new Document("Part3.docx") }
        };

        // Create the reporting engine.
        ReportingEngine engine = new ReportingEngine();

        // Build the report. The array overload allows us to pass the data source and its name ("src")
        // which matches the placeholder used in the template.
        engine.BuildReport(template, new object[] { sources }, new string[] { "src" });

        // Save the populated document.
        template.Save("Result.docx");
    }

    // Simple class used as a data source for the template.
    public class DocHolder
    {
        public Document Document { get; set; }
    }
}
