using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReporting
{
    // Simple data source class required by the LINQ Reporting engine.
    public class DocSource
    {
        public Document Document { get; set; }
    }

    public class Program
    {
        public static void Main()
        {
            // Load the source document that is in RTF format.
            Document rtfDocument = new Document("Source.rtf");

            // Create a template document that contains a LINQ Reporting placeholder.
            // The placeholder <<doc [src.Document]>> tells the engine to insert the Document
            // object from the data source named "src".
            Document template = new Document();
            DocumentBuilder builder = new DocumentBuilder(template);
            builder.Writeln("=== Report Start ===");
            builder.Writeln("<<doc [src.Document]>>");
            builder.Writeln("=== Report End ===");

            // Prepare the data source instance.
            DocSource src = new DocSource { Document = rtfDocument };

            // Build the report using the LINQ Reporting engine.
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(template, new object[] { src }, new string[] { "src" });

            // Save the resulting document.
            template.Save("Result.docx");
        }
    }
}
