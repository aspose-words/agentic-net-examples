using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingDemo
{
    // Simple class that will be used as a data source for the LINQ Reporting Engine.
    // The template will reference the property "Document" of this class.
    public class DocumentSource
    {
        public Document Document { get; set; }

        public DocumentSource(Document doc)
        {
            Document = doc;
        }
    }

    class Program
    {
        static void Main()
        {
            // -----------------------------------------------------------------
            // 1. Create a template document that contains DOT (<< >>) tags.
            //    The tags use the "doc" command to insert another document.
            //    The second tag uses the "-sourceNumbering" switch to keep the
            //    numbering of the inserted document separate from the template.
            // -----------------------------------------------------------------
            Document template = new Document();
            DocumentBuilder builder = new DocumentBuilder(template);

            // First insertion – default behavior (numbering continues).
            builder.Writeln("First insertion (numbering continues):");
            builder.Writeln("<<doc [src.Document]>>");
            builder.Writeln();

            // Second insertion – source numbering is kept as is.
            builder.Writeln("Second insertion (source numbering kept):");
            builder.Writeln("<<doc [src.Document] -sourceNumbering>>");
            builder.Writeln();

            // -----------------------------------------------------------------
            // 2. Load the document that will be inserted into the template.
            //    For demonstration we create a simple document on the fly,
            //    but any existing .docx file can be loaded here.
            // -----------------------------------------------------------------
            Document sourceDoc = new Document();
            DocumentBuilder srcBuilder = new DocumentBuilder(sourceDoc);
            srcBuilder.Writeln("List item 1.");
            srcBuilder.Writeln("List item 2.");
            srcBuilder.Writeln("List item 3.");

            // Ensure the source document has a numbered list so we can see the effect
            // of the "-sourceNumbering" switch.
            srcBuilder.ListFormat.ApplyNumberDefault();
            srcBuilder.Writeln("Numbered item A.");
            srcBuilder.Writeln("Numbered item B.");

            // -----------------------------------------------------------------
            // 3. Prepare the data source for the ReportingEngine.
            //    The engine expects an object (or an array of objects) that
            //    contains members referenced in the template.
            // -----------------------------------------------------------------
            DocumentSource dataSource = new DocumentSource(sourceDoc);

            // -----------------------------------------------------------------
            // 4. Build the report using the LINQ Reporting Engine.
            //    We pass the template, an array with a single data source object,
            //    and the corresponding name ("src") that matches the template tags.
            // -----------------------------------------------------------------
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(template, new object[] { dataSource }, new[] { "src" });

            // -----------------------------------------------------------------
            // 5. Save the resulting document.
            // -----------------------------------------------------------------
            string outputPath = Path.Combine(Environment.CurrentDirectory, "LinqReportingResult.docx");
            template.Save(outputPath);
            Console.WriteLine($"Report generated and saved to: {outputPath}");
        }
    }
}
