using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReporting
{
    // Simple class that will be used as a data source for the ReportingEngine.
    // The template will reference the property "Document" via the name "src".
    public class DocumentTestClass
    {
        public Document Document { get; }

        public DocumentTestClass(Document document)
        {
            Document = document;
        }
    }

    public class Program
    {
        public static void Main()
        {
            // -----------------------------------------------------------------
            // 1. Create a template document that contains a LINQ Reporting tag.
            //    The tag "<<doc [src.Document]>>" tells the engine to insert the
            //    document referenced by the data source named "src".
            // -----------------------------------------------------------------
            Document template = new Document();
            DocumentBuilder templateBuilder = new DocumentBuilder(template);
            templateBuilder.Writeln("Report start");
            templateBuilder.Writeln("<<doc [src.Document]>>"); // build switch
            templateBuilder.Writeln("Report end");

            // Save the template so it can be loaded later.
            template.Save("Template.docx");

            // -----------------------------------------------------------------
            // 2. Create a source document that will be inserted into the template.
            // -----------------------------------------------------------------
            Document source = new Document();
            DocumentBuilder sourceBuilder = new DocumentBuilder(source);
            sourceBuilder.Writeln("This is the dynamically inserted document.");
            sourceBuilder.Writeln("It can contain any content you need.");
            source.Save("Source.docx");

            // -----------------------------------------------------------------
            // 3. Prepare the data source for the ReportingEngine.
            //    The engine expects an object (or array of objects) whose members
            //    can be referenced from the template. Here we expose the source
            //    document via the property "Document".
            // -----------------------------------------------------------------
            DocumentTestClass dataSource = new DocumentTestClass(source);

            // -----------------------------------------------------------------
            // 4. Load the template document (lifecycle: load) and build the report.
            // -----------------------------------------------------------------
            Document report = new Document("Template.docx");

            ReportingEngine engine = new ReportingEngine
            {
                // Optional: remove empty paragraphs that may appear after processing.
                Options = ReportBuildOptions.RemoveEmptyParagraphs
            };

            // BuildReport overload that accepts multiple data sources.
            // The name "src" matches the tag used in the template.
            engine.BuildReport(report, new object[] { dataSource }, new[] { "src" });

            // -----------------------------------------------------------------
            // 5. Save the final report (lifecycle: save).
            // -----------------------------------------------------------------
            report.Save("FinalReport.docx");
        }
    }
}
