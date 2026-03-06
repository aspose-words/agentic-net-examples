using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsDemo
{
    // Simple class that will be used as the data source for the template.
    // The template will reference the property "Document" via the build switch.
    public class DocumentContainer
    {
        public Document Document { get; set; }
    }

    class Program
    {
        static void Main()
        {
            // Load the TXT template that contains the build switch, e.g.:
            // <<doc [src.Document]>>   (or with additional switches like -sourceNumbering)
            Document template = new Document("Template.txt");

            // Create the document that we want to insert dynamically.
            Document sourceDoc = new Document();
            DocumentBuilder srcBuilder = new DocumentBuilder(sourceDoc);
            srcBuilder.Writeln("First inserted paragraph.");
            srcBuilder.Writeln("Second inserted paragraph.");

            // Wrap the source document in a container object so the template can reference it.
            DocumentContainer dataSource = new DocumentContainer { Document = sourceDoc };

            // Use ReportingEngine to populate the template.
            // The third argument ("src") is the name used inside the template to reference the data source.
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(template, dataSource, "src");

            // Save the resulting document.
            template.Save("Result.docx");
        }
    }
}
