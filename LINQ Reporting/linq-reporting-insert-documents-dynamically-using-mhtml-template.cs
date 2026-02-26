using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsMhtmlInsert
{
    // Simple wrapper class that will be referenced from the template.
    public class DocumentSource
    {
        // The document that should be inserted into the template.
        public Document Document { get; set; }
    }

    class Program
    {
        static void Main()
        {
            // Path to the MHTML template that contains a LINQ Reporting tag,
            // e.g. <<doc [src.Document]>> where "src" is the name we will give to the data source.
            string templatePath = @"C:\Templates\ReportTemplate.mht";

            // Load the MHTML template. Aspose.Words automatically detects the format.
            Document template = new Document(templatePath);

            // Load the document that we want to insert dynamically.
            string insertDocPath = @"C:\Documents\InsertDocument.docx";
            Document insertDoc = new Document(insertDocPath);

            // Prepare the data source object. The template will access src.Document.
            DocumentSource src = new DocumentSource { Document = insertDoc };

            // Create the reporting engine.
            ReportingEngine engine = new ReportingEngine();

            // Build the report. The third argument is the name used inside the template.
            engine.BuildReport(template, src, "src");

            // Save the final document.
            string resultPath = @"C:\Output\FinalReport.docx";
            template.Save(resultPath);
        }
    }
}
