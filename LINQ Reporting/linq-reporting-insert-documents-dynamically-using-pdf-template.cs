using System;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Saving;
using Aspose.Words.Drawing;

namespace AsposeWordsLinqReporting
{
    // Simple data source class that holds a document to be inserted.
    public class DocumentDataSource
    {
        public Document Document { get; set; }
    }

    class Program
    {
        static void Main()
        {
            // Path to the PDF template that contains a LINQ Reporting tag, e.g. <<doc [src.Document]>>.
            const string templatePath = @"C:\Templates\ReportTemplate.pdf";

            // Path to the document that will be inserted into the template.
            const string insertDocPath = @"C:\Documents\Appendix.docx";

            // Load the PDF template as an Aspose.Words Document.
            Document template = new Document(templatePath);

            // Load the document that will be inserted.
            Document docToInsert = new Document(insertDocPath);

            // Prepare the data source. The template can reference the property "Document"
            // of the data source object using the name we will assign ("src").
            var dataSource = new DocumentDataSource { Document = docToInsert };

            // Create the ReportingEngine and build the report.
            ReportingEngine engine = new ReportingEngine();

            // BuildReport overload that accepts an array of data sources and their names.
            // This allows the template to reference the data source object itself.
            engine.BuildReport(template,
                new object[] { dataSource },               // data sources
                new[] { "src" });                         // corresponding names

            // Save the resulting document as PDF.
            const string outputPath = @"C:\Output\FinalReport.pdf";
            template.Save(outputPath, SaveFormat.Pdf);
        }
    }
}
