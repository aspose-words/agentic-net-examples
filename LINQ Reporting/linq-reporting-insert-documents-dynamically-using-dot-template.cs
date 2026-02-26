using System;
using Aspose.Words;
using Aspose.Words.Reporting;

class Program
{
    static void Main()
    {
        // -----------------------------------------------------------------
        // 1. Create a DOT (template) document that contains a placeholder
        //    for inserting other documents dynamically.
        // -----------------------------------------------------------------
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        builder.Writeln("=== Report Start ===");
        // The <<foreach>> tag iterates over the collection named "src".
        // Inside the loop we use <<doc [src.Document]>> to insert each document.
        builder.Writeln("<<foreach [src]>><<doc [src.Document]>>\n<</foreach>>");
        builder.Writeln("=== Report End ===");

        // -----------------------------------------------------------------
        // 2. Load the documents that will be inserted into the report.
        // -----------------------------------------------------------------
        Document docA = new Document("SubDocA.docx"); // replace with actual path
        Document docB = new Document("SubDocB.docx"); // replace with actual path

        // -----------------------------------------------------------------
        // 3. Prepare a data source that the ReportingEngine can iterate.
        //    Each item must expose a property named "Document" (type Document).
        // -----------------------------------------------------------------
        var dataSource = new[]
        {
            new DocumentHolder { Document = docA },
            new DocumentHolder { Document = docB }
        };

        // -----------------------------------------------------------------
        // 4. Build the report using the ReportingEngine.
        //    The data source name "src" matches the name used in the template.
        // -----------------------------------------------------------------
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(template, dataSource, "src");

        // -----------------------------------------------------------------
        // 5. Save the generated report.
        // -----------------------------------------------------------------
        template.Save("GeneratedReport.docx");
    }

    // Simple container class used as a data source item.
    public class DocumentHolder
    {
        public Document Document { get; set; }
    }
}
