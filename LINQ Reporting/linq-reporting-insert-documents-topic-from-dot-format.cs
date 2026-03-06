using System;
using Aspose.Words;
using Aspose.Words.Reporting;

class Program
{
    static void Main()
    {
        // 1. Create a template document that contains a placeholder for another document.
        Document template = new Document();
        DocumentBuilder templateBuilder = new DocumentBuilder(template);
        // The LINQ Reporting Engine uses the syntax <<doc [src.Document]>> to insert a document.
        templateBuilder.Writeln("<<doc [src.Document]>>");

        // 2. Create the document that will be inserted into the template.
        Document sourceDoc = new Document();
        DocumentBuilder sourceBuilder = new DocumentBuilder(sourceDoc);
        sourceBuilder.Writeln("This is the content of the inserted document.");

        // 3. Wrap the source document in an anonymous object that will be used as a data source.
        var dataSource = new { Document = sourceDoc };

        // 4. Populate the template using the ReportingEngine.
        ReportingEngine engine = new ReportingEngine();
        // The third argument ("src") is the name used in the template to reference the data source.
        engine.BuildReport(template, dataSource, "src");

        // 5. Save the final document.
        template.Save("Result.docx");
    }
}
