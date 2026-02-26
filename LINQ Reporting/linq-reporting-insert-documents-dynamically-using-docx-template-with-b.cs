using System;
using Aspose.Words;
using Aspose.Words.Reporting;

class LinqReportingExample
{
    static void Main()
    {
        // 1. Create a DOCX template that contains the build‑switch tags.
        //    The tags tell the ReportingEngine where to insert the source document.
        Document template = new Document();                     // create blank document
        DocumentBuilder builder = new DocumentBuilder(template);
        builder.Writeln("<<doc [src.Document]>>");               // normal insertion
        builder.Writeln("<<doc [src.Document] -sourceNumbering>>"); // keep source numbering

        // 2. Load the document that will be inserted dynamically.
        //    Replace the path with the actual location of your source DOCX file.
        Document sourceDoc = new Document("Source.docx");

        // 3. Prepare the data source for the ReportingEngine.
        //    The engine expects an array of objects and a matching array of names.
        object[] dataSources = new object[] { sourceDoc };
        string[] dataSourceNames = new string[] { "src" };

        // 4. Build the report – the engine parses the template and replaces the tags
        //    with the contents of the source document.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(template, dataSources, dataSourceNames);

        // 5. Save the populated document.
        template.Save("Result.docx");
    }
}
