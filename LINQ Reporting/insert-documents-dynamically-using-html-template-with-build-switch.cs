using System;
using Aspose.Words;
using Aspose.Words.Reporting;

class Program
{
    static void Main()
    {
        // -----------------------------------------------------------------
        // 1. Create the source document that will be inserted dynamically.
        // -----------------------------------------------------------------
        Document srcDoc = new Document();
        DocumentBuilder srcBuilder = new DocumentBuilder(srcDoc);
        srcBuilder.Writeln("First paragraph of the source document.");
        srcBuilder.Writeln("Second paragraph of the source document.");
        // Save the source document (optional, just for inspection).
        srcDoc.Save("Source.docx");

        // -----------------------------------------------------------------
        // 2. Create a template document that contains ReportingEngine tags.
        //    The tags use the <<doc [src.Document]>> syntax.
        //    The second tag includes the -sourceNumbering switch to keep the
        //    original numbering of the inserted document.
        // -----------------------------------------------------------------
        Document template = new Document();
        DocumentBuilder tmplBuilder = new DocumentBuilder(template);

        tmplBuilder.Writeln("=== Before first insertion ===");
        tmplBuilder.Writeln("<<doc [src.Document]>>"); // default behavior (continue numbering)
        tmplBuilder.Writeln("=== After first insertion ===");
        tmplBuilder.Writeln();

        tmplBuilder.Writeln("=== Before second insertion (preserve numbering) ===");
        tmplBuilder.Writeln("<<doc [src.Document] -sourceNumbering>>"); // keep source numbering
        tmplBuilder.Writeln("=== After second insertion ===");

        // Save the template document (optional, just for inspection).
        template.Save("Template.docx");

        // -----------------------------------------------------------------
        // 3. Populate the template using ReportingEngine.
        //    The data source name "src" matches the tag in the template.
        // -----------------------------------------------------------------
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(template, new object[] { srcDoc }, new string[] { "src" });

        // -----------------------------------------------------------------
        // 4. Save the final document that contains the inserted content.
        // -----------------------------------------------------------------
        template.Save("Result.docx");
    }
}
