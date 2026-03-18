using System;
using Aspose.Words;
using Aspose.Words.Reporting;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert an expression tag that calls a static method using type‑member access syntax.
        // The syntax <<[Full.Type.Name]::StaticMethod()>> is interpreted by the ReportingEngine.
        builder.Writeln("<<[aspose.words.reporting.ReportingEngine]::GetRestrictedTypes()>>");

        // Build the report. The engine evaluates the expression tag and replaces it with the result.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, new object());

        // Save the resulting document.
        doc.Save("StaticMethodCallResult.docx");
    }
}
