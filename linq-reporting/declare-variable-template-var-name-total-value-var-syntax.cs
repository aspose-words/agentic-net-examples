using System;
using Aspose.Words;
using Aspose.Words.Reporting;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Use DocumentBuilder to write the template syntax that declares a variable.
        DocumentBuilder builder = new DocumentBuilder(doc);
        // Declare a variable named "Total" with the value 123.
        builder.Writeln("<<var name=\"Total\">>123<</var>>");
        // Use the declared variable later in the document.
        builder.Writeln("The total amount is <<[Total]>>.");

        // Build the report. No external data source is required for this example.
        ReportingEngine engine = new ReportingEngine();
        // The BuildReport overload requires a data source object; an empty anonymous object works.
        engine.BuildReport(doc, new { });

        // Save the resulting document.
        doc.Save("VariableDeclarationReport.docx");
    }
}
