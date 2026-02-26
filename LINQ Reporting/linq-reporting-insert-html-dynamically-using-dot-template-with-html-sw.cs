using System;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Model
{
    // Property that holds the HTML fragment to be inserted.
    public string Html { get; set; }
}

public class Program
{
    public static void Main()
    {
        // -----------------------------------------------------------------
        // 1. Create a template document in memory.
        // -----------------------------------------------------------------
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // Insert a LINQ Reporting placeholder with the HTML switch.
        // The ":html" switch tells the engine to treat the value as HTML.
        builder.Writeln("<<[model.Html]:html>>");

        // -----------------------------------------------------------------
        // 2. Prepare the data source.
        // -----------------------------------------------------------------
        var data = new Model
        {
            Html = "<p style='color:red;'>This is <b>red</b> paragraph inserted via HTML.</p>"
        };

        // -----------------------------------------------------------------
        // 3. Build the report using ReportingEngine.
        // -----------------------------------------------------------------
        ReportingEngine engine = new ReportingEngine();

        // The third argument ("model") is the name used in the template to reference the data source.
        engine.BuildReport(template, data, "model");

        // -----------------------------------------------------------------
        // 4. Save the generated document.
        // -----------------------------------------------------------------
        template.Save("ReportWithHtml.docx");
    }
}
