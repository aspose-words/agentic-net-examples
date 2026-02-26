using System;
using Aspose.Words;
using Aspose.Words.Reporting;

class Program
{
    static void Main()
    {
        // -----------------------------------------------------------------
        // 1. Create a DOCX template that contains a Reporting Engine placeholder
        //    with the :html switch. The switch tells the engine to treat the
        //    supplied value as HTML and insert it accordingly.
        // -----------------------------------------------------------------
        Document template = new Document();                     // create blank document
        DocumentBuilder builder = new DocumentBuilder(template);
        builder.Writeln("<<[ReportData.Html]:html>>");          // placeholder with html switch
        string templatePath = "Template.docx";
        template.Save(templatePath);                            // save the template

        // -----------------------------------------------------------------
        // 2. Load the template that we just created.
        // -----------------------------------------------------------------
        Document doc = new Document(templatePath);              // load existing document

        // -----------------------------------------------------------------
        // 3. Prepare a data source object. The property name (Html) matches
        //    the placeholder field name used in the template.
        // -----------------------------------------------------------------
        var data = new ReportData
        {
            Html = "<h1 style='color:blue;'>Hello Aspose</h1>" +
                   "<p>This is <b>bold</b> and <i>italic</i> text.</p>"
        };

        // -----------------------------------------------------------------
        // 4. Populate the template using the ReportingEngine. The third
        //    argument ("ReportData") is the name used to reference the data
        //    source inside the template.
        // -----------------------------------------------------------------
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, data, "ReportData");

        // -----------------------------------------------------------------
        // 5. Save the resulting document which now contains the rendered HTML.
        // -----------------------------------------------------------------
        string outputPath = "Result.docx";
        doc.Save(outputPath);
    }

    // Simple POCO class that serves as the data source for the report.
    public class ReportData
    {
        public string Html { get; set; }
    }
}
