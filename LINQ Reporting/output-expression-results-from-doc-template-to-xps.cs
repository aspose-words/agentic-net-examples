using System;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the DOC template.
        Document doc = new Document("Template.docx");

        // Example data source for the template.
        var data = new
        {
            Title = "Quarterly Report",
            Date = DateTime.Now,
            Total = 98765.43
        };

        // Populate the template using the ReportingEngine.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, data, "ds");

        // Configure XPS save options (optional settings).
        XpsSaveOptions xpsOptions = new XpsSaveOptions();
        // Include headings up to level 2 in the XPS outline.
        xpsOptions.OutlineOptions.HeadingsOutlineLevels = 2;

        // Save the populated document as XPS.
        doc.Save("Result.xps", xpsOptions);
    }
}
