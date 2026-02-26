using System;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // 1. Create a blank Word document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // 2. Insert a LINQ Reporting expression.
        // The expression calculates the sum of the Numbers array from the data source.
        builder.Writeln("Total: <<[data.Numbers].Sum()>>");

        // 3. Prepare the data source that will be referenced by the template.
        var data = new ReportData
        {
            Numbers = new[] { 10, 20, 30 }
        };

        // 4. Populate the template using ReportingEngine.
        // The third argument ("data") is the name used inside the template to refer to the object.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, data, "data");

        // 5. Save the resulting document as SVG.
        // SaveFormat.Svg tells Aspose.Words to render each page to a separate SVG file.
        doc.Save("Report.svg", SaveFormat.Svg);
    }

    // Simple POCO class that serves as the data source for the template.
    public class ReportData
    {
        public int[] Numbers { get; set; }
    }
}
