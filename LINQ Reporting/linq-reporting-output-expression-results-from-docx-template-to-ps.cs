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

        // 2. Insert a LINQ Reporting expression that will be evaluated.
        //    The template tag <<[ds.Value]>> references the property "Value" of the data source named "ds".
        builder.Writeln("Report value: <<[ds.Value]>>");

        // 3. Prepare the data source. An anonymous object with a single property is sufficient.
        var dataSource = new { Value = 12345.67 };

        // 4. Populate the template using the ReportingEngine.
        ReportingEngine engine = new ReportingEngine();
        // The third argument is the name used inside the template to reference the data source.
        engine.BuildReport(doc, dataSource, "ds");

        // 5. Save the resulting document as PostScript (PS) using the provided save options.
        PsSaveOptions psOptions = new PsSaveOptions
        {
            SaveFormat = SaveFormat.Ps   // Explicitly set the format to PS.
        };
        doc.Save("ReportOutput.ps", psOptions);
    }
}
