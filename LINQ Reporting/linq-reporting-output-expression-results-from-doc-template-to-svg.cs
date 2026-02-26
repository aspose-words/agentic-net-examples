using System;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // 1. Create a blank Word document and insert a LINQ Reporting placeholder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Result: <<[ds.Value]>>");   // <<[ds.Value]>> will be replaced by the data source.

        // 2. Prepare a simple data source.
        var data = new ReportData { Value = 12345 };

        // 3. Populate the template using the ReportingEngine.
        ReportingEngine engine = new ReportingEngine();
        // The data source is referenced in the template as "ds".
        engine.BuildReport(doc, data, "ds");

        // 4. Save the populated document as SVG.
        // Use placed glyphs so the text is rendered as curves (no selectable text).
        SvgSaveOptions svgOptions = new SvgSaveOptions
        {
            SaveFormat = SaveFormat.Svg,
            TextOutputMode = SvgTextOutputMode.UsePlacedGlyphs,
            ShowPageBorder = false   // Optional: omit the page border in the SVG.
        };

        doc.Save("Result.svg", svgOptions);
    }

    // Simple POCO used as the data source for the report.
    public class ReportData
    {
        public int Value { get; set; }
    }
}
