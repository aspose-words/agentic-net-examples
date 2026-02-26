using System;
using System.Drawing;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Output file name.
        const string outputPath = "Result.docx";

        // -------------------------------------------------
        // 1. Create a source document that will be inserted.
        // -------------------------------------------------
        Document srcDoc = new Document();
        DocumentBuilder srcBuilder = new DocumentBuilder(srcDoc);

        // Define a custom style named "MyStyle".
        Style myStyle = srcDoc.Styles.Add(StyleType.Paragraph, "MyStyle");
        myStyle.Font.Name = "Courier New";
        myStyle.Font.Size = 14;
        myStyle.Font.Color = Color.Blue;

        // Apply the custom style to a paragraph.
        srcBuilder.ParagraphFormat.StyleName = "MyStyle";
        srcBuilder.Writeln("This paragraph uses MyStyle.");

        // -------------------------------------------------
        // 2. Create a template document that contains the
        //    ReportingEngine tag for dynamic insertion.
        // -------------------------------------------------
        Document template = new Document();
        DocumentBuilder tmplBuilder = new DocumentBuilder(template);

        tmplBuilder.Writeln("Report start");
        // The "-sourceStyles" switch tells the engine to keep the source
        // document's styles when inserting it.
        tmplBuilder.Writeln("<<doc [src.Document] -sourceStyles>>");
        tmplBuilder.Writeln("Report end");

        // -------------------------------------------------
        // 3. Prepare the data source that holds the document to
        //    be inserted. The name "src" must match the tag.
        // -------------------------------------------------
        var data = new DocumentHolder { Document = srcDoc };

        // -------------------------------------------------
        // 4. Build the report using ReportingEngine.
        // -------------------------------------------------
        ReportingEngine engine = new ReportingEngine
        {
            // Remove any empty paragraphs that may appear after insertion.
            Options = ReportBuildOptions.RemoveEmptyParagraphs
        };

        // BuildReport overload that accepts multiple data sources.
        engine.BuildReport(template, new object[] { data }, new[] { "src" });

        // -------------------------------------------------
        // 5. Save the final document.
        // -------------------------------------------------
        template.Save(outputPath);
    }

    // Simple holder class referenced from the template.
    public class DocumentHolder
    {
        public Document Document { get; set; }
    }
}
