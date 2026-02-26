using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsReportingDemo
{
    // Simple wrapper class that will be used as the data source for the template.
    // The template will reference the property "Document" via the tag <<doc [src.Document] -sourceStyles>>.
    public class DocumentContainer
    {
        public Document Document { get; set; }
    }

    class Program
    {
        static void Main()
        {
            // -----------------------------------------------------------------
            // 1. Create the DOT (template) document.
            // -----------------------------------------------------------------
            Document template = new Document();                     // create a blank document
            DocumentBuilder tmplBuilder = new DocumentBuilder(template);
            // The tag tells the ReportingEngine to insert the source document and to copy its styles.
            tmplBuilder.Writeln("<<doc [src.Document] -sourceStyles>>");

            // -----------------------------------------------------------------
            // 2. Create the source document that will be inserted into the template.
            // -----------------------------------------------------------------
            Document source = new Document();                       // create a blank document
            DocumentBuilder srcBuilder = new DocumentBuilder(source);

            // Define a custom style in the source document.
            Style customStyle = source.Styles.Add(StyleType.Paragraph, "MyCustomStyle");
            customStyle.Font.Name = "Courier New";
            customStyle.Font.Size = 14;
            customStyle.Font.Color = Color.Blue;

            // Apply the custom style to a paragraph.
            srcBuilder.ParagraphFormat.StyleName = "MyCustomStyle";
            srcBuilder.Writeln("This paragraph uses MyCustomStyle defined in the source document.");

            // -----------------------------------------------------------------
            // 3. Prepare the data source for the ReportingEngine.
            // -----------------------------------------------------------------
            DocumentContainer srcContainer = new DocumentContainer { Document = source };

            // -----------------------------------------------------------------
            // 4. Build the report – the source document will be inserted and its
            //    styles will be merged into the resulting document because of the
            //    "-sourceStyles" switch in the template tag.
            // -----------------------------------------------------------------
            ReportingEngine engine = new ReportingEngine();
            // The data source name "src" must match the name used in the template tag.
            engine.BuildReport(template, srcContainer, "src");

            // -----------------------------------------------------------------
            // 5. Save the final document.
            // -----------------------------------------------------------------
            template.Save("Result.docx");
        }
    }
}
