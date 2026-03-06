using System;
using System.Data;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // HTML fragment to be inserted.
        string html = "<p align='right'>Right aligned paragraph</p>" +
                      "<h2 align='center'>Center heading</h2>" +
                      "<div><b>Bold text</b> and <i>italic</i></div>";

        // Insert the HTML using builder formatting and remove the trailing empty paragraph.
        builder.InsertHtml(html, HtmlInsertOptions.UseBuilderFormatting | HtmlInsertOptions.RemoveLastEmptyParagraph);

        // Configure ReportingEngine to allow missing members and provide a custom message.
        ReportingEngine engine = new ReportingEngine
        {
            Options = ReportBuildOptions.AllowMissingMembers,
            MissingMemberMessage = "Missing"
        };

        // Build a report with an empty data source to demonstrate the MissingMemberMessage handling.
        engine.BuildReport(doc, new DataSet(), string.Empty);

        // Save the document as HTML, exporting images as Base64 and stripping JavaScript from links.
        HtmlSaveOptions saveOptions = new HtmlSaveOptions
        {
            ExportImagesAsBase64 = true,
            RemoveJavaScriptFromLinks = true
        };
        doc.Save("Output.html", saveOptions);
    }
}
