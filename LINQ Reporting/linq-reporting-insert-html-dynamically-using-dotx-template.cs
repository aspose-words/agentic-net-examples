using System;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Fields;

class ReportData
{
    // Simple data class that holds a title and an HTML fragment.
    public string Title { get; set; }
    public string HtmlContent { get; set; }
}

class Program
{
    static void Main()
    {
        // Load the DOTX template that contains a placeholder for the title
        // (e.g. <<[data.Title]>>) and a MERGEFIELD named HtmlContent.
        Document template = new Document("Template.dotx");

        // Prepare the data source.
        var data = new ReportData
        {
            Title = "Dynamic HTML Report",
            HtmlContent = "<h2>Report Section</h2><p>This is <b>bold</b> and <i>italic</i> text.</p>"
        };

        // Populate the non‑HTML fields using the ReportingEngine.
        // The template can reference the data source as 'data'.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(template, data, "data");

        // Insert the HTML fragment at the location of the MERGEFIELD named 'HtmlContent'.
        DocumentBuilder builder = new DocumentBuilder(template);
        builder.MoveToMergeField("HtmlContent");
        // The second argument tells InsertHtml to apply the builder's formatting as a base.
        builder.InsertHtml(data.HtmlContent, true);

        // Save the resulting document.
        template.Save("Result.docx");
    }
}
