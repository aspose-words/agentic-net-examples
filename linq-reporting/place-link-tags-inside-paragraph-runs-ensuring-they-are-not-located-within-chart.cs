using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts; // Added namespace for ChartType
using Aspose.Words.Tables;

class HyperlinkInParagraphRuns
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // -------------------------------------------------
        // First paragraph – insert a hyperlink in a run.
        // -------------------------------------------------
        builder.Write("Visit the ");
        // Set hyperlink formatting (blue and underlined) – required for visual appearance.
        builder.Font.Color = Color.Blue;
        builder.Font.Underline = Underline.Single;
        // InsertHyperlink adds a HYPERLINK field inside the current run.
        builder.InsertHyperlink("Aspose website", "https://www.aspose.com", false);
        // Reset formatting after the hyperlink.
        builder.Font.ClearFormatting();
        builder.Writeln(" for more information.");

        // -------------------------------------------------
        // Insert a chart – hyperlinks must NOT be placed inside chart elements.
        // -------------------------------------------------
        // The chart is a free‑floating shape; we insert it after the first paragraph.
        builder.InsertChart(ChartType.Column, 400, 300); // ChartType now resolved from Aspose.Words.Drawing.Charts
        // Add a paragraph break after the chart to separate it from following content.
        builder.Writeln();

        // -------------------------------------------------
        // Second paragraph – insert another hyperlink, ensuring it is outside the chart.
        // -------------------------------------------------
        builder.Write("Check out the ");
        builder.Font.Color = Color.Blue;
        builder.Font.Underline = Underline.Single;
        builder.InsertHyperlink("documentation", "https://docs.aspose.com/words/net/", false);
        builder.Font.ClearFormatting();
        builder.Writeln(" for API details.");

        // Save the document to disk.
        doc.Save("HyperlinksOutsideChart.docx");
    }
}
