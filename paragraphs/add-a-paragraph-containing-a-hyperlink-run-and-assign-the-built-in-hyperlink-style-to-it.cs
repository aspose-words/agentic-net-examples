using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a new empty paragraph and move the cursor into it.
        builder.InsertParagraph();

        // Insert a hyperlink run inside the current paragraph.
        builder.InsertHyperlink("Visit Aspose", "https://www.aspose.com", false);

        // Retrieve the paragraph that now contains the hyperlink.
        Paragraph paragraph = (Paragraph)builder.CurrentParagraph;

        // Apply the built‑in Hyperlink character style to the hyperlink run.
        if (paragraph.Runs.Count > 0)
        {
            Run hyperlinkRun = (Run)paragraph.Runs[0];
            hyperlinkRun.Font.StyleIdentifier = StyleIdentifier.Hyperlink;
        }

        // Save the document to disk.
        doc.Save("Output.docx");
    }
}
