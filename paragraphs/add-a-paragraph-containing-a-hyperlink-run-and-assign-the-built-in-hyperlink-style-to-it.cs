using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a new paragraph.
        builder.Writeln();

        // Insert a hyperlink with display text.
        string displayText = "Visit Aspose";
        string url = "https://www.aspose.com";
        builder.InsertHyperlink(displayText, url, false);

        // Get the paragraph that now contains the hyperlink.
        Paragraph paragraph = builder.CurrentParagraph;

        // Apply the built‑in Hyperlink character style to the run that holds the display text.
        foreach (Run run in paragraph.Runs)
        {
            if (run.Text == displayText)
            {
                run.Font.StyleIdentifier = StyleIdentifier.Hyperlink;
                break;
            }
        }

        // Save the document to a file.
        doc.Save("HyperlinkParagraph.docx");
    }
}
