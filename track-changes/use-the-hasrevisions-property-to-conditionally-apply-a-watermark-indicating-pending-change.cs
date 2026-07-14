using System;
using System.Drawing;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write some initial content that is not a revision.
        builder.Writeln("This is the original content.");

        // Start tracking revisions and make a change.
        doc.StartTrackRevisions("Reviewer", DateTime.Now);
        builder.Writeln("This line was added while tracking changes.");
        // Stop tracking so further edits are not recorded as revisions.
        doc.StopTrackRevisions();

        // If the document has any revisions, add a watermark indicating pending changes.
        if (doc.HasRevisions)
        {
            // Configure watermark options.
            TextWatermarkOptions options = new TextWatermarkOptions
            {
                FontFamily = "Arial",
                FontSize = 36,
                Color = Color.LightGray,
                Layout = WatermarkLayout.Diagonal,
                IsSemitrasparent = true // Semi‑transparent watermark.
            };

            // Apply the text watermark to the document.
            doc.Watermark.SetText("PENDING CHANGES", options);
        }

        // Save the document.
        doc.Save("Output.docx");
    }
}
