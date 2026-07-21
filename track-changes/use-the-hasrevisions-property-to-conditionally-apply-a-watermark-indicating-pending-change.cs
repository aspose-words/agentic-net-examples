using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add some initial content that is not a revision.
        builder.Writeln("This is the original content.");

        // Start tracking revisions.
        doc.StartTrackRevisions("Reviewer", DateTime.Now);

        // Make changes that will be recorded as revisions.
        builder.Writeln("This line was added while tracking changes.");
        builder.Writeln("Another revision line.");

        // Stop tracking revisions.
        doc.StopTrackRevisions();

        // If the document contains any revisions, add a watermark indicating pending changes.
        if (doc.HasRevisions)
        {
            // Create watermark options for a semi‑transparent diagonal text.
            TextWatermarkOptions options = new TextWatermarkOptions
            {
                FontFamily = "Arial",
                FontSize = 36,
                Color = Color.Gray,
                Layout = WatermarkLayout.Diagonal,
                IsSemitrasparent = true
            };

            // Apply the watermark to the document.
            doc.Watermark.SetText("PENDING CHANGES", options);
        }

        // Save the resulting document.
        doc.Save("Output.docx");
    }
}
