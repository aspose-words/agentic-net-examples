using System;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add initial content that is not a revision.
        builder.Writeln("Original content. ");

        // Start tracking revisions.
        doc.StartTrackRevisions("Reviewer", DateTime.Now);

        // Make some changes that will be recorded as revisions.
        builder.Writeln("This line is added while tracking changes. ");

        // Stop tracking revisions.
        doc.StopTrackRevisions();

        // If the document contains any revisions, add a watermark indicating pending changes.
        if (doc.HasRevisions)
        {
            // Configure a semi‑transparent diagonal text watermark.
            TextWatermarkOptions options = new TextWatermarkOptions
            {
                IsSemitrasparent = true,
                Layout = WatermarkLayout.Diagonal
            };

            // Apply the watermark to the document.
            doc.Watermark.SetText("Pending Changes", options);
        }

        // Save the resulting document.
        doc.Save("Output.docx");
    }
}
