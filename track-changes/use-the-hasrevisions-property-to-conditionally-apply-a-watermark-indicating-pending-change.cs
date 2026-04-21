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

        // Write some initial text that is not a revision.
        builder.Writeln("This text is not tracked.");

        // Start tracking revisions.
        doc.StartTrackRevisions("Reviewer", DateTime.Now);

        // Add text that will be recorded as a revision.
        builder.Writeln("This text is a pending change.");

        // Stop tracking revisions.
        doc.StopTrackRevisions();

        // If the document contains any revisions, add a watermark.
        if (doc.HasRevisions)
        {
            // Configure watermark options.
            TextWatermarkOptions options = new TextWatermarkOptions
            {
                FontFamily = "Arial",
                FontSize = 40,
                Color = Color.LightGray,
                Layout = WatermarkLayout.Diagonal,
                IsSemitrasparent = true // Makes the watermark appear behind the text.
            };

            // Apply the watermark to the whole document.
            doc.Watermark.SetText("PENDING CHANGES", options);
        }

        // Save the resulting document.
        doc.Save("WatermarkedWithRevisions.docx");
    }
}
