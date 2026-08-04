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

        // Write some initial content that is not a revision.
        builder.Writeln("Original content. ");

        // Start tracking revisions.
        doc.StartTrackRevisions("Author", DateTime.Now);

        // Make a change that will be recorded as a revision.
        builder.Writeln("This line is added while tracking changes.");

        // Stop tracking revisions.
        doc.StopTrackRevisions();

        // If the document contains any revisions, add a watermark.
        if (doc.HasRevisions)
        {
            // Configure watermark options.
            TextWatermarkOptions options = new TextWatermarkOptions
            {
                FontFamily = "Arial",
                FontSize = 36,
                Color = Color.LightGray,
                Layout = WatermarkLayout.Diagonal,
                IsSemitrasparent = false
            };

            // Apply the text watermark to the document.
            doc.Watermark.SetText("PENDING CHANGES", options);
        }

        // Save the document to a file.
        doc.Save("Output.docx");
    }
}
