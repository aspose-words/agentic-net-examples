using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new document and add initial content.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Original paragraph.");

        // Enable track changes.
        doc.StartTrackRevisions("John Doe", DateTime.Now);

        // Perform modifications while tracking is enabled.
        builder.Writeln("Added paragraph while tracking.");
        builder.MoveToDocumentStart();
        builder.Write("Inserted at start. ");

        // Stop tracking changes.
        doc.StopTrackRevisions();

        // If the document contains any revisions, add a watermark.
        if (doc.HasRevisions)
        {
            // Add watermark to each section's primary header.
            foreach (Section section in doc.Sections)
            {
                // Ensure the primary header exists.
                HeaderFooter header = section.HeadersFooters[HeaderFooterType.HeaderPrimary];
                if (header == null)
                {
                    header = new HeaderFooter(doc, HeaderFooterType.HeaderPrimary);
                    section.HeadersFooters.Add(header);
                }

                // Move the builder to the header.
                DocumentBuilder watermarkBuilder = new DocumentBuilder(doc);
                watermarkBuilder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);

                // Create a text shape that will act as the watermark.
                Shape watermark = new Shape(doc, ShapeType.TextPlainText);
                watermark.TextPath.Text = "Pending Changes";
                watermark.TextPath.FontFamily = "Arial";
                watermark.Width = 500;
                watermark.Height = 100;
                watermark.Rotation = -40;
                watermark.Fill.Color = Color.LightGray;
                watermark.StrokeColor = Color.LightGray;
                watermark.RelativeHorizontalPosition = RelativeHorizontalPosition.Page;
                watermark.RelativeVerticalPosition = RelativeVerticalPosition.Page;
                watermark.WrapType = WrapType.None;
                watermark.VerticalAlignment = VerticalAlignment.Center;
                watermark.HorizontalAlignment = HorizontalAlignment.Center;

                // Insert the watermark shape into the header.
                watermarkBuilder.InsertNode(watermark);
            }
        }

        // Save the resulting document.
        doc.Save("Result.docx");
    }
}
