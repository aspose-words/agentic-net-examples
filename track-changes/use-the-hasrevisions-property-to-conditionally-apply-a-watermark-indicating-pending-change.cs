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

        // Add initial content.
        builder.Writeln("This is the original paragraph.");
        builder.Writeln("Another line of text.");

        // Start tracking revisions.
        doc.StartTrackRevisions("John Doe", DateTime.Now);

        // Make some changes while tracking is enabled.
        builder.Writeln("This line was added while tracking changes.");
        builder.MoveToDocumentStart();
        builder.Write("Inserted at the beginning. ");

        // Simulate a deletion revision by striking through text.
        builder.MoveTo(doc.FirstSection.Body.FirstParagraph);
        builder.Font.StrikeThrough = true;
        builder.Write("original ");
        builder.Font.StrikeThrough = false;

        // Change formatting of a paragraph.
        builder.MoveTo(doc.FirstSection.Body.Paragraphs[1]);
        builder.Font.Bold = true;
        builder.Write("Bolded text.");
        builder.Font.Bold = false;

        // Stop tracking revisions.
        doc.StopTrackRevisions();

        // If the document contains any revisions, add a watermark.
        if (doc.HasRevisions)
        {
            foreach (Section section in doc.Sections)
            {
                // Ensure the primary header exists.
                HeaderFooter header = section.HeadersFooters[HeaderFooterType.HeaderPrimary];
                if (header == null)
                {
                    header = new HeaderFooter(doc, HeaderFooterType.HeaderPrimary);
                    section.HeadersFooters.Add(header);
                }

                // Create a paragraph to hold the watermark shape.
                Paragraph para = new Paragraph(doc);
                header.AppendChild(para);

                // Create the watermark shape.
                Shape watermark = new Shape(doc, ShapeType.TextPlainText);
                watermark.TextPath.Text = "PENDING CHANGES";
                watermark.TextPath.FontFamily = "Arial";
                watermark.TextPath.Bold = true;
                watermark.Fill.Color = Color.LightGray;
                watermark.Stroke.Color = Color.LightGray;
                watermark.Rotation = -40;
                watermark.WrapType = WrapType.None;
                watermark.RelativeHorizontalPosition = RelativeHorizontalPosition.Page;
                watermark.RelativeVerticalPosition = RelativeVerticalPosition.Page;
                watermark.HorizontalAlignment = HorizontalAlignment.Center;
                watermark.VerticalAlignment = VerticalAlignment.Center;
                watermark.Width = 500;
                watermark.Height = 100;

                // Insert the watermark into the paragraph.
                para.AppendChild(watermark);
            }
        }

        // Save the document.
        doc.Save("Output.docx");
    }
}
