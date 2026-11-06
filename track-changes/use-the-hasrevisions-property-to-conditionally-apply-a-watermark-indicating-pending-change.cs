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

        // Start tracking revisions.
        doc.StartTrackRevisions("DemoAuthor", DateTime.Now);

        // Make some changes while tracking is enabled.
        builder.Writeln("This paragraph was added while tracking changes.");
        builder.MoveToDocumentStart();
        builder.Write("Inserted at the beginning. ");

        // Stop tracking revisions.
        doc.StopTrackRevisions();

        // If the document has any revisions, add a watermark indicating pending changes.
        if (doc.HasRevisions)
        {
            // Create a watermark shape.
            Shape watermark = new Shape(doc, ShapeType.TextPlainText);
            watermark.TextPath.Text = "PENDING CHANGES";
            watermark.TextPath.FontFamily = "Arial";
            watermark.Width = 500;
            watermark.Height = 100;
            watermark.Rotation = -40;
            watermark.FillColor = Color.LightGray;
            watermark.StrokeColor = Color.LightGray;
            watermark.RelativeHorizontalPosition = RelativeHorizontalPosition.Page;
            watermark.RelativeVerticalPosition = RelativeVerticalPosition.Page;
            watermark.WrapType = WrapType.None;
            watermark.VerticalAlignment = VerticalAlignment.Center;
            watermark.HorizontalAlignment = HorizontalAlignment.Center;
            watermark.BehindText = true;

            // Insert the watermark into each header of each section.
            foreach (Section section in doc.Sections)
            {
                HeaderFooter header = section.HeadersFooters[HeaderFooterType.HeaderPrimary];
                if (header == null)
                {
                    header = new HeaderFooter(doc, HeaderFooterType.HeaderPrimary);
                    section.HeadersFooters.Add(header);
                }

                // The header must contain a paragraph; add the shape to a new paragraph.
                Paragraph para = new Paragraph(doc);
                para.AppendChild(watermark.Clone(true));
                header.AppendChild(para);
            }
        }

        // Save the document.
        string outputPath = "Output.docx";
        doc.Save(outputPath);

        // Indicate completion.
        Console.WriteLine($"Document saved to '{outputPath}'. Watermark applied: {doc.HasRevisions}");
    }
}
