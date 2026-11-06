using System;
using System.IO;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a folder for sample documents.
        string folderPath = Path.Combine(Directory.GetCurrentDirectory(), "Docs");
        Directory.CreateDirectory(folderPath);

        // Generate sample DOCX files.
        for (int i = 1; i <= 2; i++)
        {
            string filePath = Path.Combine(folderPath, $"Sample{i}.docx");
            Document sampleDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(sampleDoc);
            builder.Writeln($"This is sample document {i}.");
            sampleDoc.Save(filePath);
        }

        // Process each DOCX file in the folder, inserting a semi‑transparent watermark.
        foreach (string file in Directory.GetFiles(folderPath, "*.docx"))
        {
            Document doc = new Document(file);
            InsertWatermark(doc, "CONFIDENTIAL");
            doc.Save(file);
        }

        Console.WriteLine("Watermark insertion completed.");
    }

    private static void InsertWatermark(Document doc, string watermarkText)
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

            // Ensure the header contains at least one paragraph.
            Paragraph para;
            if (header.Paragraphs.Count == 0)
            {
                para = new Paragraph(doc);
                header.AppendChild(para);
            }
            else
            {
                para = header.Paragraphs[0];
            }

            // Build the watermark shape.
            Shape watermark = new Shape(doc, ShapeType.TextPlainText);
            watermark.TextPath.Text = watermarkText;
            watermark.TextPath.FontFamily = "Arial";
            watermark.Width = 500;
            watermark.Height = 100;
            watermark.Rotation = -40;
            watermark.Fill.Color = Color.Gray;
            watermark.Fill.Transparency = 0.5f; // Semi‑transparent.
            watermark.Stroke.Color = Color.Gray;
            watermark.RelativeHorizontalPosition = RelativeHorizontalPosition.Page;
            watermark.RelativeVerticalPosition = RelativeVerticalPosition.Page;
            watermark.WrapType = WrapType.None;
            watermark.HorizontalAlignment = HorizontalAlignment.Center;
            watermark.VerticalAlignment = VerticalAlignment.Center;

            // Insert the shape into the header paragraph.
            para.AppendChild(watermark);
        }
    }
}
