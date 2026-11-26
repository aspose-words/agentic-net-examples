using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Notes;
using Aspose.Drawing;

namespace FootnoteImageExtraction
{
    public class Program
    {
        public static void Main()
        {
            // Create a deterministic sample image file.
            const string sampleImagePath = "sample.png";
            CreateSampleImage(sampleImagePath);

            // Create a Word document with footnotes that contain images.
            const string docPath = "DocumentWithFootnoteImages.docx";
            CreateDocumentWithFootnoteImages(docPath, sampleImagePath);

            // Extract images from footnotes and save them using footnote numbers.
            ExtractFootnoteImages(docPath);
        }

        private static void CreateSampleImage(string path)
        {
            // 100x100 white bitmap with a red rectangle.
            Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(100, 100);
            try
            {
                Aspose.Drawing.Graphics graphics = Aspose.Drawing.Graphics.FromImage(bitmap);
                try
                {
                    graphics.Clear(Aspose.Drawing.Color.White);
                    using (Aspose.Drawing.Pen pen = new Aspose.Drawing.Pen(Aspose.Drawing.Color.Red, 5))
                    {
                        graphics.DrawRectangle(pen, 10, 10, 80, 80);
                    }
                }
                finally
                {
                    graphics.Dispose();
                }

                bitmap.Save(path);
            }
            finally
            {
                bitmap.Dispose();
            }

            if (!File.Exists(path))
                throw new InvalidOperationException($"Failed to create sample image at '{path}'.");
        }

        private static void CreateDocumentWithFootnoteImages(string docPath, string imagePath)
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // First paragraph with a footnote containing an image.
            builder.Writeln("This is a paragraph with a footnote reference.");
            Footnote footnote1 = builder.InsertFootnote(FootnoteType.Footnote, "1");
            // Move the builder into the footnote paragraph and insert the image.
            builder.MoveTo(footnote1.FirstParagraph);
            builder.InsertImage(imagePath);
            builder.Writeln(); // End of footnote paragraph.

            // Second paragraph with another footnote containing the same image.
            builder.Writeln("Another paragraph with a second footnote.");
            Footnote footnote2 = builder.InsertFootnote(FootnoteType.Footnote, "2");
            builder.MoveTo(footnote2.FirstParagraph);
            builder.InsertImage(imagePath);
            builder.Writeln();

            // Save the document.
            doc.Save(docPath);

            if (!File.Exists(docPath))
                throw new InvalidOperationException($"Failed to save document at '{docPath}'.");
        }

        private static void ExtractFootnoteImages(string docPath)
        {
            Document doc = new Document(docPath);
            NodeCollection footnotes = doc.GetChildNodes(NodeType.Footnote, true);

            int extractedCount = 0;

            foreach (Footnote footnote in footnotes)
            {
                // Enumerate all Shape nodes inside the footnote.
                NodeCollection shapes = footnote.GetChildNodes(NodeType.Shape, true);
                int imageIndex = 0;

                foreach (Shape shape in shapes)
                {
                    if (shape.HasImage)
                    {
                        // Use the footnote's reference mark (the number shown in the document) for naming.
                        string footnoteNumber = footnote.ReferenceMark ?? "unknown";
                        string outputFileName = $"footnote-{footnoteNumber}-{imageIndex}.png";

                        shape.ImageData.Save(outputFileName);
                        if (!File.Exists(outputFileName))
                            throw new InvalidOperationException($"Image was not saved to '{outputFileName}'.");

                        extractedCount++;
                        imageIndex++;
                    }
                }
            }

            if (extractedCount == 0)
                throw new InvalidOperationException("No images were extracted from footnotes.");

            Console.WriteLine($"Extracted {extractedCount} image(s) from footnotes.");
        }
    }
}
