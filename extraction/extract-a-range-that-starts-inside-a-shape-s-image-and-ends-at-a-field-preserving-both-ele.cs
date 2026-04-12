using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Fields;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // Create a new blank document and a builder.
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);

        // ---------- Create a sample image ----------
        // Generate a simple red square bitmap using Aspose.Drawing.
        using (Bitmap bitmap = new Bitmap(100, 100))
        {
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                graphics.Clear(Color.Red);
            }

            // Save the bitmap to a memory stream.
            using (MemoryStream imageStream = new MemoryStream())
            {
                bitmap.Save(imageStream, ImageFormat.Png);
                imageStream.Position = 0;

                // Insert the image inline into the document and keep a reference to the created shape.
                Shape imageShape = builder.InsertImage(imageStream);

                // Insert a DATE field right after the image, in the same paragraph.
                builder.InsertField(FieldType.FieldDate, true);

                // Save the source document (optional, for inspection).
                sourceDoc.Save("SourceDocument.docx");

                // ---------- Validate that the shape contains an image ----------
                if (!imageShape.HasImage)
                    throw new InvalidOperationException("The inserted shape does not contain an image.");

                // ---------- Locate the field start ----------
                // The field is placed inside the first paragraph.
                Paragraph paragraph = sourceDoc.FirstSection.Body.FirstParagraph;
                FieldStart fieldStart = paragraph.GetChildNodes(NodeType.FieldStart, true)
                                                 .Cast<FieldStart>()
                                                 .FirstOrDefault();

                if (fieldStart == null)
                    throw new InvalidOperationException("The document does not contain a field.");

                // ---------- Extract the range ----------
                // The range we need is the whole paragraph that contains both the image shape and the field.
                // Use NodeImporter to import the paragraph into a new document.
                Document extractedDoc = new Document();
                extractedDoc.RemoveAllChildren();

                // Build minimal document structure: Section -> Body.
                Section section = new Section(extractedDoc);
                extractedDoc.AppendChild(section);
                Body body = new Body(extractedDoc);
                section.AppendChild(body);

                // Import the paragraph from the source document.
                NodeImporter importer = new NodeImporter(sourceDoc, extractedDoc, ImportFormatMode.KeepSourceFormatting);
                Node importedParagraph = importer.ImportNode(paragraph, true);
                body.AppendChild(importedParagraph);

                // Save the extracted document.
                extractedDoc.Save("ExtractedRange.docx");

                // ---------- Validation ----------
                Shape extractedShape = extractedDoc.GetChildNodes(NodeType.Shape, true)
                                                  .Cast<Shape>()
                                                  .FirstOrDefault();

                FieldStart extractedFieldStart = extractedDoc.GetChildNodes(NodeType.FieldStart, true)
                                                            .Cast<FieldStart>()
                                                            .FirstOrDefault();

                if (extractedShape == null || !extractedShape.HasImage)
                    throw new InvalidOperationException("Extracted document does not contain the image shape.");

                if (extractedFieldStart == null)
                    throw new InvalidOperationException("Extracted document does not contain the field.");

                Console.WriteLine("Extraction completed successfully.");
            }
        }
    }
}
