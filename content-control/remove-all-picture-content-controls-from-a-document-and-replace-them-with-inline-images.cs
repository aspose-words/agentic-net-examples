using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Markup;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a sample document that contains picture content controls.
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);

        // First picture content control.
        StructuredDocumentTag pictureSdt1 = builder.InsertStructuredDocumentTag(SdtType.Picture);
        // Insert an image into the picture content control.
        builder.InsertImage(GetSampleImageBytes());

        // Add some text between controls.
        builder.Writeln();
        builder.Write("Some text between picture controls.");
        builder.Writeln();

        // Second picture content control.
        StructuredDocumentTag pictureSdt2 = builder.InsertStructuredDocumentTag(SdtType.Picture);
        builder.InsertImage(GetSampleImageBytes());

        // Save the source document (optional, just for demonstration).
        sourceDoc.Save("Input.docx");

        // Process the document: replace each picture content control with an inline image.
        NodeCollection sdtNodes = sourceDoc.GetChildNodes(NodeType.StructuredDocumentTag, true);
        DocumentBuilder editBuilder = new DocumentBuilder(sourceDoc);

        // Iterate over a copy because we will modify the document.
        foreach (StructuredDocumentTag sdt in sdtNodes.Cast<StructuredDocumentTag>().ToList())
        {
            if (sdt.SdtType == SdtType.Picture)
            {
                // Find the Shape (image) inside the picture content control.
                Shape shape = sdt.GetChildNodes(NodeType.Shape, true)
                                 .Cast<Shape>()
                                 .FirstOrDefault();

                if (shape != null && shape.HasImage)
                {
                    // Get the image bytes from the shape.
                    byte[] imageBytes = shape.ImageData.ImageBytes;

                    // Insert the image just before the content control.
                    editBuilder.MoveTo(sdt);
                    editBuilder.InsertImage(imageBytes);
                }

                // Remove the original picture content control (its inner content is now empty).
                sdt.Remove();
            }
        }

        // Save the modified document.
        sourceDoc.Save("Output.docx");
    }

    // Returns a simple 1x1 red PNG image as a byte array.
    private static byte[] GetSampleImageBytes()
    {
        // This is a base64‑encoded PNG of a 1×1 red pixel.
        const string base64Png =
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAIAAACQd1PeAAAADUlEQVR42mP8z/C/HwAFgwJ/lKXK5wAAAABJRU5ErkJggg==";

        return Convert.FromBase64String(base64Png);
    }
}
