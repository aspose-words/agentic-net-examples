using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Markup;

public class Program
{
    public static void Main()
    {
        // Create a simple PNG image (1x1 pixel) and save it to a local file.
        const string imageFileName = "sample.png";
        byte[] pngData = Convert.FromBase64String(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+X9eUAAAAASUVORK5CYII=");
        File.WriteAllBytes(imageFileName, pngData);

        // -----------------------------------------------------------------
        // Step 1: Build a sample document that contains picture content controls.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert three picture content controls, each with the same image.
        for (int i = 0; i < 3; i++)
        {
            // Insert a picture content control at the current cursor position.
            StructuredDocumentTag pictureSdt = builder.InsertStructuredDocumentTag(SdtType.Picture);

            // Move the cursor inside the newly created content control.
            builder.MoveTo(pictureSdt);

            // Insert the image (inline) inside the content control.
            builder.InsertImage(imageFileName);

            // Add a paragraph break after each control for readability.
            builder.Writeln();
        }

        // Save the source document.
        const string inputPath = "input.docx";
        doc.Save(inputPath);

        // -----------------------------------------------------------------
        // Step 2: Load the document and replace picture content controls with inline images.
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(inputPath);

        // Find all picture content controls in the document.
        var pictureControls = loadedDoc
            .GetChildNodes(NodeType.StructuredDocumentTag, true)
            .OfType<StructuredDocumentTag>()
            .Where(sdt => sdt.SdtType == SdtType.Picture)
            .ToList();

        // For each picture content control, remove the control but keep its inner image.
        foreach (var sdt in pictureControls)
        {
            // RemoveSelfOnly keeps the child nodes (the image shape) in the document.
            sdt.RemoveSelfOnly();
        }

        // Save the modified document.
        const string outputPath = "output.docx";
        loadedDoc.Save(outputPath);
    }
}
