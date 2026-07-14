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
        // Create a sample document that contains a picture content control.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a picture content control at the current cursor position.
        StructuredDocumentTag pictureSdt = builder.InsertStructuredDocumentTag(SdtType.Picture);

        // Insert a tiny PNG image (1x1 pixel) into the content control.
        // The image is provided as a Base64‑encoded byte array to avoid external files.
        byte[] pngBytes = Convert.FromBase64String(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XK6cAAAAASUVORK5CYII=");
        builder.InsertImage(pngBytes);

        // Save the source document.
        const string inputPath = "input.docx";
        doc.Save(inputPath);

        // Load the document that contains picture content controls.
        Document loadedDoc = new Document(inputPath);

        // Find all picture content controls in the document.
        var pictureControls = loadedDoc
            .GetChildNodes(NodeType.StructuredDocumentTag, true)
            .OfType<StructuredDocumentTag>()
            .Where(sdt => sdt.SdtType == SdtType.Picture)
            .ToList();

        // Replace each picture content control with its inner image.
        foreach (var sdt in pictureControls)
        {
            // Remove the SDT node but keep its child nodes (the image shape) in the document.
            sdt.RemoveSelfOnly();
        }

        // Save the resulting document where picture content controls have been removed.
        const string outputPath = "output.docx";
        loadedDoc.Save(outputPath);
    }
}
