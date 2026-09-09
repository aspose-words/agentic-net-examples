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
        // -----------------------------------------------------------------
        // 1. Create a tiny PNG image (1x1 pixel) to be used in the sample document.
        // -----------------------------------------------------------------
        const string base64Png =
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XK8cAAAAASUVORK5CYII=";
        byte[] pngBytes = Convert.FromBase64String(base64Png);
        const string imageFileName = "sample.png";
        File.WriteAllBytes(imageFileName, pngBytes);

        // -----------------------------------------------------------------
        // 2. Create a sample document that contains a picture content control.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a picture content control at the current cursor position.
        StructuredDocumentTag pictureSdt = builder.InsertStructuredDocumentTag(SdtType.Picture);

        // Insert the image inside the picture content control.
        builder.InsertImage(imageFileName);

        // Save the source document.
        const string inputPath = "input.docx";
        doc.Save(inputPath);

        // -----------------------------------------------------------------
        // 3. Load the document and replace each picture content control with an inline image.
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(inputPath);

        // Find all picture content controls in the document.
        var pictureSdts = loadedDoc
            .GetChildNodes(NodeType.StructuredDocumentTag, true)
            .OfType<StructuredDocumentTag>()
            .Where(sdt => sdt.SdtType == SdtType.Picture)
            .ToList();

        foreach (StructuredDocumentTag sdt in pictureSdts)
        {
            // Remove the content control but keep its children (the image) in the document.
            // This method preserves the original inline shape without needing to clone or re‑insert it.
            sdt.RemoveSelfOnly();
        }

        // Save the resulting document.
        const string outputPath = "output.docx";
        loadedDoc.Save(outputPath);
    }
}
