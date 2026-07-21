using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // -----------------------------------------------------------------
        // Setup working directories
        // -----------------------------------------------------------------
        const string workDir = "Work";
        Directory.CreateDirectory(workDir);
        string docPath = Path.Combine(workDir, "sample.doc");
        string outputDir = Path.Combine(workDir, "Extracted");
        Directory.CreateDirectory(outputDir);

        // -----------------------------------------------------------------
        // 1. Create a temporary document and render its first page to EMF.
        // -----------------------------------------------------------------
        Document tempDoc = new Document();
        DocumentBuilder tempBuilder = new DocumentBuilder(tempDoc);
        tempBuilder.Writeln("Temporary EMF content");

        // Save the first page of the temporary document as an EMF image into a memory stream.
        using (MemoryStream emfStream = new MemoryStream())
        {
            ImageSaveOptions emfOptions = new ImageSaveOptions(SaveFormat.Emf);
            // Render only the first page.
            emfOptions.PageSet = new PageSet(0);
            tempDoc.Save(emfStream, emfOptions);
            emfStream.Position = 0; // Reset for reading.

            // -----------------------------------------------------------------
            // 2. Create the main document and embed the EMF image.
            // -----------------------------------------------------------------
            Document mainDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(mainDoc);
            // Insert the EMF image from the stream.
            builder.InsertImage(emfStream);
            mainDoc.Save(docPath);
        }

        // -----------------------------------------------------------------
        // 3. Load the document and extract all embedded EMF vector images.
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(docPath);

        var emfShapes = loadedDoc.GetChildNodes(NodeType.Shape, true)
                                 .Cast<Shape>()
                                 .Where(s => s.HasImage && s.ImageData.ImageType == ImageType.Emf)
                                 .ToList();

        if (!emfShapes.Any())
            throw new InvalidOperationException("No EMF vector images were found in the document.");

        int index = 0;
        foreach (var shape in emfShapes)
        {
            string extractedPath = Path.Combine(outputDir,
                $"extracted_{index}{FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType)}");
            shape.ImageData.Save(extractedPath);
            Console.WriteLine($"Extracted EMF image saved to: {extractedPath}");
            index++;
        }

        // -----------------------------------------------------------------
        // Validation: ensure at least one EMF file was written.
        // -----------------------------------------------------------------
        int emfCount = Directory.GetFiles(outputDir, "*.emf").Length;
        if (emfCount == 0)
            throw new InvalidOperationException("Extraction failed – no EMF files were created.");

        Console.WriteLine($"Extraction completed. {emfCount} EMF file(s) saved to '{outputDir}'.");
    }
}
