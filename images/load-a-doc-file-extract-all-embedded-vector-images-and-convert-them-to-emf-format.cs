using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Loading;
using Aspose.Words.Saving;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Deterministic file names.
        const string vectorEmfPath = "vector.emf";
        const string docPath = "sample.doc";

        // -----------------------------------------------------------------
        // 1. Create a simple EMF vector image using Aspose.Words rendering.
        // -----------------------------------------------------------------
        // Create a temporary document with some content.
        Document tempDoc = new Document();
        DocumentBuilder tempBuilder = new DocumentBuilder(tempDoc);
        tempBuilder.Writeln("Sample vector content for EMF image.");

        // Save the first page of the temporary document as an EMF file.
        ImageSaveOptions emfSaveOptions = new ImageSaveOptions(SaveFormat.Emf);
        tempDoc.Save(vectorEmfPath, emfSaveOptions);

        // Verify that the EMF file was created.
        if (!File.Exists(vectorEmfPath))
            throw new InvalidOperationException("Failed to create the EMF image.");

        // -----------------------------------------------------------------
        // 2. Create a DOC file and embed the EMF image.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert the EMF image into the document.
        Shape insertedShape = builder.InsertImage(vectorEmfPath);

        // Ensure the shape actually contains an image.
        if (!insertedShape.HasImage)
            throw new InvalidOperationException("The inserted shape does not contain an image.");

        // Save the document.
        doc.Save(docPath);

        // Verify that the DOC file was saved.
        if (!File.Exists(docPath))
            throw new InvalidOperationException("Failed to save the DOC file.");

        // -----------------------------------------------------------------
        // 3. Load the DOC file without converting metafiles to PNG.
        // -----------------------------------------------------------------
        LoadOptions loadOptions = new LoadOptions
        {
            ConvertMetafilesToPng = false // Preserve vector images.
        };

        Document loadedDoc = new Document(docPath, loadOptions);

        // -----------------------------------------------------------------
        // 4. Extract all embedded vector images (EMF or WMF) and save as EMF.
        // -----------------------------------------------------------------
        int extractedCount = 0;
        foreach (Shape shape in loadedDoc.GetChildNodes(NodeType.Shape, true).OfType<Shape>())
        {
            if (!shape.HasImage)
                continue;

            ImageType imgType = shape.ImageData.ImageType;
            if (imgType == ImageType.Emf || imgType == ImageType.Wmf)
            {
                string outFile = $"extracted_vector_{extractedCount}.emf";

                // Save the image data. Aspose.Words will write EMF when the extension is .emf.
                shape.ImageData.Save(outFile);
                if (!File.Exists(outFile))
                    throw new InvalidOperationException($"Failed to save extracted image '{outFile}'.");

                extractedCount++;
            }
        }

        // -----------------------------------------------------------------
        // 5. Validate that at least one EMF file was created.
        // -----------------------------------------------------------------
        if (extractedCount == 0)
            throw new InvalidOperationException("No vector images were extracted from the document.");

        Console.WriteLine($"Successfully extracted {extractedCount} vector image(s).");
    }
}
