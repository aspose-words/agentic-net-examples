using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Loading;
using Aspose.Words.Drawing;

class ExtractVectorImagesToEmf
{
    static void Main()
    {
        // Path to the source DOC file (placed in the program's working directory).
        string inputPath = Path.Combine(Environment.CurrentDirectory, "source.doc");

        // Ensure a document exists at the specified path.
        if (!File.Exists(inputPath))
        {
            // Create a simple document so the example can run without external files.
            Document newDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(newDoc);
            builder.Writeln("This is a placeholder document created for the example.");
            newDoc.Save(inputPath);
        }

        // Folder where the extracted EMF files will be saved.
        string outputFolder = Path.Combine(Environment.CurrentDirectory, "ExtractedEmf");
        Directory.CreateDirectory(outputFolder);

        // Load the document without converting metafiles to raster images.
        LoadOptions loadOptions = new LoadOptions { ConvertMetafilesToPng = false };
        Document doc = new Document(inputPath, loadOptions);

        int imageCounter = 0;

        // Iterate over all Shape nodes in the document.
        foreach (Shape shape in doc.GetChildNodes(NodeType.Shape, true))
        {
            if (!shape.HasImage) continue;

            ImageType imgType = shape.ImageData.ImageType;
            if (imgType == ImageType.Wmf || imgType == ImageType.Emf)
            {
                string outFile = Path.Combine(outputFolder, $"vector_{imageCounter}.emf");
                shape.ImageData.Save(outFile);
                imageCounter++;
            }
        }

        Console.WriteLine($"Extracted {imageCounter} vector image(s) to EMF format.");
    }
}
