using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fonts;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define a folder for temporary files.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // -----------------------------------------------------------------
        // 1. Create a simple DOCX document and save it locally.
        // -----------------------------------------------------------------
        string sourceDocPath = Path.Combine(artifactsDir, "Sample.docx");
        Document docToSave = new Document();
        DocumentBuilder builder = new DocumentBuilder(docToSave);
        builder.Writeln("This is a sample document used for rendering to a 1‑bpp TIFF image.");
        docToSave.Save(sourceDocPath);

        // -----------------------------------------------------------------
        // 2. Load the DOCX document.
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(sourceDocPath);

        // -----------------------------------------------------------------
        // 3. Configure FontSettings (no OpenType feature toggling needed).
        // -----------------------------------------------------------------
        FontSettings fontSettings = new FontSettings();
        loadedDoc.FontSettings = fontSettings;

        // -----------------------------------------------------------------
        // 4. Set up ImageSaveOptions for 1‑bpp TIFF output.
        // -----------------------------------------------------------------
        ImageSaveOptions tiffOptions = new ImageSaveOptions(SaveFormat.Tiff)
        {
            // Use CCITT Group 4 compression, suitable for 1‑bpp images.
            TiffCompression = TiffCompression.Ccitt4,

            // Force the pixel format to 1‑bpp indexed.
            PixelFormat = ImagePixelFormat.Format1bppIndexed,

            // Set a reasonable resolution.
            Resolution = 300,

            // Disable anti‑aliasing and high‑quality rendering to keep the file size minimal.
            UseAntiAliasing = false,
            UseHighQualityRendering = false
        };

        // -----------------------------------------------------------------
        // 5. Render the document to a single-page TIFF file.
        // -----------------------------------------------------------------
        string outputTiffPath = Path.Combine(artifactsDir, "Rendered.tiff");
        loadedDoc.Save(outputTiffPath, tiffOptions);

        // -----------------------------------------------------------------
        // 6. Verify that the output file was created.
        // -----------------------------------------------------------------
        if (!File.Exists(outputTiffPath))
            throw new InvalidOperationException("The TIFF file was not created.");

        // Output the size of the generated file.
        long fileSize = new FileInfo(outputTiffPath).Length;
        Console.WriteLine($"TIFF file created at: {outputTiffPath}");
        Console.WriteLine($"File size: {fileSize} bytes");
    }
}
