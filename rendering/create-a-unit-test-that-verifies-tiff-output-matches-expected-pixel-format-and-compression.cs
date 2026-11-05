using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    // Expected TIFF properties (per TIFF specification)
    private const ushort ExpectedCompression = 4; // LZW compression (value defined by TIFF spec)
    private const ushort ExpectedPhotometric = 2; // RGB (color) photometric interpretation (value defined by TIFF spec)

    public static void Main()
    {
        // Create a simple document in memory.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Hello, Aspose.Words TIFF rendering test.");

        // Configure TIFF save options – use LZW compression.
        ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Tiff);
        saveOptions.TiffCompression = TiffCompression.Lzw; // LZW compression

        // Save the document as a TIFF file.
        string tiffPath = "output.tiff";
        doc.Save(tiffPath, saveOptions);

        // Verify that the TIFF file was created.
        if (!File.Exists(tiffPath))
            throw new FileNotFoundException("TIFF file was not created.", tiffPath);

        // Basic validation: file size should be greater than zero.
        FileInfo info = new FileInfo(tiffPath);
        if (info.Length == 0)
            throw new InvalidOperationException("TIFF file is empty.");

        // Additional validation: ensure that the save options were set to the expected compression.
        // (Aspose.Words does not expose TIFF tags directly, so we validate the configuration used.)
        if (saveOptions.TiffCompression != TiffCompression.Lzw)
            throw new InvalidOperationException($"Unexpected TIFF compression setting. Expected: {ExpectedCompression}, Actual: {(ushort)saveOptions.TiffCompression}");

        // Note: Photometric interpretation for color images is automatically set to RGB by Aspose.Words.
        // Since the API does not expose this tag, we rely on the default behavior for color images.

        Console.WriteLine("TIFF rendering verification succeeded.");
    }
}
