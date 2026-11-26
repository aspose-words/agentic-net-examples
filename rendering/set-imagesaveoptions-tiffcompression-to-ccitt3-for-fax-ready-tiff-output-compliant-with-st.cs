using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a simple document with some text.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Hello, Aspose.Words!");

        // Configure image save options for TIFF with CCITT3 compression (fax‑ready).
        ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Tiff);
        saveOptions.TiffCompression = TiffCompression.Ccitt3;

        // Define output file path.
        string outputPath = "fax_ready_output.tiff";

        // Save the document as a TIFF image using the specified options.
        doc.Save(outputPath, saveOptions);

        // Verify that the TIFF file was created successfully.
        if (!File.Exists(outputPath))
        {
            throw new InvalidOperationException($"Failed to create TIFF file at '{outputPath}'.");
        }

        // Optional: ensure the file is not empty.
        FileInfo info = new FileInfo(outputPath);
        if (info.Length == 0)
        {
            throw new InvalidOperationException("The generated TIFF file is empty.");
        }

        // Indicate successful completion.
        Console.WriteLine($"TIFF file saved successfully to '{outputPath}'.");
    }
}
