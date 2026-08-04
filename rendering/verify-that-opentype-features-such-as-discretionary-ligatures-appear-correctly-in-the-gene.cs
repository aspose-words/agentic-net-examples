using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write text that contains characters which can form discretionary ligatures (e.g., "fi", "fl").
        builder.Font.Name = "Times New Roman";
        builder.Font.Size = 48;
        builder.Writeln("Office"); // Contains "ff" and "fi" ligatures in many fonts.

        // Configure image save options for TIFF output.
        ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Tiff)
        {
            // Optional: increase resolution for better quality.
            Resolution = 300
            // No need to set PageSet; the default renders all pages.
        };

        // Define the output file path.
        string outputPath = Path.Combine(Environment.CurrentDirectory, "DiscretionaryLigatures.tiff");

        // Save the document as a TIFF image.
        doc.Save(outputPath, saveOptions);

        // Verify that the TIFF file was created and has a non‑zero size.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The TIFF file was not created.");

        long fileSize = new FileInfo(outputPath).Length;
        if (fileSize == 0)
            throw new InvalidOperationException("The generated TIFF file is empty.");

        // Indicate success.
        Console.WriteLine($"TIFF file created successfully at: {outputPath}");
        Console.WriteLine($"File size: {fileSize} bytes");
    }
}
