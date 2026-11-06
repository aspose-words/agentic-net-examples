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

        // Use a font that contains discretionary ligatures (e.g., Times New Roman).
        builder.Font.Name = "Times New Roman";

        // Write text that can form a discretionary ligature such as "fi".
        builder.Writeln("office");

        // Write the Unicode ligature character directly (ﬁ) to ensure it is rendered as a single glyph.
        builder.Writeln("\uFB01"); // Unicode for 'ﬁ'

        // Path for the output TIFF file.
        string tiffPath = "output.tiff";

        // Configure image save options for TIFF.
        // No need to set PageCount – by default all pages are rendered.
        ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Tiff);

        // Save the document as a multipage TIFF.
        doc.Save(tiffPath, saveOptions);

        // Verify that the TIFF file was created and is not empty.
        if (!File.Exists(tiffPath))
            throw new Exception("TIFF file was not created.");

        FileInfo info = new FileInfo(tiffPath);
        if (info.Length == 0)
            throw new Exception("TIFF file is empty.");

        // Output a simple confirmation.
        Console.WriteLine($"TIFF rendered successfully. Pages: {doc.PageCount}, File size: {info.Length} bytes.");
    }
}
