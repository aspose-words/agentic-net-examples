using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a new document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Use a font that contains discretionary ligatures (e.g., Times New Roman).
        builder.Font.Name = "Times New Roman";
        builder.Font.Size = 48;

        // Write text that includes a discretionary ligature (the "fi" sequence).
        builder.Writeln("office"); // The "fi" in "office" can be rendered as a ligature.

        // Define the output TIFF file path.
        string tiffPath = "output.tiff";

        // Configure image save options for TIFF.
        ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Tiff)
        {
            // Render each page as a separate TIFF frame (default behavior).
            // No additional options are required for this verification.
        };

        // Save the document as a multipage TIFF.
        doc.Save(tiffPath, saveOptions);

        // Verify that the TIFF file was created.
        if (!File.Exists(tiffPath))
            throw new InvalidOperationException($"TIFF file was not created at '{tiffPath}'.");

        // Verify that the file size is greater than zero (indicates content was written).
        FileInfo info = new FileInfo(tiffPath);
        if (info.Length == 0)
            throw new InvalidOperationException("Generated TIFF file is empty.");

        // Verify that the source document has the expected page count.
        // For this simple example the document should be a single page.
        int pageCount = doc.PageCount;
        if (pageCount != 1)
            throw new InvalidOperationException($"Unexpected page count: {pageCount}. Expected 1.");

        // If all checks pass, indicate success.
        Console.WriteLine("TIFF rendering completed successfully. File size: " + info.Length + " bytes.");
    }
}
