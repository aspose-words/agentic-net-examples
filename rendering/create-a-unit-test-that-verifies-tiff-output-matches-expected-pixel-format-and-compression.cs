using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        const string docPath = "sample.docx";
        const string tiffPath = "output.tiff";

        // Create a sample document with three pages.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Page 1");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Page 2");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Page 3");
        doc.Save(docPath);

        // Render the document to a multipage TIFF with specific compression.
        ImageSaveOptions tiffOptions = new ImageSaveOptions(SaveFormat.Tiff);
        tiffOptions.TiffCompression = TiffCompression.Ccitt4; // Correct enum value
        doc.Save(tiffPath, tiffOptions);

        // Verify that the TIFF file was created.
        if (!File.Exists(tiffPath))
            throw new Exception("TIFF file was not created.");

        // Verify that the file is not empty.
        FileInfo tiffInfo = new FileInfo(tiffPath);
        if (tiffInfo.Length == 0)
            throw new Exception("TIFF file is empty.");

        // Verify that the number of pages in the source document matches the expected output size.
        // Each page should produce a frame; we approximate by ensuring the file size is reasonable.
        int pageCount = doc.PageCount;
        if (tiffInfo.Length < pageCount * 100) // arbitrary minimal size per page
            throw new Exception("TIFF file size is unexpectedly small for the number of pages.");

        Console.WriteLine("TIFF rendering test passed.");
    }
}
