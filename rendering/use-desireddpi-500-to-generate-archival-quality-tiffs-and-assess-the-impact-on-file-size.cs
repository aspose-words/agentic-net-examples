using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a sample document with two pages.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Hello World! This is a sample document for TIFF rendering.");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Second page content for TIFF rendering test.");

        // Render TIFF with default DPI.
        string defaultTiffPath = "output_default.tiff";
        ImageSaveOptions defaultOptions = new ImageSaveOptions(SaveFormat.Tiff);
        doc.Save(defaultTiffPath, defaultOptions);

        // Render TIFF with high DPI (500) using the correct property.
        string highDpiTiffPath = "output_500dpi.tiff";
        ImageSaveOptions highDpiOptions = new ImageSaveOptions(SaveFormat.Tiff);
        highDpiOptions.Resolution = 500; // Set desired DPI.
        doc.Save(highDpiTiffPath, highDpiOptions);

        // Verify that both files were created.
        if (!File.Exists(defaultTiffPath))
            throw new FileNotFoundException("Default DPI TIFF was not created.", defaultTiffPath);
        if (!File.Exists(highDpiTiffPath))
            throw new FileNotFoundException("500 DPI TIFF was not created.", highDpiTiffPath);

        // Compare file sizes.
        long defaultSize = new FileInfo(defaultTiffPath).Length;
        long highDpiSize = new FileInfo(highDpiTiffPath).Length;

        Console.WriteLine($"Default DPI TIFF size: {defaultSize} bytes");
        Console.WriteLine($"500 DPI TIFF size: {highDpiSize} bytes");
        Console.WriteLine($"Size increase: {highDpiSize - defaultSize} bytes ({(double)highDpiSize / defaultSize:P2} of original)");
    }
}
