using System;
using System.IO;
using System.IO.Compression;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a simple ZIP archive in memory with two files.
        byte[] zipBytes;
        using (var ms = new MemoryStream())
        {
            using (var archive = new ZipArchive(ms, ZipArchiveMode.Create, true))
            {
                var entry1 = archive.CreateEntry("File1.txt");
                using (var entryStream = entry1.Open())
                using (var writer = new StreamWriter(entryStream))
                {
                    writer.Write("Content of file 1");
                }

                var entry2 = archive.CreateEntry("File2.txt");
                using (var entryStream = entry2.Open())
                using (var writer = new StreamWriter(entryStream))
                {
                    writer.Write("Content of file 2");
                }
            }
            zipBytes = ms.ToArray();
        }

        // Create a new document and insert the ZIP as an OLE Package.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        using (MemoryStream zipStream = new MemoryStream(zipBytes))
        {
            // Insert the OLE object as a package and display it as an icon.
            Shape oleShape = builder.InsertOleObject(zipStream, "Package", true, null);

            // Set package properties for identification.
            oleShape.OleFormat.OlePackage.FileName = "SampleArchive.zip";
            oleShape.OleFormat.OlePackage.DisplayName = "Sample Archive";
        }

        // Iterate through all shapes to find OLE Packages and enumerate their details.
        var shapes = doc.GetChildNodes(NodeType.Shape, true).OfType<Shape>();
        foreach (var shape in shapes)
        {
            OleFormat oleFormat = shape.OleFormat;
            if (oleFormat != null && oleFormat.OlePackage != null)
            {
                OlePackage package = oleFormat.OlePackage;
                Console.WriteLine("Found OLE Package:");
                Console.WriteLine($"  FileName   : {package.FileName}");
                Console.WriteLine($"  DisplayName: {package.DisplayName}");

                // Optionally, list the raw OLE entries (e.g., the embedded file streams).
                // The package itself stores the original file; we can extract it to inspect its contents.
                using (MemoryStream extracted = new MemoryStream())
                {
                    oleFormat.Save(extracted);
                    extracted.Position = 0;
                    using (var archive = new ZipArchive(extracted, ZipArchiveMode.Read))
                    {
                        Console.WriteLine("  Contained entries:");
                        foreach (var entry in archive.Entries)
                        {
                            Console.WriteLine($"    - {entry.FullName} ({entry.Length} bytes)");
                        }
                    }
                }
            }
        }

        // Save the document to a temporary file (optional, demonstrates saving).
        string outputPath = Path.Combine(Path.GetTempPath(), "OlePackageDemo.docx");
        doc.Save(outputPath);
        Console.WriteLine($"Document saved to: {outputPath}");
    }
}
