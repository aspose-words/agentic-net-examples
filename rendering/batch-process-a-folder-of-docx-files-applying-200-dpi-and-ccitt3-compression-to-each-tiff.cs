using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a temporary source folder and generate sample DOCX files.
        string sourceFolder = Path.Combine(Path.GetTempPath(), "DocxBatchSource");
        Directory.CreateDirectory(sourceFolder);
        for (int i = 1; i <= 3; i++)
        {
            string docPath = Path.Combine(sourceFolder, $"Sample{i}.docx");
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln($"This is sample document {i}.");
            doc.Save(docPath);
        }

        // Create an output folder for the TIFF files.
        string outputFolder = Path.Combine(Path.GetTempPath(), "TiffBatchOutput");
        Directory.CreateDirectory(outputFolder);

        // Process each DOCX file in the source folder.
        string[] docxFiles = Directory.GetFiles(sourceFolder, "*.docx");
        foreach (string docxFile in docxFiles)
        {
            Document document = new Document(docxFile);

            // Configure image save options for TIFF with 200 DPI and CCITT3 compression.
            ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Tiff)
            {
                Resolution = 200,                     // Sets both DpiX and DpiY.
                TiffCompression = TiffCompression.Ccitt3
            };

            string tiffFileName = Path.GetFileNameWithoutExtension(docxFile) + ".tiff";
            string tiffPath = Path.Combine(outputFolder, tiffFileName);

            document.Save(tiffPath, saveOptions);

            // Verify that the TIFF file was created.
            if (!File.Exists(tiffPath))
                throw new Exception($"Failed to create TIFF file: {tiffPath}");
        }

        // Indicate completion.
        Console.WriteLine($"Processed {docxFiles.Length} DOCX files. TIFF files are saved in: {outputFolder}");
    }
}
