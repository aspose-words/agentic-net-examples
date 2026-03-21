using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Base directory of the application.
        string baseDir = AppContext.BaseDirectory;

        // Folder containing the source RTF files.
        string inputFolder = Path.Combine(baseDir, "InputRtf");

        // Folder where the resulting PDF files will be placed.
        string outputFolder = Path.Combine(baseDir, "OutputPdf");

        // Ensure both directories exist.
        Directory.CreateDirectory(inputFolder);
        Directory.CreateDirectory(outputFolder);

        // If there are no RTF files, create a simple sample file.
        if (Directory.GetFiles(inputFolder, "*.rtf").Length == 0)
        {
            string sampleRtfPath = Path.Combine(inputFolder, "Sample.rtf");
            var sampleDoc = new Document();
            var builder = new DocumentBuilder(sampleDoc);
            builder.Writeln("This is a sample RTF document generated for the batch conversion example.");
            sampleDoc.Save(sampleRtfPath, SaveFormat.Rtf);
        }

        // Process each .rtf file in the input folder.
        foreach (string rtfPath in Directory.GetFiles(inputFolder, "*.rtf"))
        {
            // Load the RTF document.
            Document doc = new Document(rtfPath);

            // Build the output PDF file name.
            string pdfFileName = Path.GetFileNameWithoutExtension(rtfPath) + ".pdf";
            string pdfPath = Path.Combine(outputFolder, pdfFileName);

            // Save the document as PDF using the default layout and options.
            doc.Save(pdfPath, SaveFormat.Pdf);
        }

        Console.WriteLine($"Conversion completed. PDFs are located in: {outputFolder}");
    }
}
