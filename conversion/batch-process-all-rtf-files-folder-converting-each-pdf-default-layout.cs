using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Folder containing the source RTF files.
        string inputFolder = @"C:\InputRtf";

        // Folder where the resulting PDF files will be placed.
        string outputFolder = @"C:\OutputPdf";

        // Ensure the output directory exists.
        Directory.CreateDirectory(outputFolder);

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
    }
}
