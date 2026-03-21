using System;
using System.IO;
using Aspose.Words;

class HyphenateBatch
{
    static void Main()
    {
        // Folder that contains the source Word documents (relative to the executable).
        string sourceFolder = Path.Combine(Environment.CurrentDirectory, "Input");

        // Folder where the hyphenated PDFs will be saved (relative to the executable).
        string outputFolder = Path.Combine(Environment.CurrentDirectory, "Output");

        // Ensure both directories exist.
        Directory.CreateDirectory(sourceFolder);
        Directory.CreateDirectory(outputFolder);

        // Retrieve all .docx files from the source folder.
        string[] docFiles = Directory.GetFiles(sourceFolder, "*.docx");

        if (docFiles.Length == 0)
        {
            Console.WriteLine($"No .docx files found in \"{sourceFolder}\". Place documents there and rerun.");
            return;
        }

        foreach (string docPath in docFiles)
        {
            // Load the document from the file system.
            Document doc = new Document(docPath);

            // Enable automatic hyphenation for the document.
            doc.HyphenationOptions.AutoHyphenation = true;

            // Optional: fine‑tune hyphenation behavior.
            doc.HyphenationOptions.ConsecutiveHyphenLimit = 2; // limit consecutive hyphenated lines
            doc.HyphenationOptions.HyphenationZone = 720;      // 0.5 inch zone
            doc.HyphenationOptions.HyphenateCaps = true;      // hyphenate all‑caps words

            // Construct the output PDF file name.
            string pdfFileName = Path.GetFileNameWithoutExtension(docPath) + ".pdf";
            string pdfPath = Path.Combine(outputFolder, pdfFileName);

            // Save the document as PDF; the .pdf extension selects the PDF format.
            doc.Save(pdfPath);
            Console.WriteLine($"Saved hyphenated PDF: {pdfPath}");
        }
    }
}
