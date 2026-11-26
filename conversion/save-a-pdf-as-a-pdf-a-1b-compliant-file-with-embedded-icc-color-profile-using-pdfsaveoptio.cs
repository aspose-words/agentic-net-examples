using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a sample Word document in memory.
        Document source = new Document();
        DocumentBuilder builder = new DocumentBuilder(source);
        builder.Writeln("This is a sample document to be saved as PDF/A‑1b with an embedded ICC profile.");

        // Configure PDF save options for PDF/A‑1b compliance.
        // Note: In the current Aspose.Words version the IccProfileId property is not available.
        // The example demonstrates the required compliance setting; embedding a specific ICC profile
        // would require using IccProfilePath or IccProfileBytes if a profile file is supplied.
        PdfSaveOptions saveOptions = new PdfSaveOptions
        {
            Compliance = PdfCompliance.PdfA1b
        };

        // Define the output PDF file path.
        string outputPath = "output.pdf";

        // Save the document as PDF/A‑1b with the specified options.
        source.Save(outputPath, saveOptions);

        // Verify that the PDF file was created.
        if (!File.Exists(outputPath))
        {
            throw new InvalidOperationException("The PDF/A‑1b file was not created as expected.");
        }

        // Ensure the file is not empty.
        FileInfo info = new FileInfo(outputPath);
        if (info.Length == 0)
        {
            throw new InvalidOperationException("The created PDF/A‑1b file is empty.");
        }

        // Indicate successful completion.
        Console.WriteLine("PDF/A‑1b file created successfully: " + Path.GetFullPath(outputPath));
    }
}
