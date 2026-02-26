using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

class PdfComplianceExample
{
    static void Main()
    {
        // Define the directories used in the example.
        // In a real project these could be read from configuration or passed as arguments.
        const string MyDir = @"C:\Input\";        // Folder that contains the source document.
        const string ArtifactsDir = @"C:\Output\"; // Folder where the resulting PDF will be saved.

        // Ensure the output directory exists.
        Directory.CreateDirectory(ArtifactsDir);

        // Path to the source document.
        string inputPath = Path.Combine(MyDir, "input.docx");

        // Load the document using the standard Document constructor.
        Document doc = new Document(inputPath);

        // Determine the desired PDF compliance level.
        // This could be based on user input, configuration, etc.
        // For demonstration, we set it to PDF/A‑1b.
        PdfCompliance targetCompliance = PdfCompliance.PdfA1b;

        // Create a PdfSaveOptions instance via the provided factory method.
        PdfSaveOptions saveOptions = (PdfSaveOptions)SaveOptions.CreateSaveOptions(SaveFormat.Pdf);

        // Apply the chosen compliance level.
        saveOptions.Compliance = targetCompliance;

        // Save the document as PDF using the configured options.
        string outputPath = Path.Combine(ArtifactsDir, "output.pdf");
        doc.Save(outputPath, saveOptions);

        Console.WriteLine($"Document saved to: {outputPath}\nCompliance: {targetCompliance}");
    }
}
