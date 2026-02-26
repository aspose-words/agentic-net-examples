using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

class DownsampleDocument
{
    static void Main()
    {
        // Paths to the source and destination files.
        string dataDir = @"C:\Docs\";
        string inputPath = Path.Combine(dataDir, "Input.docx");
        string outputPath = Path.Combine(dataDir, "Output_Downsampled.pdf");

        // Load the Word document.
        Document doc = new Document(inputPath);

        // Create PDF save options that will be used to downsample images.
        PdfSaveOptions saveOptions = new PdfSaveOptions();

        // Enable downsampling and configure the target resolution (ppi) and threshold.
        // Images with a resolution higher than the threshold will be reduced to the target resolution.
        saveOptions.DownsampleOptions.DownsampleImages = true;   // Ensure downsampling is active.
        saveOptions.DownsampleOptions.Resolution = 72;          // Target resolution in pixels per inch.
        saveOptions.DownsampleOptions.ResolutionThreshold = 150; // Only downsample images >150 ppi.

        // Save the document as a PDF with the specified downsampling options.
        doc.Save(outputPath, saveOptions);
    }
}
