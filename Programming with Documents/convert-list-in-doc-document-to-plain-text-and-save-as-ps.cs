using System;
using Aspose.Words;
using Aspose.Words.Saving;

namespace ListToPlainTextPs
{
    class Program
    {
        static void Main()
        {
            // Path to the source DOC file that contains a list.
            const string inputPath = @"C:\Docs\SourceDocument.doc";

            // Path to the output PostScript file.
            const string outputPath = @"C:\Docs\ListAsPlainText.ps";

            // Load the DOC document.
            Document doc = new Document(inputPath);

            // Ensure that list labels are up‑to‑date so they appear correctly in the output.
            doc.UpdateListLabels();

            // Save the document as PostScript (PS) format.
            // Using the SaveFormat enum directly satisfies the lifecycle rule for saving.
            doc.Save(outputPath, SaveFormat.Ps);
        }
    }
}
