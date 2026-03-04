using System;
using Aspose.Words;
using Aspose.Words.Saving;

namespace DocumentConversionExample
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the source DOTM (macro‑enabled template) file.
            string inputPath = @"C:\Docs\Template.dotm";

            // Path where the resulting MHTML (Web archive) file will be saved.
            string outputPath = @"C:\Docs\Template.mht";

            // Load the DOTM document from the file system.
            Document doc = new Document(inputPath);

            // Save the loaded document in MHTML format.
            // The SaveFormat enumeration value Mhtml corresponds to the MHTML (Web archive) format.
            doc.Save(outputPath, SaveFormat.Mhtml);
        }
    }
}
