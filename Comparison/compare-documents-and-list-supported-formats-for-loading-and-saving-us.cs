using System;
using System.IO;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Saving;

namespace AsposeWordsDemo
{
    class Program
    {
        static void Main()
        {
            // Paths to the documents to compare.
            string originalPath = @"Documents\Original.docx";
            string editedPath   = @"Documents\Edited.docx";

            // Load the original and edited documents using the Document(string) constructor.
            Document originalDoc = new Document(originalPath);
            Document editedDoc   = new Document(editedPath);

            // Compare the documents. The revisions will be added to the original document.
            originalDoc.Compare(editedDoc, "Comparer", DateTime.Now);

            // Save the comparison result to a new DOCX file.
            string comparisonResultPath = @"Output\ComparisonResult.docx";
            originalDoc.Save(comparisonResultPath);

            // -----------------------------------------------------------------
            // List all load formats supported by Aspose.Words.
            // -----------------------------------------------------------------
            Console.WriteLine("Supported Load Formats:");
            foreach (LoadFormat loadFormat in Enum.GetValues(typeof(LoadFormat)))
            {
                // Skip the 'Unknown' and 'Auto' entries for clarity.
                if (loadFormat == LoadFormat.Unknown || loadFormat == LoadFormat.Auto)
                    continue;

                Console.WriteLine($"- {loadFormat} ({(int)loadFormat})");
            }

            // -----------------------------------------------------------------
            // List all save formats supported by Aspose.Words.
            // -----------------------------------------------------------------
            Console.WriteLine("\nSupported Save Formats:");
            foreach (SaveFormat saveFormat in Enum.GetValues(typeof(SaveFormat)))
            {
                // Skip the 'Unknown' entry.
                if (saveFormat == SaveFormat.Unknown)
                    continue;

                Console.WriteLine($"- {saveFormat} ({(int)saveFormat})");
            }

            // -----------------------------------------------------------------
            // Demonstrate saving the original DOCX document to a few other formats.
            // -----------------------------------------------------------------
            string outputDir = @"Output";

            // Ensure the output directory exists.
            Directory.CreateDirectory(outputDir);

            // Save as PDF.
            originalDoc.Save(Path.Combine(outputDir, "Original.pdf"), SaveFormat.Pdf);

            // Save as HTML.
            originalDoc.Save(Path.Combine(outputDir, "Original.html"), SaveFormat.Html);

            // Save as plain text.
            originalDoc.Save(Path.Combine(outputDir, "Original.txt"), SaveFormat.Text);

            // Save as ODT.
            originalDoc.Save(Path.Combine(outputDir, "Original.odt"), SaveFormat.Odt);

            Console.WriteLine("\nConversion completed. Check the Output folder for generated files.");
        }
    }
}
