using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Drawing;
using Aspose.Words.Loading;
using Aspose.Words.Saving;
using Aspose.Words.Replacing;
using Aspose.Words.Fonts;
using Aspose.Drawing; // Required package, not used directly but included per requirements
using Newtonsoft.Json; // Required package, not used directly but included per requirements

namespace FontReplacementExample
{
    public class Program
    {
        public static void Main()
        {
            // Paths for the sample and result documents
            string originalPath = "Original.docx";
            string resultPath = "Modified.docx";

            // Create a sample document with two different fonts
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // First paragraph uses the font we want to replace
            builder.Font.Name = "Arial";
            builder.Writeln("This paragraph is in Arial font.");

            // Second paragraph uses a different font (should remain unchanged)
            builder.Font.Name = "Times New Roman";
            builder.Writeln("This paragraph is in Times New Roman font.");

            // Save the original document
            doc.Save(originalPath, SaveFormat.Docx);

            // Load the document back (demonstrates load rule usage)
            Document loadedDoc = new Document(originalPath);

            // Replace all occurrences of the source font with the target font
            string sourceFont = "Arial";
            string targetFont = "Calibri";

            foreach (Run run in loadedDoc.GetChildNodes(NodeType.Run, true))
            {
                if (string.Equals(run.Font.Name, sourceFont, StringComparison.OrdinalIgnoreCase))
                {
                    run.Font.Name = targetFont;
                }
            }

            // Validate that the replacement was successful
            bool replacementSucceeded = true;
            foreach (Run run in loadedDoc.GetChildNodes(NodeType.Run, true))
            {
                if (string.Equals(run.Font.Name, sourceFont, StringComparison.OrdinalIgnoreCase))
                {
                    replacementSucceeded = false;
                    break;
                }
            }

            // Save the modified document
            loadedDoc.Save(resultPath, SaveFormat.Docx);

            // Validate that the output file exists
            bool fileExists = File.Exists(resultPath);
            Console.WriteLine($"Font replacement completed. Output file exists: {fileExists}. Replacement succeeded: {replacementSucceeded}");
        }
    }
}
