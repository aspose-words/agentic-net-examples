using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Markup; // Required for StructuredDocumentTag
using Aspose.Words.Saving;

namespace DocumentConversionDemo
{
    public class DotxConverter
    {
        /// <summary>
        /// Loads a DOTX template, replaces all content controls (StructuredDocumentTag) with a placeholder text,
        /// and saves the result as a DOCX file.
        /// </summary>
        /// <param name="inputDotxPath">Full path to the source .dotx file.</param>
        /// <param name="outputDocxPath">Full path where the converted .docx will be saved.</param>
        public static void ConvertDotxToDocx(string inputDotxPath, string outputDocxPath)
        {
            // Load the DOTX document from the file system.
            Document templateDoc = new Document(inputDotxPath);

            // Iterate over all content controls (StructuredDocumentTag nodes) in the document.
            // Replace each control's content with a simple placeholder text.
            foreach (StructuredDocumentTag sdt in templateDoc.GetChildNodes(NodeType.StructuredDocumentTag, true))
            {
                // Clear any existing child nodes inside the content control.
                sdt.RemoveAllChildren();

                // Insert a new Run node with the replacement text.
                sdt.AppendChild(new Run(templateDoc, "Placeholder"));
            }

            // Save the modified document as a DOCX file.
            // The Save method automatically determines the format from the file extension,
            // but we explicitly specify SaveFormat.Docx for clarity.
            templateDoc.Save(outputDocxPath, SaveFormat.Docx);
        }

        // Example usage.
        public static void Main()
        {
            // Paths can be adjusted as needed.
            string inputPath = @"C:\Docs\Template.dotx";
            string outputPath = @"C:\Docs\Result.docx";

            // Ensure the input file exists before attempting conversion.
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            try
            {
                ConvertDotxToDocx(inputPath, outputPath);
                Console.WriteLine($"Conversion successful. Output saved to: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred during conversion: {ex.Message}");
            }
        }
    }
}
