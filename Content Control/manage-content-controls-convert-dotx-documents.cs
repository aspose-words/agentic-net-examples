using System;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Markup;

namespace AsposeWordsDotxConversion
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the source DOTX template.
            string dotxPath = @"C:\Templates\SampleTemplate.dotx";

            // Load the DOTX document. The constructor automatically detects the format.
            Document doc = new Document(dotxPath);

            // Iterate through all content controls (StructuredDocumentTag nodes) in the document.
            NodeCollection contentControls = doc.GetChildNodes(NodeType.StructuredDocumentTag, true);
            foreach (StructuredDocumentTag sdt in contentControls)
            {
                // Clear any existing content inside the content control.
                sdt.RemoveAllChildren();

                // Insert the placeholder text as a new Run node.
                Run placeholder = new Run(doc, $"[Placeholder for {sdt.Title}]");
                sdt.AppendChild(placeholder);
            }

            // Prepare save options to convert the document to DOCX format.
            OoxmlSaveOptions saveOptions = new OoxmlSaveOptions(SaveFormat.Docx)
            {
                // Ensure that the generator name is not embedded (optional).
                ExportGeneratorName = false,
                // Keep the original formatting without pretty formatting (optional).
                PrettyFormat = false
            };

            // Path for the output DOCX file.
            string outputPath = @"C:\Output\ConvertedDocument.docx";

            // Save the modified document using the specified options.
            doc.Save(outputPath, saveOptions);

            Console.WriteLine("DOTX document has been converted and saved as DOCX.");
        }
    }
}
