using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

namespace ExtractionExample
{
    public class Program
    {
        public static void Main()
        {
            // -----------------------------------------------------------------
            // Create a sample source document with several paragraphs.
            // -----------------------------------------------------------------
            Document sourceDoc = new Document();
            DocumentBuilder srcBuilder = new DocumentBuilder(sourceDoc);
            srcBuilder.Writeln("First paragraph.");
            srcBuilder.Writeln("Second paragraph.");
            srcBuilder.Writeln("Third paragraph.");

            const string sourcePath = "source.docx";
            sourceDoc.Save(sourcePath);

            // -----------------------------------------------------------------
            // Load the source document.
            // -----------------------------------------------------------------
            Document loadedSource = new Document(sourcePath);

            // Extract all paragraph nodes from the source document.
            NodeCollection paragraphNodes = loadedSource.GetChildNodes(NodeType.Paragraph, true);
            if (paragraphNodes == null || paragraphNodes.Count == 0)
                throw new InvalidOperationException("No paragraphs were found in the source document.");

            // -----------------------------------------------------------------
            // Create a new destination document with a clean structure.
            // -----------------------------------------------------------------
            Document resultDoc = new Document();
            resultDoc.RemoveAllChildren(); // Remove the default empty section/body.

            Section section = new Section(resultDoc);
            resultDoc.AppendChild(section);
            Body body = new Body(resultDoc);
            section.AppendChild(body);

            // Add an existing paragraph that will appear after the prepended content.
            Paragraph existingParagraph = new Paragraph(resultDoc);
            existingParagraph.AppendChild(new Run(resultDoc, "Existing content in the new document."));
            body.AppendChild(existingParagraph);

            // -----------------------------------------------------------------
            // Import (clone) the extracted paragraphs into the new document.
            // Use NodeImporter to transfer nodes between documents.
            // -----------------------------------------------------------------
            NodeImporter importer = new NodeImporter(loadedSource, resultDoc, ImportFormatMode.KeepSourceFormatting);

            foreach (Paragraph para in paragraphNodes)
            {
                // Import the paragraph (deep clone) into the destination document.
                Node importedNode = importer.ImportNode(para, true);
                // Insert before the existing paragraph so the extracted content appears at the start.
                body.InsertBefore(importedNode, existingParagraph);
            }

            // -----------------------------------------------------------------
            // Save the resulting document.
            // -----------------------------------------------------------------
            const string resultPath = "result.docx";
            resultDoc.Save(resultPath, SaveFormat.Docx);

            // Verify that the output file was created.
            if (!File.Exists(resultPath))
                throw new InvalidOperationException("The result document was not created.");

            // Optional cleanup (comment out if you want to inspect the files).
            // File.Delete(sourcePath);
            // File.Delete(resultPath);
        }
    }
}
