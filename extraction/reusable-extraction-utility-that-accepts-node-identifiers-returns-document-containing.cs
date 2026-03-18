using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Saving;

namespace ExtractionUtility
{
    /// <summary>
    /// Provides methods to extract nodes identified by their CustomNodeId
    /// from a source document and return a new document containing those nodes.
    /// </summary>
    public static class NodeExtractor
    {
        /// <summary>
        /// Extracts nodes from a document loaded from the specified file path.
        /// </summary>
        /// <param name="sourcePath">Path to the source Word document.</param>
        /// <param name="nodeIds">Collection of CustomNodeId values identifying the nodes to extract.</param>
        /// <returns>A new <see cref="Document"/> containing the extracted nodes.</returns>
        public static Document Extract(string sourcePath, IEnumerable<int> nodeIds)
        {
            if (string.IsNullOrEmpty(sourcePath))
                throw new ArgumentException("Source path must be provided.", nameof(sourcePath));

            Document srcDoc = new Document(sourcePath);
            return Extract(srcDoc, nodeIds);
        }

        /// <summary>
        /// Extracts nodes from an already loaded document.
        /// </summary>
        /// <param name="srcDoc">The source document.</param>
        /// <param name="nodeIds">Collection of CustomNodeId values identifying the nodes to extract.</param>
        /// <returns>A new <see cref="Document"/> containing the extracted nodes.</returns>
        public static Document Extract(Document srcDoc, IEnumerable<int> nodeIds)
        {
            if (srcDoc == null) throw new ArgumentNullException(nameof(srcDoc));
            if (nodeIds == null) throw new ArgumentNullException(nameof(nodeIds));

            // Create an empty destination document with the minimal required structure.
            Document dstDoc = new Document();
            dstDoc.RemoveAllChildren(); // start from a clean slate
            Section dstSection = new Section(dstDoc);
            dstDoc.AppendChild(dstSection);
            Body dstBody = new Body(dstDoc);
            dstSection.AppendChild(dstBody);
            // Add a placeholder paragraph – it will be removed once the first real node is imported.
            dstBody.AppendChild(new Paragraph(dstDoc));

            // Prepare a NodeImporter for efficient repeated imports.
            NodeImporter importer = new NodeImporter(srcDoc, dstDoc, ImportFormatMode.KeepSourceFormatting);

            foreach (int id in nodeIds)
            {
                // Search the entire source document for a node with the matching CustomNodeId.
                Node srcNode = srcDoc.GetChildNodes(NodeType.Any, true)
                                     .FirstOrDefault(n => n.CustomNodeId == id);

                if (srcNode == null)
                    throw new InvalidOperationException($"Node with CustomNodeId {id} was not found in the source document.");

                // Import the node (deep clone) into the destination document.
                Node importedNode = importer.ImportNode(srcNode, true);

                // Remove the placeholder paragraph if it is still present.
                if (dstBody.LastParagraph != null && string.IsNullOrWhiteSpace(dstBody.LastParagraph.GetText()))
                {
                    dstBody.LastParagraph.Remove();
                }

                dstBody.AppendChild(importedNode);
            }

            return dstDoc;
        }

        /// <summary>
        /// Saves the extracted document to the specified file path using the given format.
        /// </summary>
        /// <param name="doc">Document to save.</param>
        /// <param name="outputPath">Destination file path.</param>
        /// <param name="format">Save format (default is DOCX).</param>
        public static void Save(Document doc, string outputPath, SaveFormat format = SaveFormat.Docx)
        {
            if (doc == null) throw new ArgumentNullException(nameof(doc));
            if (string.IsNullOrEmpty(outputPath)) throw new ArgumentException("Output path must be provided.", nameof(outputPath));

            doc.Save(outputPath, format);
        }
    }

    /// <summary>
    /// Simple entry point to make the project a valid console application.
    /// Demonstrates how to call the extraction utility.
    /// </summary>
    public static class Program
    {
        public static void Main()
        {
            // Example usage – adjust paths and IDs as needed.
            const string sourcePath = "source.docx";
            var ids = new List<int> { 1, 2, 3 };

            try
            {
                Document extracted = NodeExtractor.Extract(sourcePath, ids);
                NodeExtractor.Save(extracted, "extracted.docx");
                Console.WriteLine("Extraction completed successfully.");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
