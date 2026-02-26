using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Math;   // Namespace for OfficeMath nodes

class OfficeMathSplitter
{
    static void Main()
    {
        // Path to the source DOCX file that contains OfficeMath equations.
        string inputPath = @"C:\Docs\InputDocument.docx";

        // Folder where each extracted OfficeMath equation will be saved as a separate DOCX file.
        string outputFolder = @"C:\Docs\ExtractedEquations";

        // Ensure the output directory exists.
        if (!Directory.Exists(outputFolder))
            Directory.CreateDirectory(outputFolder);

        // Load the source document.
        Document srcDoc = new Document(inputPath);

        // Collect all OfficeMath nodes in the document (including those inside tables, paragraphs, etc.).
        NodeCollection mathNodes = srcDoc.GetChildNodes(NodeType.OfficeMath, true);

        // Iterate through each OfficeMath node and save it to its own document.
        int equationIndex = 1;
        foreach (OfficeMath mathNode in mathNodes)
        {
            // Create a new blank document.
            Document newDoc = new Document();

            // Remove the automatically created empty section/paragraphs.
            newDoc.RemoveAllChildren();

            // Build the minimal structure required for a valid Word document.
            Section section = new Section(newDoc);
            newDoc.AppendChild(section);

            Body body = new Body(newDoc);
            section.AppendChild(body);

            Paragraph paragraph = new Paragraph(newDoc);
            body.AppendChild(paragraph);

            // Import the OfficeMath node from the source document into the new document.
            NodeImporter importer = new NodeImporter(srcDoc, newDoc, ImportFormatMode.KeepSourceFormatting);
            Node importedMath = importer.ImportNode(mathNode, true);

            // Append the imported OfficeMath node to the paragraph.
            paragraph.AppendChild(importedMath);

            // Save the new document using the standard Save method.
            string outFile = Path.Combine(outputFolder, $"Equation_{equationIndex}.docx");
            newDoc.Save(outFile);

            equationIndex++;
        }

        Console.WriteLine($"Extracted {equationIndex - 1} OfficeMath equations to \"{outputFolder}\".");
    }
}
