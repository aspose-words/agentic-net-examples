using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Define file paths in the current working directory.
        string sourcePath = Path.Combine(Directory.GetCurrentDirectory(), "Source.docx");
        string resultPath = Path.Combine(Directory.GetCurrentDirectory(), "Extracted.docx");

        // -----------------------------------------------------------------
        // 1. Create a sample source document with several paragraphs.
        // -----------------------------------------------------------------
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);

        builder.Writeln("Paragraph 1");
        builder.Writeln("Paragraph 2");
        builder.Writeln("Paragraph 3");
        builder.Writeln("Paragraph 4");
        builder.Writeln("Paragraph 5");

        sourceDoc.Save(sourcePath);

        // -----------------------------------------------------------------
        // 2. Load the document from the file system.
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(sourcePath);

        // -----------------------------------------------------------------
        // 3. Locate the start and end paragraphs (by index in this example).
        //    We will extract paragraphs 2 through 4 inclusive.
        // -----------------------------------------------------------------
        NodeCollection allParagraphs = loadedDoc.GetChildNodes(NodeType.Paragraph, true);
        int startIndex = 1; // second paragraph (zero‑based)
        int endIndex = 3;   // fourth paragraph

        if (startIndex < 0 || endIndex >= allParagraphs.Count || startIndex > endIndex)
            throw new InvalidOperationException("Invalid paragraph range specified.");

        // -----------------------------------------------------------------
        // 4. Prepare the destination document.
        // -----------------------------------------------------------------
        Document destDoc = new Document();
        destDoc.RemoveAllChildren(); // clear the default section/paragraph

        Section destSection = new Section(destDoc);
        destDoc.AppendChild(destSection);

        Body destBody = new Body(destDoc);
        destSection.AppendChild(destBody);

        // -----------------------------------------------------------------
        // 5. Import the selected paragraphs into the new document.
        // -----------------------------------------------------------------
        NodeImporter importer = new NodeImporter(loadedDoc, destDoc, ImportFormatMode.KeepSourceFormatting);

        for (int i = startIndex; i <= endIndex; i++)
        {
            Paragraph srcParagraph = (Paragraph)allParagraphs[i];
            Node importedNode = importer.ImportNode(srcParagraph, true);
            destBody.AppendChild(importedNode);
        }

        // -----------------------------------------------------------------
        // 6. Save the extracted content as a new DOCX file.
        // -----------------------------------------------------------------
        destDoc.Save(resultPath);

        // Verify that the output file was created.
        if (!File.Exists(resultPath))
            throw new FileNotFoundException("The extracted document was not created.", resultPath);
    }
}
