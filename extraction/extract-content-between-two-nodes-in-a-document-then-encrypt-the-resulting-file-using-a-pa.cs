using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // -------------------------------------------------
        // 1. Create a sample source document with identifiable paragraphs.
        // -------------------------------------------------
        Document source = new Document();
        DocumentBuilder builder = new DocumentBuilder(source);
        builder.Writeln("Intro paragraph.");
        builder.Writeln("Start paragraph - include this.");
        builder.Writeln("Middle paragraph - also include.");
        builder.Writeln("End paragraph - include this.");
        builder.Writeln("Final paragraph.");
        source.Save("source.docx");

        // -------------------------------------------------
        // 2. Load the source document from disk.
        // -------------------------------------------------
        Document loaded = new Document("source.docx");

        // -------------------------------------------------
        // 3. Locate the start and end paragraphs that bound the extraction range.
        // -------------------------------------------------
        Paragraph startPara = loaded.FirstSection.Body.Paragraphs[1];
        Paragraph endPara = loaded.FirstSection.Body.Paragraphs[3];
        if (startPara == null || endPara == null)
            throw new InvalidOperationException("Boundary paragraphs not found.");

        // -------------------------------------------------
        // 4. Prepare a new empty document that will hold the extracted content.
        // -------------------------------------------------
        Document result = new Document();
        result.RemoveAllChildren();                     // Remove the default empty section.
        Section section = new Section(result);
        result.AppendChild(section);
        Body body = new Body(result);
        section.AppendChild(body);

        // -------------------------------------------------
        // 5. Import (clone) nodes from the source document into the result document.
        //    Use NodeImporter to transfer nodes between documents safely.
        // -------------------------------------------------
        NodeImporter importer = new NodeImporter(loaded, result, ImportFormatMode.KeepSourceFormatting);
        Node current = startPara;
        while (current != null)
        {
            // Import the current node (deep clone) into the destination document.
            Node importedNode = importer.ImportNode(current, true);
            body.AppendChild(importedNode);

            if (current == endPara)
                break;

            current = current.NextSibling;
        }

        // -------------------------------------------------
        // 6. Encrypt the extracted document with a password and save it.
        // -------------------------------------------------
        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions(SaveFormat.Docx)
        {
            Password = "SecretPassword123"
        };
        result.Save("extracted-encrypted.docx", saveOptions);

        // -------------------------------------------------
        // 7. Verify that the encrypted file was created.
        // -------------------------------------------------
        if (!File.Exists("extracted-encrypted.docx"))
            throw new InvalidOperationException("Encrypted output file was not created.");
    }
}
