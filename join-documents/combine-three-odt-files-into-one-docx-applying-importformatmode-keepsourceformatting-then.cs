using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Define file names
        string doc1Path = "doc1.odt";
        string doc2Path = "doc2.odt";
        string doc3Path = "doc3.odt";
        string mergedPath = "merged.docx";

        // Create first ODT document
        var doc1 = new Document();
        var builder1 = new DocumentBuilder(doc1);
        builder1.Writeln("Content of Document 1");
        doc1.Save(doc1Path, SaveFormat.Odt);

        // Create second ODT document
        var doc2 = new Document();
        var builder2 = new DocumentBuilder(doc2);
        builder2.Writeln("Content of Document 2");
        doc2.Save(doc2Path, SaveFormat.Odt);

        // Create third ODT document
        var doc3 = new Document();
        var builder3 = new DocumentBuilder(doc3);
        builder3.Writeln("Content of Document 3");
        doc3.Save(doc3Path, SaveFormat.Odt);

        // Load the first document as the destination
        var mergedDoc = new Document(doc1Path);

        // Append the second and third documents, preserving source formatting
        mergedDoc.AppendDocument(new Document(doc2Path), ImportFormatMode.KeepSourceFormatting);
        mergedDoc.AppendDocument(new Document(doc3Path), ImportFormatMode.KeepSourceFormatting);

        // Save the merged document as DOCX
        mergedDoc.Save(mergedPath, SaveFormat.Docx);

        // Validation: ensure the merged file exists
        if (!File.Exists(mergedPath))
            throw new Exception("Merged DOCX file was not created.");

        // Validation: ensure content from all source documents is present
        var resultDoc = new Document(mergedPath);
        string allText = resultDoc.GetText();

        if (!allText.Contains("Content of Document 1") ||
            !allText.Contains("Content of Document 2") ||
            !allText.Contains("Content of Document 3"))
        {
            throw new Exception("Merged document does not contain content from all source files.");
        }

        // Cleanup: optional removal of temporary files (comment out if inspection needed)
        // File.Delete(doc1Path);
        // File.Delete(doc2Path);
        // File.Delete(doc3Path);
    }
}
