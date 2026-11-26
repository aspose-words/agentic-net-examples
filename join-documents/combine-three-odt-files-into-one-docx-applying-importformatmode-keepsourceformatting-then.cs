using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Define file names for the source ODT files and the merged DOCX output.
        string[] sourceFiles = { "Doc1.odt", "Doc2.odt", "Doc3.odt" };
        string mergedFile = "Merged.docx";

        // Create three sample ODT documents with distinct content.
        for (int i = 0; i < sourceFiles.Length; i++)
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln($"This is the content of source document {i + 1}.");
            doc.Save(sourceFiles[i], SaveFormat.Odt);
        }

        // Load the first document which will serve as the destination.
        Document mergedDoc = new Document(sourceFiles[0]);

        // Append the remaining documents using KeepSourceFormatting.
        for (int i = 1; i < sourceFiles.Length; i++)
        {
            Document srcDoc = new Document(sourceFiles[i]);
            mergedDoc.AppendDocument(srcDoc, ImportFormatMode.KeepSourceFormatting);
        }

        // Save the merged document as DOCX.
        mergedDoc.Save(mergedFile, SaveFormat.Docx);

        // Validation: ensure the merged file exists.
        if (!File.Exists(mergedFile))
        {
            throw new InvalidOperationException($"Merged file '{mergedFile}' was not created.");
        }

        // Validation: ensure the merged document contains three sections (one per source).
        Document validationDoc = new Document(mergedFile);
        if (validationDoc.Sections.Count != sourceFiles.Length)
        {
            throw new InvalidOperationException($"Merged document should contain {sourceFiles.Length} sections, but contains {validationDoc.Sections.Count}.");
        }

        // Optional cleanup: delete temporary source files (comment out if you want to keep them).
        foreach (string file in sourceFiles)
        {
            if (File.Exists(file))
                File.Delete(file);
        }
    }
}
