using System;
using Aspose.Words;

class OdtMerger
{
    static void Main()
    {
        // Create sample ODT files if they do not exist.
        string[] sourceFiles = { "Document1.odt", "Document2.odt", "Document3.odt" };
        for (int i = 0; i < sourceFiles.Length; i++)
        {
            if (!System.IO.File.Exists(sourceFiles[i]))
            {
                var doc = new Document();
                var builder = new DocumentBuilder(doc);
                builder.Writeln($"This is content of {sourceFiles[i]}");
                doc.Save(sourceFiles[i], SaveFormat.Odt);
            }
        }

        // Create a blank destination document.
        Document destination = new Document();

        // Load each ODT file and append it to the destination document,
        // preserving the original formatting of each source.
        foreach (string filePath in sourceFiles)
        {
            Document source = new Document(filePath);
            destination.AppendDocument(source, ImportFormatMode.KeepSourceFormatting);
        }

        // Save the combined document as DOCX.
        destination.Save("Combined.docx", SaveFormat.Docx);
    }
}
