using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a sample document with two sections.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Section 1 - First paragraph.");
        builder.InsertBreak(BreakType.SectionBreakNewPage);
        builder.Writeln("Section 2 - First paragraph.");

        // Save the source document (optional, just for reference).
        string sourcePath = "Sample.docx";
        doc.Save(sourcePath);

        // Prepare output folder.
        string outputFolder = "SplitOutput";
        Directory.CreateDirectory(outputFolder);

        // Split the document by sections: create a new document for each section.
        int sectionIndex = 1;
        foreach (Section srcSection in doc.Sections)
        {
            // Create a new empty document.
            Document splitDoc = new Document();
            // Remove the default empty section that a new Document contains.
            splitDoc.RemoveAllChildren();

            // Import the source section into the new document.
            NodeImporter importer = new NodeImporter(doc, splitDoc, ImportFormatMode.KeepSourceFormatting);
            Section importedSection = (Section)importer.ImportNode(srcSection, true);
            splitDoc.AppendChild(importedSection);

            // Save each split document as HTML.
            string splitFileName = $"Split_section_{sectionIndex}.html";
            string splitPath = Path.Combine(outputFolder, splitFileName);
            splitDoc.Save(splitPath, SaveFormat.Html);
            sectionIndex++;
        }

        // Validate that at least two split files were created (one per section).
        string[] htmlFiles = Directory.GetFiles(outputFolder, "*.html");
        if (htmlFiles.Length < 2)
        {
            throw new Exception($"Expected at least 2 split HTML files, but found {htmlFiles.Length}.");
        }

        // Optionally, clean up created files (comment out if you want to inspect them).
        // foreach (var file in htmlFiles) File.Delete(file);
        // Directory.Delete(outputFolder);
    }
}
