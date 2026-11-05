using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a sample document with several sections.
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);

        for (int i = 1; i <= 5; i++)
        {
            // Insert a section break before each new section (except the first).
            if (i > 1)
                builder.InsertBreak(BreakType.SectionBreakNewPage);

            builder.Writeln($"This is content of Section {i}.");
        }

        // Split the document by sections, saving even‑indexed parts as DOCX and odd‑indexed parts as PDF.
        for (int i = 0; i < sourceDoc.Sections.Count; i++)
        {
            Section originalSection = sourceDoc.Sections[i];

            // Create a new document that will contain only this section.
            Document partDoc = new Document();
            partDoc.RemoveAllChildren(); // Remove the default empty section.

            // Import the section from the source document into the new document.
            Section importedSection = (Section)partDoc.ImportNode(originalSection, true);
            partDoc.AppendChild(importedSection);

            // Determine file name and format.
            bool isEven = i % 2 == 0; // 0‑based index: even → DOCX, odd → PDF.
            string extension = isEven ? ".docx" : ".pdf";
            string fileName = $"Part_{i + 1}{extension}";

            // Save using the appropriate format.
            if (isEven)
                partDoc.Save(fileName, SaveFormat.Docx);
            else
                partDoc.Save(fileName, SaveFormat.Pdf);
        }

        // Verify that the expected files were created.
        for (int i = 0; i < 5; i++)
        {
            string extension = (i % 2 == 0) ? ".docx" : ".pdf";
            string expectedFile = $"Part_{i + 1}{extension}";
            if (!File.Exists(expectedFile))
                throw new Exception($"Expected file not found: {expectedFile}");
        }

        Console.WriteLine("Document split completed successfully.");
    }
}
