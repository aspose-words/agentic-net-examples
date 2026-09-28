using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a sample document with four sections, each on a new page.
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);

        for (int i = 1; i <= 4; i++)
        {
            builder.Writeln($"This is content of Section {i}.");
            if (i < 4)
                builder.InsertBreak(BreakType.SectionBreakNewPage);
        }

        // Split the document by sections.
        // Even‑indexed parts (0‑based) will be saved as DOCX, odd‑indexed parts as PDF.
        for (int i = 0; i < sourceDoc.Sections.Count; i++)
        {
            Section srcSection = sourceDoc.Sections[i];

            // Create a new empty document and import the current section.
            Document partDoc = new Document();
            partDoc.RemoveAllChildren(); // Remove the default empty section.

            Node importedSection = partDoc.ImportNode(srcSection, true);
            partDoc.AppendChild(importedSection);

            // Determine format and file name.
            bool isEven = (i % 2 == 0);
            SaveFormat format = isEven ? SaveFormat.Docx : SaveFormat.Pdf;
            string extension = isEven ? ".docx" : ".pdf";
            string fileName = $"Part_{i + 1}{extension}";

            // Save the split part.
            partDoc.Save(fileName, format);
        }

        // Validate that the expected files were created.
        for (int i = 1; i <= 4; i++)
        {
            string expectedFile = $"Part_{i}" + (i % 2 == 1 ? ".docx" : ".pdf");
            if (!File.Exists(expectedFile))
                throw new FileNotFoundException($"Expected split file not found: {expectedFile}");
        }

        Console.WriteLine("Document split completed successfully.");
    }
}
