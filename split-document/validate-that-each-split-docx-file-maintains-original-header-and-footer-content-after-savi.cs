using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a sample document with two sections, each having the same header and footer.
        Document sourceDoc = new Document();

        // Remove the default empty section created by the Document constructor.
        sourceDoc.Sections.Clear();

        for (int i = 0; i < 2; i++)
        {
            // Create a new section and add it to the document.
            Section section = new Section(sourceDoc);
            sourceDoc.Sections.Add(section);

            // Every section must contain a Body node before we can add paragraphs.
            Body body = new Body(sourceDoc);
            section.AppendChild(body);

            // Add header with marker text.
            HeaderFooter header = new HeaderFooter(sourceDoc, HeaderFooterType.HeaderPrimary);
            Paragraph headerPara = new Paragraph(sourceDoc);
            headerPara.AppendChild(new Run(sourceDoc, "Header Text"));
            header.AppendChild(headerPara);
            section.HeadersFooters.Add(header);

            // Add footer with marker text.
            HeaderFooter footer = new HeaderFooter(sourceDoc, HeaderFooterType.FooterPrimary);
            Paragraph footerPara = new Paragraph(sourceDoc);
            footerPara.AppendChild(new Run(sourceDoc, "Footer Text"));
            footer.AppendChild(footerPara);
            section.HeadersFooters.Add(footer);

            // Add body content.
            Paragraph bodyPara = new Paragraph(sourceDoc);
            bodyPara.AppendChild(new Run(sourceDoc, $"Section {i + 1} content."));
            body.AppendChild(bodyPara);
        }

        // Split the document by sections, preserving headers and footers.
        string outputFolder = Path.Combine(Directory.GetCurrentDirectory(), "SplitOutputs");
        Directory.CreateDirectory(outputFolder);

        for (int i = 0; i < sourceDoc.Sections.Count; i++)
        {
            Section originalSection = sourceDoc.Sections[i];

            // Create a new empty document for the split part.
            Document splitDoc = new Document();

            // Import the section into the new document.
            NodeImporter importer = new NodeImporter(sourceDoc, splitDoc, ImportFormatMode.KeepSourceFormatting);
            Section importedSection = (Section)importer.ImportNode(originalSection, true);

            // Remove the default empty section and add the imported one.
            splitDoc.Sections.Clear();
            splitDoc.Sections.Add(importedSection);

            // Save the split document.
            string splitPath = Path.Combine(outputFolder, $"Section_{i + 1}.docx");
            splitDoc.Save(splitPath);
        }

        // Validate that each split file exists and retains the original header/footer text.
        for (int i = 0; i < sourceDoc.Sections.Count; i++)
        {
            string splitPath = Path.Combine(outputFolder, $"Section_{i + 1}.docx");

            if (!File.Exists(splitPath))
                throw new Exception($"Expected split file not found: {splitPath}");

            Document loaded = new Document(splitPath);
            HeaderFooter loadedHeader = loaded.FirstSection.HeadersFooters[HeaderFooterType.HeaderPrimary];
            HeaderFooter loadedFooter = loaded.FirstSection.HeadersFooters[HeaderFooterType.FooterPrimary];

            if (loadedHeader == null || !loadedHeader.GetText().Contains("Header Text"))
                throw new Exception($"Header text missing or incorrect in file: {splitPath}");

            if (loadedFooter == null || !loadedFooter.GetText().Contains("Footer Text"))
                throw new Exception($"Footer text missing or incorrect in file: {splitPath}");
        }

        // Indicate successful completion.
        Console.WriteLine("All split documents created and validated successfully.");
    }
}
