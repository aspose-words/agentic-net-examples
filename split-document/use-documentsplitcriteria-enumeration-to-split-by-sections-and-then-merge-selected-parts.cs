using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class SplitAndMergeExample
{
    public static void Main()
    {
        // -------------------------------------------------
        // 1. Create a sample document with three sections.
        // -------------------------------------------------
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);

        // Section 1
        builder.Writeln("Section 1 - Introduction");
        builder.InsertBreak(BreakType.SectionBreakNewPage);

        // Section 2
        builder.Writeln("Section 2 - Body");
        builder.InsertBreak(BreakType.SectionBreakNewPage);

        // Section 3
        builder.Writeln("Section 3 - Conclusion");

        // -------------------------------------------------
        // 2. Split the document by sections and save each part as HTML.
        // -------------------------------------------------
        string baseName = "SplitDocument";
        string[] splitFiles = new string[sourceDoc.Sections.Count];

        for (int i = 0; i < sourceDoc.Sections.Count; i++)
        {
            // Create a new document that will contain only the current section.
            Document partDoc = new Document();
            // Import the section from the source document.
            Section srcSection = sourceDoc.Sections[i];
            Node imported = partDoc.ImportNode(srcSection, true, ImportFormatMode.KeepSourceFormatting);
            partDoc.AppendChild(imported);

            // Prepare the file name.
            string fileName = $"{baseName}_section{i + 1}.html";
            splitFiles[i] = fileName;

            // Save as HTML. ExportHeadersFootersMode.PerSection ensures headers/footers are kept.
            HtmlSaveOptions saveOptions = new HtmlSaveOptions
            {
                ExportHeadersFootersMode = ExportHeadersFootersMode.PerSection
            };
            partDoc.Save(fileName, saveOptions);
        }

        // -------------------------------------------------
        // 3. Verify that split files were created.
        // -------------------------------------------------
        foreach (string file in splitFiles)
        {
            if (!File.Exists(file))
                throw new Exception($"Expected split file not found: {file}");
        }

        // -------------------------------------------------
        // 4. Merge selected sections (Section 1 and Section 3) into a new document.
        // -------------------------------------------------
        Document mergedDoc = new Document();
        // Remove the initially empty section that a new Document contains.
        mergedDoc.RemoveAllChildren();

        // Load and import Section 1.
        Document part1 = new Document(splitFiles[0]);
        foreach (Section sec in part1.Sections)
        {
            Node imported = mergedDoc.ImportNode(sec, true, ImportFormatMode.KeepSourceFormatting);
            mergedDoc.AppendChild(imported);
        }

        // Load and import Section 3.
        Document part3 = new Document(splitFiles[2]);
        foreach (Section sec in part3.Sections)
        {
            Node imported = mergedDoc.ImportNode(sec, true, ImportFormatMode.KeepSourceFormatting);
            mergedDoc.AppendChild(imported);
        }

        // Save the merged document.
        string mergedPath = "MergedDocument.docx";
        mergedDoc.Save(mergedPath);

        // Verify merged document exists.
        if (!File.Exists(mergedPath))
            throw new Exception("Merged document was not created.");

        // End of example.
    }
}
