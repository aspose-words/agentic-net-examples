using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Layout;
using Aspose.Words.Fields; // Needed for FieldType enum

public class JoinDocumentsValidatePageNumbers
{
    public static void Main()
    {
        // Folder for temporary files
        string workDir = Path.Combine(Directory.GetCurrentDirectory(), "Work");
        Directory.CreateDirectory(workDir);

        // Create first source document with a PAGE field
        Document srcDoc1 = new Document();
        DocumentBuilder builder1 = new DocumentBuilder(srcDoc1);
        builder1.Writeln("Source Document 1");
        builder1.InsertField(FieldType.FieldPage, true); // page number field
        string srcPath1 = Path.Combine(workDir, "Source1.docx");
        srcDoc1.Save(srcPath1, SaveFormat.Docx);

        // Create second source document with a PAGE field
        Document srcDoc2 = new Document();
        DocumentBuilder builder2 = new DocumentBuilder(srcDoc2);
        builder2.Writeln("Source Document 2");
        builder2.InsertField(FieldType.FieldPage, true);
        string srcPath2 = Path.Combine(workDir, "Source2.docx");
        srcDoc2.Save(srcPath2, SaveFormat.Docx);

        // Destination document – start with some content
        Document dstDoc = new Document();
        DocumentBuilder dstBuilder = new DocumentBuilder(dstDoc);
        dstBuilder.Writeln("Destination Document Start");
        dstBuilder.InsertField(FieldType.FieldPage, true);

        // Append the two source documents, each starting on a new page
        Document src1 = new Document(srcPath1);
        Document src2 = new Document(srcPath2);
        dstDoc.AppendDocument(src1, ImportFormatMode.KeepSourceFormatting);
        dstDoc.AppendDocument(src2, ImportFormatMode.KeepSourceFormatting);

        // Update fields so PAGE fields contain actual numbers
        dstDoc.UpdateFields();

        // Save the merged document
        string mergedPath = Path.Combine(workDir, "Merged.docx");
        dstDoc.Save(mergedPath, SaveFormat.Docx);

        // Verify that the merged file exists
        if (!File.Exists(mergedPath))
            throw new InvalidOperationException("Merged document was not created.");

        // Use LayoutCollector to map nodes to page numbers
        LayoutCollector collector = new LayoutCollector(dstDoc);

        // Validate that each section starts on the expected page number
        for (int i = 0; i < dstDoc.Sections.Count; i++)
        {
            Section section = dstDoc.Sections[i];
            // Ensure the section has at least one paragraph
            if (section.Body.FirstParagraph == null)
                throw new InvalidOperationException($"Section {i + 1} has no paragraphs.");

            int actualPage = collector.GetStartPageIndex(section.Body.FirstParagraph);
            int expectedPage = i + 1; // sections are 1‑based pages because each starts on a new page

            if (actualPage != expectedPage)
                throw new InvalidOperationException(
                    $"Section {i + 1} starts on page {actualPage}, but expected page {expectedPage}.");
        }

        // If we reach this point, validation succeeded
        Console.WriteLine("All sections contain correct page numbers. Merged document saved to:");
        Console.WriteLine(mergedPath);
    }
}
