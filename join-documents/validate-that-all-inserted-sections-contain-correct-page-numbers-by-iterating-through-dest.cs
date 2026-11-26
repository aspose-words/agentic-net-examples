using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Layout;

public class Program
{
    public static void Main()
    {
        // Prepare file paths
        string folder = Directory.GetCurrentDirectory();
        string sourcePath1 = Path.Combine(folder, "Source1.docx");
        string sourcePath2 = Path.Combine(folder, "Source2.docx");
        string mergedPath = Path.Combine(folder, "Merged.docx");

        // Create first source document with two pages and PAGE fields
        var doc1 = new Document();
        var builder1 = new DocumentBuilder(doc1);
        builder1.Writeln("Document 1 - Page 1");
        builder1.InsertField("PAGE", "");
        builder1.InsertBreak(BreakType.PageBreak);
        builder1.Writeln("Document 1 - Page 2");
        builder1.InsertField("PAGE", "");
        doc1.Save(sourcePath1, SaveFormat.Docx);

        // Create second source document with two pages and PAGE fields
        var doc2 = new Document();
        var builder2 = new DocumentBuilder(doc2);
        builder2.Writeln("Document 2 - Page 1");
        builder2.InsertField("PAGE", "");
        builder2.InsertBreak(BreakType.PageBreak);
        builder2.Writeln("Document 2 - Page 2");
        builder2.InsertField("PAGE", "");
        doc2.Save(sourcePath2, SaveFormat.Docx);

        // Load source documents
        var sourceDoc1 = new Document(sourcePath1);
        var sourceDoc2 = new Document(sourcePath2);

        // Create destination document and append source documents
        var destDoc = new Document();
        destDoc.AppendDocument(sourceDoc1, ImportFormatMode.KeepSourceFormatting);
        destDoc.AppendDocument(sourceDoc2, ImportFormatMode.KeepSourceFormatting);

        // Update fields and layout to ensure PAGE fields are calculated
        destDoc.UpdateFields();
        destDoc.UpdatePageLayout();

        // Validate that each inserted section starts on the correct page number
        var collector = new LayoutCollector(destDoc);
        int previousStart = 0;
        int previousPageCount = 0;

        for (int i = 0; i < destDoc.Sections.Count; i++)
        {
            Section section = destDoc.Sections[i];
            Paragraph firstParagraph = section.Body.FirstParagraph;
            if (firstParagraph == null)
                throw new InvalidOperationException($"Section {i + 1} has no paragraphs.");

            int startPage = collector.GetStartPageIndex(firstParagraph);
            int expectedStart = (i == 0) ? 1 : previousStart + previousPageCount;

            if (startPage != expectedStart)
                throw new Exception($"Section {i + 1} starts on page {startPage}, but expected {expectedStart}.");

            // Determine page count of this section
            int pageCount;
            if (i < destDoc.Sections.Count - 1)
            {
                Paragraph nextFirstParagraph = destDoc.Sections[i + 1].Body.FirstParagraph;
                int nextStart = collector.GetStartPageIndex(nextFirstParagraph);
                pageCount = nextStart - startPage;
            }
            else
            {
                pageCount = destDoc.PageCount - startPage + 1;
            }

            previousStart = startPage;
            previousPageCount = pageCount;
        }

        // Save merged document
        destDoc.Save(mergedPath, SaveFormat.Docx);

        // Verify merged file exists
        if (!File.Exists(mergedPath))
            throw new FileNotFoundException("Merged document was not created.", mergedPath);

        // Clean up temporary source files
        File.Delete(sourcePath1);
        File.Delete(sourcePath2);
    }
}
