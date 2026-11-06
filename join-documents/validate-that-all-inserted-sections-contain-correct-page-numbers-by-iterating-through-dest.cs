using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Layout;

public class Program
{
    public static void Main()
    {
        // Prepare a temporary folder for the sample files.
        string tempFolder = Path.Combine(Path.GetTempPath(), "AsposeJoinDemo");
        Directory.CreateDirectory(tempFolder);

        // ---------- Create first source document ----------
        Document srcDoc1 = new Document();
        DocumentBuilder builder1 = new DocumentBuilder(srcDoc1);
        builder1.Writeln("Document 1 - Page 1");
        builder1.InsertBreak(BreakType.PageBreak);
        builder1.Writeln("Document 1 - Page 2");
        string srcPath1 = Path.Combine(tempFolder, "Source1.docx");
        srcDoc1.Save(srcPath1, SaveFormat.Docx);

        // ---------- Create second source document ----------
        Document srcDoc2 = new Document();
        DocumentBuilder builder2 = new DocumentBuilder(srcDoc2);
        builder2.Writeln("Document 2 - Page 1");
        builder2.InsertBreak(BreakType.PageBreak);
        builder2.Writeln("Document 2 - Page 2");
        string srcPath2 = Path.Combine(tempFolder, "Source2.docx");
        srcDoc2.Save(srcPath2, SaveFormat.Docx);

        // Load the source documents (simulating separate files).
        Document source1 = new Document(srcPath1);
        Document source2 = new Document(srcPath2);

        // Use the first document as the destination.
        Document destination = source1;

        // Append the second document, preserving its original formatting.
        destination.AppendDocument(source2, ImportFormatMode.KeepSourceFormatting);

        // Ensure layout information (page numbers) is up‑to‑date.
        destination.UpdatePageLayout();

        // Use LayoutCollector to retrieve the start page of each section.
        LayoutCollector collector = new LayoutCollector(destination);

        int previousStartPage = 0;
        int sectionIndex = 0;
        foreach (Section section in destination.Sections)
        {
            int startPage = collector.GetStartPageIndex(section);

            // Validate that sections appear in non‑decreasing page order.
            if (sectionIndex > 0 && startPage < previousStartPage)
            {
                throw new InvalidOperationException(
                    $"Section {sectionIndex} starts on page {startPage}, which is before the previous section's start page {previousStartPage}.");
            }

            previousStartPage = startPage;
            sectionIndex++;
        }

        // Save the merged document.
        string mergedPath = Path.Combine(tempFolder, "Merged.docx");
        destination.Save(mergedPath, SaveFormat.Docx);

        // Verify that the merged file was created.
        if (!File.Exists(mergedPath))
        {
            throw new FileNotFoundException("Merged document was not created.", mergedPath);
        }

        // Indicate successful completion.
        Console.WriteLine("Document merged and validated successfully.");
    }
}
