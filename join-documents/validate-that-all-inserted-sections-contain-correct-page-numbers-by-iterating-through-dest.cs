using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.BuildingBlocks;
using Aspose.Words.Fields;   // Needed for FieldType

public class Program
{
    public static void Main()
    {
        // Prepare a folder for temporary files.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // Create two source documents with simple content.
        string srcPath1 = Path.Combine(artifactsDir, "Source1.docx");
        string srcPath2 = Path.Combine(artifactsDir, "Source2.docx");

        CreateSourceDocument(srcPath1, "First source document.");
        CreateSourceDocument(srcPath2, "Second source document.");

        // Load the source documents.
        Document srcDoc1 = new Document(srcPath1);
        Document srcDoc2 = new Document(srcPath2);

        // Create the destination document (initially empty).
        Document dstDoc = new Document();

        // Append the source documents to the destination.
        dstDoc.AppendDocument(srcDoc1, ImportFormatMode.KeepSourceFormatting);
        dstDoc.AppendDocument(srcDoc2, ImportFormatMode.KeepSourceFormatting);

        // Ensure layout is up‑to‑date before inserting page numbers.
        dstDoc.UpdatePageLayout();

        // Insert a PAGE field into the primary footer of each section.
        for (int i = 0; i < dstDoc.Sections.Count; i++)
        {
            DocumentBuilder builder = new DocumentBuilder(dstDoc);
            builder.MoveToSection(i);
            builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);
            // Insert a PAGE field that will display the current page number.
            builder.InsertField(FieldType.FieldPage, true);
        }

        // Update fields so that the PAGE fields contain their results.
        dstDoc.UpdateFields();

        // Validate that each footer's PAGE field shows the expected page number.
        int expectedPage = 1;
        foreach (Section section in dstDoc.Sections)
        {
            HeaderFooter footer = section.HeadersFooters[HeaderFooterType.FooterPrimary];
            if (footer != null)
            {
                // Find the first field in the footer (there should be exactly one PAGE field).
                var pageField = footer.Range.Fields.FirstOrDefault();
                if (pageField == null)
                    throw new InvalidOperationException($"Section {expectedPage} does not contain a PAGE field.");

                string fieldResult = pageField.Result?.Trim();
                if (fieldResult != expectedPage.ToString())
                    throw new InvalidOperationException(
                        $"Page number mismatch in section {expectedPage}: expected {expectedPage}, got {fieldResult}.");
            }
            expectedPage++;
        }

        // Save the merged document.
        string mergedPath = Path.Combine(artifactsDir, "Merged.docx");
        dstDoc.Save(mergedPath);

        // Verify that the file was created.
        if (!File.Exists(mergedPath))
            throw new FileNotFoundException("Merged document was not saved.", mergedPath);
    }

    // Helper method to create a simple one‑section document with given text.
    private static void CreateSourceDocument(string filePath, string text)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln(text);
        doc.Save(filePath);
    }
}
