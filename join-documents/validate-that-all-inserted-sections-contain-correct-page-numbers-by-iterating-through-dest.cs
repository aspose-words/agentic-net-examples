using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;

public class JoinDocumentsValidatePageNumbers
{
    public static void Main()
    {
        // Prepare a folder for temporary files.
        string dataDir = Path.Combine(Directory.GetCurrentDirectory(), "Data");
        Directory.CreateDirectory(dataDir);

        // Paths for the sample source documents and the merged result.
        string src1Path = Path.Combine(dataDir, "Source1.docx");
        string src2Path = Path.Combine(dataDir, "Source2.docx");
        string mergedPath = Path.Combine(dataDir, "Merged.docx");

        // ---------- Create first source document ----------
        Document srcDoc1 = new Document();
        DocumentBuilder builder1 = new DocumentBuilder(srcDoc1);
        builder1.Writeln("First document");
        // Insert a PAGE field that will display the current page number.
        builder1.InsertField(FieldType.FieldPage, true);
        srcDoc1.Save(src1Path);

        // ---------- Create second source document ----------
        Document srcDoc2 = new Document();
        DocumentBuilder builder2 = new DocumentBuilder(srcDoc2);
        builder2.Writeln("Second document");
        builder2.InsertField(FieldType.FieldPage, true);
        srcDoc2.Save(src2Path);

        // ---------- Load the first document as the destination ----------
        Document dstDoc = new Document(src1Path);

        // Append the second document while keeping its original formatting.
        dstDoc.AppendDocument(srcDoc2, ImportFormatMode.KeepSourceFormatting);

        // Update fields and layout so that PAGE fields reflect correct numbers.
        dstDoc.UpdateFields();
        dstDoc.UpdatePageLayout();

        // ---------- Validation: each PAGE field must contain the expected page number ----------
        // Expected page numbers after the merge: 1 for the first document, 2 for the second.
        int[] expectedPageNumbers = { 1, 2 };
        int index = 0;

        foreach (Field field in dstDoc.Range.Fields)
        {
            if (field.Type == FieldType.FieldPage)
            {
                // The field result is the page number as a string.
                string result = field.Result.Trim();

                if (!int.TryParse(result, out int actualPage))
                {
                    throw new InvalidOperationException($"Unable to parse page number from field result '{result}'.");
                }

                if (index >= expectedPageNumbers.Length)
                {
                    throw new InvalidOperationException("More PAGE fields found than expected.");
                }

                if (actualPage != expectedPageNumbers[index])
                {
                    throw new InvalidOperationException(
                        $"Page number mismatch in section {index + 1}: expected {expectedPageNumbers[index]}, got {actualPage}.");
                }

                index++;
            }
        }

        if (index != expectedPageNumbers.Length)
        {
            throw new InvalidOperationException(
                $"Expected {expectedPageNumbers.Length} PAGE fields, but found {index}.");
        }

        // Save the merged document.
        dstDoc.Save(mergedPath);

        // Verify that the merged file was created.
        if (!File.Exists(mergedPath))
        {
            throw new FileNotFoundException("Merged document was not saved correctly.", mergedPath);
        }

        // Program completed successfully.
        Console.WriteLine("Document merged and page numbers validated successfully.");
    }
}
