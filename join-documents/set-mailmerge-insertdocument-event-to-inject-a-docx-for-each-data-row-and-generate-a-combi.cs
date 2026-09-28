using System;
using System.Data;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Prepare a temporary working folder.
        string workFolder = Path.Combine(Path.GetTempPath(), "AsposeJoinDemo");
        Directory.CreateDirectory(workFolder);

        // Paths for the sample documents.
        string templatePath = Path.Combine(workFolder, "Template.docx");
        string insertDocPath = Path.Combine(workFolder, "Insert.docx");
        string resultPdfPath = Path.Combine(workFolder, "Combined.pdf");

        // -----------------------------------------------------------------
        // Create the main template document (Header + Footer).
        // -----------------------------------------------------------------
        var templateBuilder = new DocumentBuilder();
        templateBuilder.Writeln("Header of the main document.");
        templateBuilder.Writeln("Footer of the main document.");
        templateBuilder.Document.Save(templatePath, SaveFormat.Docx);

        // -----------------------------------------------------------------
        // Create the document that will be inserted for each data row.
        // -----------------------------------------------------------------
        var insertBuilder = new DocumentBuilder();
        insertBuilder.Writeln("=== Inserted Document Content ===");
        insertBuilder.Writeln("This content comes from the inserted document.");
        insertBuilder.Document.Save(insertDocPath, SaveFormat.Docx);

        // -----------------------------------------------------------------
        // Load the documents.
        // -----------------------------------------------------------------
        Document mainDoc = new Document(templatePath);
        Document insertDoc = new Document(insertDocPath);

        // -----------------------------------------------------------------
        // Prepare a data source with two rows; each row will trigger an insertion.
        // -----------------------------------------------------------------
        DataTable data = new DataTable("Data");
        data.Columns.Add("InsertDoc", typeof(string));
        data.Rows.Add("Row1");
        data.Rows.Add("Row2");

        // -----------------------------------------------------------------
        // Append the insert document for each data row.
        // -----------------------------------------------------------------
        DocumentBuilder builder = new DocumentBuilder(mainDoc);
        bool first = true;
        foreach (DataRow row in data.Rows)
        {
            // Optionally add a page break before each inserted document except the first.
            if (!first)
            {
                builder.InsertBreak(BreakType.PageBreak);
            }

            // Insert the whole document at the current builder position.
            builder.InsertDocument(insertDoc, ImportFormatMode.KeepSourceFormatting);
            first = false;
        }

        // -----------------------------------------------------------------
        // Save the combined result as PDF.
        // -----------------------------------------------------------------
        mainDoc.Save(resultPdfPath, SaveFormat.Pdf);

        // -----------------------------------------------------------------
        // Validate that the PDF was created.
        // -----------------------------------------------------------------
        if (!File.Exists(resultPdfPath))
        {
            throw new InvalidOperationException("The combined PDF file was not created.");
        }

        // Optional: indicate success.
        Console.WriteLine($"Combined PDF created at: {resultPdfPath}");
    }
}
