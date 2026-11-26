using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Prepare a temporary working folder.
        string workFolder = Path.Combine(Path.GetTempPath(), "AsposeJoinExample");
        Directory.CreateDirectory(workFolder);

        // Create the document that will be inserted for each data row.
        Document insertDoc = new Document();
        DocumentBuilder insertBuilder = new DocumentBuilder(insertDoc);
        insertBuilder.Writeln("=== Inserted Content ===");
        insertBuilder.Writeln("This paragraph is inserted for each record.");

        // Create the main template document with a merge field.
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);
        builder.Writeln("=== Mail Merge Report ===");
        builder.Writeln("Name: <<Name>>");
        builder.Writeln("Below is the inserted document for this record:");

        // Sample data rows.
        var data = new[]
        {
            new Person { Name = "Alice" },
            new Person { Name = "Bob" },
            new Person { Name = "Charlie" }
        };

        Document? masterDoc = null;

        foreach (var person in data)
        {
            // Clone the template for the current record.
            Document currentDoc = (Document)templateDoc.Clone(true);
            // Perform mail merge for the current record.
            currentDoc.MailMerge.Execute(new[] { "Name" }, new object[] { person.Name });

            // Append the insert document to the current document.
            currentDoc.AppendDocument(insertDoc, ImportFormatMode.KeepSourceFormatting);

            // Combine with the master document.
            if (masterDoc == null)
            {
                masterDoc = currentDoc;
            }
            else
            {
                masterDoc.AppendDocument(currentDoc, ImportFormatMode.KeepSourceFormatting);
            }
        }

        if (masterDoc == null)
        {
            throw new InvalidOperationException("No documents were generated.");
        }

        // Save the combined result as PDF.
        string pdfPath = Path.Combine(workFolder, "CombinedReport.pdf");
        masterDoc.Save(pdfPath, SaveFormat.Pdf);

        // Validate that the PDF was created.
        if (!File.Exists(pdfPath))
        {
            throw new InvalidOperationException("The combined PDF file was not created.");
        }
    }

    private class Person
    {
        public string Name { get; set; } = string.Empty;
    }
}
