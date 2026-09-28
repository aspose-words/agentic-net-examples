using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Tables;   // Required for Table class

public class Program
{
    public static void Main()
    {
        // Create the destination document with a table that contains a bookmark.
        Document destDoc = new Document();
        DocumentBuilder destBuilder = new DocumentBuilder(destDoc);

        // Start a table and add a single cell.
        Table table = destBuilder.StartTable();
        destBuilder.InsertCell();

        // Insert a bookmark named "InsertHere" inside the cell.
        destBuilder.StartBookmark("InsertHere");
        destBuilder.Writeln("Placeholder text before insertion.");
        destBuilder.EndBookmark("InsertHere");

        // Close the cell, row, and table.
        destBuilder.EndRow();
        destBuilder.EndTable();

        // Save the destination document as DOCX (optional, for inspection).
        const string destPath = "Destination.docx";
        destDoc.Save(destPath, SaveFormat.Docx);

        // Create the source document that will be inserted.
        Document srcDoc = new Document();
        DocumentBuilder srcBuilder = new DocumentBuilder(srcDoc);
        srcBuilder.Writeln("This is the content from the source DOCX document.");
        srcBuilder.Writeln("It will be inserted at the bookmark location.");
        const string srcPath = "Source.docx";
        srcDoc.Save(srcPath, SaveFormat.Docx);

        // Load the source document (demonstrating loading from file).
        Document sourceToInsert = new Document(srcPath);

        // Move the builder to the bookmark in the destination document.
        destBuilder.MoveToBookmark("InsertHere");

        // Insert the source document at the bookmark, preserving its formatting.
        destBuilder.InsertDocument(sourceToInsert, ImportFormatMode.KeepSourceFormatting);

        // Save the merged document as ODT, preserving the table structure.
        const string resultPath = "Result.odt";
        destDoc.Save(resultPath, SaveFormat.Odt);

        // Verify that the output file was created.
        if (!File.Exists(resultPath))
        {
            throw new InvalidOperationException($"Failed to create the output file: {resultPath}");
        }

        // Optional validation: ensure the inserted text is present in the resulting ODT.
        Document validationDoc = new Document(resultPath);
        bool containsInsertedText = validationDoc.GetText()
            .Contains("This is the content from the source DOCX document.");
        if (!containsInsertedText)
        {
            throw new InvalidOperationException("The inserted content was not found in the resulting ODT document.");
        }

        // Program completed successfully.
    }
}
