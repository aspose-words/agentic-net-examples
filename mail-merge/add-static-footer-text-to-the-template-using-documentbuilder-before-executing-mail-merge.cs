using System;
using Aspose.Words;
using Aspose.Words.BuildingBlocks;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a static footer to the document.
        // Move the builder to the primary footer of the first section.
        builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);
        builder.Write("This is a static footer added before mail merge.");
        // Return the builder to the main body of the document.
        builder.MoveToDocumentEnd();

        // Build the mail‑merge template in the main body.
        builder.Writeln(); // Ensure we are on a new line after the footer.
        builder.Write("Dear ");
        builder.InsertField("MERGEFIELD FirstName", "<FirstName>");
        builder.Write(" ");
        builder.InsertField("MERGEFIELD LastName", "<LastName>");
        builder.Writeln(":");
        builder.InsertField("MERGEFIELD Message", "<Message>");

        // Prepare simple mail‑merge data.
        string[] fieldNames = { "FirstName", "LastName", "Message" };
        object[] fieldValues = { "John", "Doe", "Hello! This document contains a static footer." };

        // Execute the mail merge.
        doc.MailMerge.Execute(fieldNames, fieldValues);

        // Save the resulting document.
        doc.Save("Result.docx");
    }
}
