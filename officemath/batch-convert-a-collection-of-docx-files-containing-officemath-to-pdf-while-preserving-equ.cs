using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Math;

public class Program
{
    public static void Main()
    {
        // Define folders for input DOCX files and output PDF files.
        string inputFolder = Path.Combine(Directory.GetCurrentDirectory(), "InputDocs");
        string outputFolder = Path.Combine(Directory.GetCurrentDirectory(), "OutputPdfs");

        Directory.CreateDirectory(inputFolder);
        Directory.CreateDirectory(outputFolder);

        // Create a few sample DOCX documents that contain OfficeMath equations.
        for (int i = 1; i <= 3; i++)
        {
            string docPath = Path.Combine(inputFolder, $"Sample{i}.docx");
            CreateSampleDocumentWithEquation(docPath, i);
        }

        // Batch convert each DOCX file to PDF while preserving equation fidelity.
        foreach (string docxFile in Directory.GetFiles(inputFolder, "*.docx"))
        {
            // Load the DOCX document.
            Document doc = new Document(docxFile);

            // Save as PDF in the output folder.
            string pdfPath = Path.Combine(outputFolder, Path.GetFileNameWithoutExtension(docxFile) + ".pdf");
            doc.Save(pdfPath, SaveFormat.Pdf);

            // Validate that the PDF was created.
            if (!File.Exists(pdfPath))
                throw new InvalidOperationException($"PDF conversion failed for '{docxFile}'.");
        }

        // Optional: indicate completion (no interactive input required).
        Console.WriteLine("Batch conversion completed successfully.");
    }

    // Creates a DOCX file containing a simple OfficeMath equation using the deterministic EQ-field bootstrap workflow.
    private static void CreateSampleDocumentWithEquation(string filePath, int index)
    {
        // Create a blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a title paragraph.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln($"Sample Document {index}");

        // Add a normal paragraph.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("The following is a simple equation:");

        // Insert an OfficeMath equation (fraction 1/2) using the EQ field bootstrap.
        InsertOfficeMath(builder, @"\f(1,2)");

        // Add another paragraph to separate documents.
        builder.Writeln("End of document.");

        // Save the document as DOCX.
        doc.Save(filePath, SaveFormat.Docx);
    }

    // Inserts an OfficeMath equation into the document using the EQ field bootstrap pattern.
    private static void InsertOfficeMath(DocumentBuilder builder, string eqArguments)
    {
        // Insert an EQ field.
        FieldEQ field = (FieldEQ)builder.InsertField(FieldType.FieldEquation, true);

        // Write the EQ arguments into the field separator.
        builder.MoveTo(field.Separator);
        builder.Write(eqArguments);

        // Return the builder to the paragraph that contains the field.
        builder.MoveTo(field.Start.ParentNode);

        // Convert the EQ field to a real OfficeMath object.
        OfficeMath officeMath = field.AsOfficeMath();

        // If conversion succeeded, replace the field with the OfficeMath node.
        if (officeMath != null)
        {
            // Insert the OfficeMath node before the field start.
            field.Start.ParentNode.InsertBefore(officeMath, field.Start);
            // Remove the original field.
            field.Remove();
        }

        // Add a line break after the equation for readability.
        builder.Writeln();
    }
}
