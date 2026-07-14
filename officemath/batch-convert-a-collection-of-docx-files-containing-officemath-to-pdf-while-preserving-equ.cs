using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Math;
using Aspose.Words.Saving;

public class BatchOfficeMathToPdf
{
    public static void Main()
    {
        // Prepare folders.
        string baseDir = Directory.GetCurrentDirectory();
        string inputFolder = Path.Combine(baseDir, "InputDocs");
        string outputFolder = Path.Combine(baseDir, "OutputPdfs");

        Directory.CreateDirectory(inputFolder);
        Directory.CreateDirectory(outputFolder);

        // Create sample DOCX files with simple OfficeMath equations.
        CreateSampleDocx(Path.Combine(inputFolder, "Equation1.docx"), @"\f(1,2)");   // fraction 1/2
        CreateSampleDocx(Path.Combine(inputFolder, "Equation2.docx"), @"\r(3,x)");   // cube root of x
        CreateSampleDocx(Path.Combine(inputFolder, "Equation3.docx"), @"\s \up2(a)"); // superscript a²

        // Batch convert each DOCX to PDF.
        foreach (string docxPath in Directory.GetFiles(inputFolder, "*.docx"))
        {
            Document doc = new Document(docxPath);
            string pdfPath = Path.Combine(outputFolder, Path.GetFileNameWithoutExtension(docxPath) + ".pdf");
            doc.Save(pdfPath, SaveFormat.Pdf);

            if (!File.Exists(pdfPath))
                throw new InvalidOperationException($"PDF was not created: {pdfPath}");
        }

        Console.WriteLine("Batch conversion completed successfully.");
    }

    private static void CreateSampleDocx(string filePath, string eqArgument)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a descriptive paragraph.
        builder.Writeln("Sample equation:");

        // Insert the EQ field and convert it to a real OfficeMath node.
        InsertEquation(builder, eqArgument);

        // Add a blank line after the equation.
        builder.Writeln();

        // Save the document as DOCX.
        doc.Save(filePath, SaveFormat.Docx);
    }

    private static void InsertEquation(DocumentBuilder builder, string eqArgument)
    {
        // Insert an EQ field.
        FieldEQ field = (FieldEQ)builder.InsertField(FieldType.FieldEquation, true);

        // Write the EQ argument into the field separator.
        builder.MoveTo(field.Separator);
        builder.Write(eqArgument);

        // Update the field so that the EQ code is recognized.
        field.Update();

        // Return to the field start.
        builder.MoveTo(field.Start);

        // Convert the field to OfficeMath.
        OfficeMath officeMath = field.AsOfficeMath();

        if (officeMath == null)
            throw new InvalidOperationException("Failed to convert EQ field to OfficeMath.");

        // Insert the OfficeMath node before the field start.
        field.Start.ParentNode.InsertBefore(officeMath, field.Start);

        // Remove the original field.
        field.Remove();
    }
}
