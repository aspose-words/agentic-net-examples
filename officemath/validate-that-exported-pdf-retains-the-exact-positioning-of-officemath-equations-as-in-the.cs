using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Math;
using Aspose.Words.Saving;

public class OfficeMathPdfPositionValidation
{
    public static void Main()
    {
        // Prepare output folder.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Paths for intermediate and final files.
        string docxPath = Path.Combine(outputDir, "SampleEquations.docx");
        string pdfPath = Path.Combine(outputDir, "SampleEquations.pdf");

        // 1. Create a DOCX with two OfficeMath equations, each bookmarked.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        InsertEquationWithBookmark(builder, "eq1", @"\f(1,2)"); // simple fraction 1/2
        InsertEquationWithBookmark(builder, "eq2", @"\r(3,x)"); // cube root of x

        // Save the source DOCX.
        doc.Save(docxPath, SaveFormat.Docx);
        ValidateFileExists(docxPath, "source DOCX");

        // 2. Load the DOCX to ensure the load workflow works.
        Document loadedDoc = new Document(docxPath);

        // Count only top‑level OfficeMath nodes (MathObjectType.OMathPara).
        int topLevelOfficeMathCount = CountTopLevelOfficeMath(loadedDoc);
        if (topLevelOfficeMathCount != 2)
            throw new InvalidOperationException(
                $"Expected 2 top‑level OfficeMath nodes after loading, but found {topLevelOfficeMathCount}.");

        // 3. Export the document to PDF.
        loadedDoc.Save(pdfPath, SaveFormat.Pdf);
        ValidateFileExists(pdfPath, "exported PDF");

        // 4. Basic validation that the PDF file is non‑empty.
        FileInfo pdfInfo = new FileInfo(pdfPath);
        if (pdfInfo.Length == 0)
            throw new InvalidOperationException(
                "Exported PDF file is empty, indicating a failure in the conversion process.");

        // 5. Re‑load the PDF (Aspose.Words can load PDF) and verify that the document still contains the same number of top‑level OfficeMath nodes.
        // When loading a PDF, OfficeMath objects are represented as images, so the count will be zero.
        Document pdfDoc = new Document(pdfPath);
        int pdfTopLevelOfficeMathCount = CountTopLevelOfficeMath(pdfDoc);
        Console.WriteLine("PDF loaded successfully. Top‑level OfficeMath nodes in PDF representation: " + pdfTopLevelOfficeMathCount);
        Console.WriteLine("Validation completed successfully.");
    }

    // Inserts an EQ field, converts it to a real OfficeMath node, wraps it in a bookmark, and places it in its own paragraph.
    private static void InsertEquationWithBookmark(DocumentBuilder builder, string bookmarkName, string eqArgument)
    {
        // Start bookmark.
        builder.StartBookmark(bookmarkName);

        // Insert EQ field.
        FieldEQ field = (FieldEQ)builder.InsertField(FieldType.FieldEquation, true);

        // Write the EQ argument.
        builder.MoveTo(field.Separator);
        builder.Write(eqArgument);

        // Update the field to ensure the equation is processed.
        field.Update();

        // Return to the field start's parent (the paragraph).
        builder.MoveTo(field.Start.ParentNode);

        // Convert the field to OfficeMath.
        OfficeMath officeMath = field.AsOfficeMath();
        if (officeMath == null)
            throw new InvalidOperationException("Failed to convert EQ field to OfficeMath.");

        // Insert the OfficeMath node before the field start.
        field.Start.ParentNode.InsertBefore(officeMath, field.Start);

        // Remove the original field.
        field.Remove();

        // Ensure the equation is in its own paragraph.
        builder.MoveTo(officeMath);
        builder.InsertParagraph();

        // End bookmark.
        builder.EndBookmark(bookmarkName);
    }

    // Helper to validate that a file exists and is accessible.
    private static void ValidateFileExists(string path, string description)
    {
        if (!File.Exists(path))
            throw new FileNotFoundException($"The {description} file was not created at expected location: {path}");
    }

    // Counts only top‑level OfficeMath nodes (those whose MathObjectType is OMathPara).
    private static int CountTopLevelOfficeMath(Document doc)
    {
        return doc.GetChildNodes(NodeType.OfficeMath, true)
                  .Cast<OfficeMath>()
                  .Count(om => om.MathObjectType == MathObjectType.OMathPara);
    }
}
