using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Math;

public class ReplaceOfficeMathExample
{
    public static void Main()
    {
        // Prepare output folder.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);
        string outputPath = Path.Combine(outputDir, "ReplaceOfficeMath.docx");

        // 1. Create a new document and a builder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // 2. Insert an initial equation using the EQ field bootstrap workflow.
        //    This will create a fraction 1/2 as the first OfficeMath object.
        FieldEQ initialField = InsertFieldEQ(builder, @"\f(1,2)");
        // Ensure the field is up‑to‑date before conversion.
        initialField.Update();

        OfficeMath oldOfficeMath = initialField.AsOfficeMath()
            ?? throw new InvalidOperationException("Failed to convert initial EQ field to OfficeMath.");

        // Replace the field with the real OfficeMath node.
        initialField.Start.ParentNode.InsertBefore(oldOfficeMath, initialField.Start);
        initialField.Remove();

        // 3. Create a new equation that will replace the existing one.
        //    Example: a cubic root of x using the radical switch.
        // Move the builder to the paragraph that contains the old equation.
        builder.MoveTo(oldOfficeMath.ParentParagraph);
        // Insert a new EQ field after the old equation.
        FieldEQ newField = InsertFieldEQ(builder, @"\r(3,x)");
        newField.Update();

        OfficeMath newOfficeMath = newField.AsOfficeMath()
            ?? throw new InvalidOperationException("Failed to convert new EQ field to OfficeMath.");

        // Insert the new OfficeMath node after the old one and remove the old node.
        oldOfficeMath.ParentNode.InsertAfter(newOfficeMath, oldOfficeMath);
        oldOfficeMath.Remove();

        // Remove the temporary field.
        newField.Remove();

        // 4. Save the resulting document.
        doc.Save(outputPath);
    }

    // Helper method that inserts an EQ field, writes the argument string,
    // updates the field, and moves the builder back to the field's parent paragraph.
    private static FieldEQ InsertFieldEQ(DocumentBuilder builder, string args)
    {
        // Insert an empty EQ field.
        FieldEQ field = (FieldEQ)builder.InsertField(FieldType.FieldEquation, true);

        // Write the equation arguments into the field separator.
        builder.MoveTo(field.Separator);
        builder.Write(args);

        // Return the builder to the field's parent paragraph.
        builder.MoveTo(field.Start.ParentNode);
        builder.InsertParagraph(); // Ensure the next content starts on a new line.

        return field;
    }
}
