using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Math;

public class OfficeMathSaveExample
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Introductory paragraph.
        builder.Writeln("Sample document with OfficeMath equations:");

        // Insert three safe EQ‑field equations.
        InsertOfficeMath(builder, @"\f(1,2)");          // Fraction 1/2
        InsertOfficeMath(builder, @"\r(3,x)");          // Cube root of x
        InsertOfficeMath(builder, @"\i \su(n=1,5,n)"); // Integral with summation

        // Save the document as DOCX.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "ModifiedDocument.docx");
        doc.Save(outputPath, SaveFormat.Docx);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The document was not saved correctly.");

        Console.WriteLine("Document saved to: " + outputPath);
    }

    // Inserts an EQ field, converts it to a real OfficeMath node,
    // applies display formatting to top‑level equations,
    // and removes the original field.
    private static void InsertOfficeMath(DocumentBuilder builder, string eqArguments)
    {
        // Insert an empty EQ field.
        FieldEQ field = (FieldEQ)builder.InsertField(FieldType.FieldEquation, true);

        // Write the EQ arguments after the field separator.
        builder.MoveTo(field.Separator);
        builder.Write(eqArguments);

        // Update the field so that Word evaluates the EQ code.
        field.Update();

        // Convert the EQ field to an OfficeMath object.
        OfficeMath officeMath = field.AsOfficeMath();

        if (officeMath == null)
            throw new InvalidOperationException("Failed to convert EQ field to OfficeMath.");

        // Insert the OfficeMath node before the field start.
        field.Start.ParentNode.InsertBefore(officeMath, field.Start);

        // Remove the original EQ field.
        field.Remove();

        // Apply display formatting only to top‑level equations.
        if (officeMath.MathObjectType == MathObjectType.OMathPara)
        {
            officeMath.DisplayType = OfficeMathDisplayType.Display;
            officeMath.Justification = OfficeMathJustification.Left;
        }

        // Move the builder to the paragraph after the inserted equation
        // so subsequent insertions continue on a new line.
        builder.MoveTo(officeMath.ParentParagraph);
        builder.InsertParagraph();
    }
}
