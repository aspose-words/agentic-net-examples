using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Math;

public class InsertOfficeMathFromMathML
{
    public static void Main()
    {
        // The original MathML is kept as a comment because Aspose.Words does not parse it directly.
        // This string is provided for reference only.
        string mathMl = @"<math xmlns=""http://www.w3.org/1998/Math/MathML""><mfrac><mi>a</mi><mi>b</mi></mfrac></math>";

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a paragraph that will contain the equation.
        builder.Writeln("Equation inserted from MathML:");

        // Insert an EQ field that represents a simple fraction a/b.
        // The "\f" switch creates a fraction; the arguments are the numerator and denominator.
        FieldEQ eqField = InsertFieldEQ(builder, @"\f(a,b)");

        // Ensure the field is up‑to‑date before converting it.
        eqField.Update();

        // Convert the EQ field to a real OfficeMath object.
        OfficeMath officeMath = eqField.AsOfficeMath();

        if (officeMath == null)
            throw new InvalidOperationException("Failed to convert EQ field to OfficeMath.");

        // Insert the OfficeMath node before the field start and remove the original field.
        eqField.Start.ParentNode.InsertBefore(officeMath, eqField.Start);
        eqField.Remove();

        // Set display formatting for the top‑level equation.
        officeMath.DisplayType = OfficeMathDisplayType.Display;
        officeMath.Justification = OfficeMathJustification.Left;

        // Save the document.
        string outputPath = Path.Combine(Environment.CurrentDirectory, "OfficeMathFromMathML.docx");
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new FileNotFoundException("The output document was not created.", outputPath);
    }

    // Helper that follows the deterministic EQ‑field bootstrap pattern.
    private static FieldEQ InsertFieldEQ(DocumentBuilder builder, string args)
    {
        // Insert an empty EQ field.
        FieldEQ field = (FieldEQ)builder.InsertField(FieldType.FieldEquation, true);

        // Move to the field separator and write the EQ argument string.
        builder.MoveTo(field.Separator);
        builder.Write(args);

        // Return the builder to the field's paragraph and start a new paragraph for subsequent content.
        builder.MoveTo(field.Start.ParentNode);
        builder.InsertParagraph();

        return field;
    }
}
