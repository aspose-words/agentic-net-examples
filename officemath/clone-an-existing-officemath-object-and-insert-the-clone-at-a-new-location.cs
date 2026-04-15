using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Math;

public class CloneOfficeMathExample
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a paragraph with introductory text.
        builder.Writeln("Original equation follows:");

        // Insert an EQ field that will be converted to a real OfficeMath object.
        // The field is created with the EQ switch; the actual equation arguments are written after the separator.
        FieldEQ eqField = (FieldEQ)builder.InsertField(FieldType.FieldEquation, true);

        // Move to the field separator and write a simple fraction equation: \f(1,2)
        builder.MoveTo(eqField.Separator);
        builder.Write(@"\f(1,2)");

        // Update the field so that the EQ code is processed before conversion.
        eqField.Update();

        // Return the builder to the field's parent paragraph.
        builder.MoveTo(eqField.Start.ParentNode);

        // Convert the EQ field to an OfficeMath node.
        OfficeMath originalMath = eqField.AsOfficeMath();
        if (originalMath == null)
            throw new InvalidOperationException("Failed to convert EQ field to OfficeMath.");

        // Replace the field with the real OfficeMath node.
        eqField.Start.ParentNode.InsertBefore(originalMath, eqField.Start);
        eqField.Remove();

        // Clone the existing OfficeMath node (deep clone).
        OfficeMath clonedMath = (OfficeMath)originalMath.Clone(true);
        if (clonedMath == null)
            throw new InvalidOperationException("Failed to clone OfficeMath node.");

        // Insert the cloned OfficeMath into a new paragraph after the original equation.
        builder.MoveTo(originalMath.ParentParagraph);
        builder.InsertParagraph(); // Creates a new paragraph after the current one.
        Paragraph newParagraph = (Paragraph)builder.CurrentParagraph;
        newParagraph.AppendChild(clonedMath);

        // Save the resulting document.
        string outputPath = "ClonedOfficeMath.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new FileNotFoundException("The output document was not created.", outputPath);
    }
}
