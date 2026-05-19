using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Math;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a paragraph with introductory text.
        builder.Writeln("Original equation:");

        // Insert an EQ field that will be converted to a real OfficeMath object.
        Field field = builder.InsertField(FieldType.FieldEquation, true);
        FieldEQ fieldEQ = (FieldEQ)field;

        // Write a simple fraction equation into the field separator.
        builder.MoveTo(fieldEQ.Separator);
        builder.Write(@"\f(1,2)"); // Fraction 1/2.

        // Update the field so that the EQ code is recognized.
        field.Update();

        // Return the builder cursor to the paragraph that contains the field.
        builder.MoveTo(fieldEQ.Start.ParentNode);

        // Convert the EQ field to an OfficeMath object.
        OfficeMath originalMath = fieldEQ.AsOfficeMath();
        if (originalMath == null)
            throw new InvalidOperationException("Failed to convert EQ field to OfficeMath.");

        // Insert the OfficeMath node before the field start node.
        fieldEQ.Start.ParentNode.InsertBefore(originalMath, fieldEQ.Start);

        // Remove the original field – the document now contains only the real OfficeMath node.
        fieldEQ.Remove();

        // Clone the existing OfficeMath node (deep clone to copy all child elements).
        OfficeMath clonedMath = (OfficeMath)originalMath.Clone(true);

        // Insert the cloned OfficeMath into a new paragraph after the original one.
        Paragraph originalParagraph = originalMath.ParentParagraph;
        builder.MoveTo(originalParagraph);
        builder.InsertParagraph(); // Creates a new empty paragraph after the current one.
        Paragraph newParagraph = builder.CurrentParagraph;
        newParagraph.PrependChild(clonedMath);

        // Save the document to disk.
        string outputPath = "ClonedOfficeMath.docx";
        doc.Save(outputPath);

        // Validate that the file was created.
        if (!File.Exists(outputPath))
            throw new FileNotFoundException("The output document was not saved.", outputPath);
    }
}
