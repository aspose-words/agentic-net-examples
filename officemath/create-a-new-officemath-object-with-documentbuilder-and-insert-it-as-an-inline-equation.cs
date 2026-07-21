using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Math;

public class OfficeMathExample
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add some introductory text.
        builder.Writeln("This paragraph precedes an inline equation:");

        // Insert an EQ field that will be converted to OfficeMath.
        Field field = builder.InsertField(FieldType.FieldEquation, true);
        FieldEQ fieldEq = field as FieldEQ;
        if (fieldEq == null)
            throw new InvalidOperationException("Failed to create FieldEQ.");

        // Write a simple EQ argument (fraction 1/2) at the field separator.
        builder.MoveTo(fieldEq.Separator);
        builder.Write(@"\f(1,2)");

        // Update the field so that it can be converted to a real OfficeMath object.
        field.Update();

        // Convert the field to an OfficeMath object.
        OfficeMath officeMath = fieldEq.AsOfficeMath();
        if (officeMath == null)
            throw new InvalidOperationException("EQ field could not be converted to OfficeMath.");

        // Insert the OfficeMath node before the field start and remove the original field.
        CompositeNode parent = field.Start.ParentNode as CompositeNode;
        if (parent == null)
            throw new InvalidOperationException("Field start does not have a valid composite parent.");

        parent.InsertBefore(officeMath, field.Start);
        field.Remove();

        // Move the builder after the inserted OfficeMath so we can continue writing.
        builder.MoveTo(officeMath);
        builder.Writeln(); // End the paragraph that contains the equation.

        // Add some trailing text.
        builder.Writeln("This paragraph follows the inline equation.");

        // Save the document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "OfficeMathExample.docx");
        doc.Save(outputPath);

        // Validate that the file was created.
        if (!File.Exists(outputPath))
            throw new FileNotFoundException("The output document was not saved.", outputPath);

        // Reload the document and verify the OfficeMath node exists.
        Document loadedDoc = new Document(outputPath);
        int mathCount = loadedDoc.GetChildNodes(NodeType.OfficeMath, true).Count;
        if (mathCount == 0)
            throw new InvalidOperationException("No OfficeMath nodes were found in the saved document.");

        // Indicate successful completion.
        Console.WriteLine($"Document saved to '{outputPath}' with {mathCount} OfficeMath node(s).");
    }
}
