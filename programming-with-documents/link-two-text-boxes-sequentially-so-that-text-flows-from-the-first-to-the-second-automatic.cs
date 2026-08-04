using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert the first floating text box.
        Shape shape1 = builder.InsertShape(ShapeType.TextBox, 200, 100);
        TextBox textBox1 = shape1.TextBox;

        // Insert the second floating text box.
        Shape shape2 = builder.InsertShape(ShapeType.TextBox, 200, 100);
        TextBox textBox2 = shape2.TextBox;

        // Link the first text box to the second one if the link is valid.
        if (textBox1.IsValidLinkTarget(textBox2))
        {
            textBox1.Next = textBox2;
        }

        // Move the cursor inside the first text box and add a long paragraph.
        // The text that does not fit into the first box will automatically flow into the linked second box.
        builder.MoveTo(shape1.LastParagraph);
        builder.Writeln("Lorem ipsum dolor sit amet, consectetur adipiscing elit. " +
                        "Sed non risus. Suspendisse lectus tortor, dignissim sit amet, " +
                        "adipiscing nec, ultricies sed, dolor. Cras elementum ultrices diam. " +
                        "Maecenas ligula massa, varius a, semper congue, euismod non, mi. " +
                        "Proin porttitor, orci nec nonummy molestie, enim est eleifend mi, " +
                        "non fermentum diam nisl sit amet erat. Duis semper. " +
                        "Duis arcu massa, scelerisque vitae, consequat in, pretium a, enim. " +
                        "Pellentesque congue. Ut in risus volutpat libero pharetra tempor. " +
                        "Cras vestibulum bibendum augue. Praesent egestas leo in pede. " +
                        "Praesent blandit odio eu enim. Pellentesque sed dui ut augue blandit sodales. " +
                        "Vestibulum ante ipsum primis in faucibus orci luctus et ultrices posuere cubilia Curae; " +
                        "Aliquam nibh. Mauris ac mauris sed pede pellentesque fermentum. " +
                        "Maecenas adipiscing ante non diam sodales hendrerit.");

        // Save the document to the current directory.
        string outputPath = Path.Combine(Environment.CurrentDirectory, "LinkedTextBoxes.docx");
        doc.Save(outputPath);
    }
}
