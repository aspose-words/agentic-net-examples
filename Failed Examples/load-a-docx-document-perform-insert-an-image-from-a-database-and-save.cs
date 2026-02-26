// ALL ATTEMPTS FAILED. Below is the last generated code.

using System;
using System.Data;               // For database access (placeholder)
using System.Data.SqlClient;    // For SQL Server (example)
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Drawing;

class Program
{
    static void Main()
    {
        // Path to the source DOCX file.
        string docPath = @"C:\Input\Template.docx";

        // Load the existing Word document.
        Document doc = new Document(docPath);

        // Create a DocumentBuilder to modify the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Retrieve the image bytes from the database.
        // This method is a placeholder – replace with actual data‑access code.
        byte[] imageBytes = LoadImageFromDatabase();

        // Insert the image at the current cursor position.
        // Uses the InsertImage(byte[]) overload defined in DocumentBuilder.
        builder.InsertImage(imageBytes);

        // Prepare image save options to render the document as a PNG.
        ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Png)
        {
            // Optional: set resolution, background color, etc.
            Resolution = 300,                     // 300 DPI for high quality
            PaperColor = System.Drawing.Color.Transparent
        };

        // Path for the output PNG file.
        string pngPath = @"C:\Output\Result.png";

        // Save the document as a PNG image using the Save(string, SaveOptions) overload.
        doc.Save(pngPath, saveOptions);
    }

    // Placeholder method that simulates fetching an image from a database.
    // Replace the implementation with actual ADO.NET or ORM code as needed.
    static byte[] LoadImageFromDatabase()
    {
        // Example using a SQL Server connection.
        // Adjust connection string, query, and parameters to match your schema.
        const string connectionString = "Data Source=SERVER;Initial Catalog=MyDatabase;Integrated Security=True;";
        const string query = "SELECT ImageData FROM Images WHERE ImageId = @Id";

        using (SqlConnection conn = new SqlConnection(connectionString))
        using (SqlCommand cmd = new SqlCommand(query, conn))
        {
            cmd.Parameters.Add("@Id", SqlDbType.Int).Value = 1; // Example ID

            conn.Open();
            object result = cmd.ExecuteScalar();

            if (result != null && result != DBNull.Value)
                return (byte[])result;
            else
                throw new InvalidOperationException("Image not found in database.");
        }
    }
}
