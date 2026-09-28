using System;
using System.Collections.Generic;

namespace MailMergeDemo
{
    // Arguments for the ImageFieldMerging event
    public class ImageFieldMergingArgs : EventArgs
    {
        public string FieldName { get; }
        public object FieldValue { get; }
        public string ImagePath { get; set; }
        public bool Handled { get; set; }

        public ImageFieldMergingArgs(string fieldName, object fieldValue)
        {
            FieldName = fieldName;
            FieldValue = fieldValue;
            ImagePath = string.Empty;
            Handled = false;
        }
    }

    // Delegate for the ImageFieldMerging event
    public delegate void ImageFieldMergingEventHandler(object sender, ImageFieldMergingArgs e);

    // Simple mail‑merge engine that raises ImageFieldMerging for image fields
    public class MailMergeEngine
    {
        private readonly Dictionary<string, object> _fields = new Dictionary<string, object>();

        // Event raised when an image field is being merged
        public event ImageFieldMergingEventHandler ImageFieldMerging;

        // Add a field to the data source
        public void AddField(string name, object value)
        {
            _fields[name] = value;
        }

        // Execute the merge (simulated)
        public void Execute()
        {
            foreach (var kvp in _fields)
            {
                // For this demo we treat any field whose name ends with "Image" as an image field
                if (kvp.Key.EndsWith("Image", StringComparison.OrdinalIgnoreCase))
                {
                    var args = new ImageFieldMergingArgs(kvp.Key, kvp.Value);
                    OnImageFieldMerging(args);

                    if (args.Handled)
                    {
                        Console.WriteLine($"Image field '{args.FieldName}' merged with custom image path: {args.ImagePath}");
                    }
                    else
                    {
                        Console.WriteLine($"Image field '{args.FieldName}' merged with default handling.");
                    }
                }
                else
                {
                    Console.WriteLine($"Text field '{kvp.Key}' merged with value: {kvp.Value}");
                }
            }
        }

        protected virtual void OnImageFieldMerging(ImageFieldMergingArgs e)
        {
            ImageFieldMerging?.Invoke(this, e);
        }
    }

    public class Program
    {
        public static void Main()
        {
            var engine = new MailMergeEngine();

            // Add sample fields
            engine.AddField("FirstName", "John");
            engine.AddField("LastName", "Doe");
            engine.AddField("ProfileImage", "john_doe"); // value used to build image file name

            // Subscribe to the ImageFieldMerging event to provide a custom image path
            engine.ImageFieldMerging += (sender, e) =>
            {
                // Custom logic: build an image path based on the field value
                if (e.FieldValue is string imageKey && !string.IsNullOrWhiteSpace(imageKey))
                {
                    e.ImagePath = $"./Images/{imageKey}.png";
                    e.Handled = true; // indicate that we have handled the insertion
                }
            };

            // Run the merge simulation
            engine.Execute();

            // Program ends automatically
        }
    }
}
