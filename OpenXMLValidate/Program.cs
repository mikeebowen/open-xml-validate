using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using System;
using System.Collections.Generic;
using System.Reflection;
using System.IO;

namespace inspect_it
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Enter the path to the OOXML file");
            string path = Console.ReadLine();

            OpenPresentation(path);
        }
        public static void OpenPresentation(string fileName)
        {
            //PresentationDocument doc = PresentationDocument.Open(fileName, false);
            if (File.Exists(fileName))
            {
                WordprocessingDocument doc = WordprocessingDocument.Open(fileName, false);
                OpenXmlValidator openXmlValidator = new OpenXmlValidator();
                IEnumerable<ValidationErrorInfo> validations = openXmlValidator.Validate(doc);
                foreach (ValidationErrorInfo validationErrorInfo in validations)
                {
                    Console.WriteLine($"Validation Error: {validationErrorInfo.Description}");
                    Console.WriteLine($"Validation XPath: {validationErrorInfo.Path.XPath}");
                }
            }
        }
    }
}
