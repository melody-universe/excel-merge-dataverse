using ExcelMerge.Library;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MergeDocumentsConsoleApp
{
    class Program
    {
        static void Main(string[] args)
        {
            var template = GetTemplateContents();
            var inputFiles = GetInputFilesContents();
            var outputContents = Excel.Merge(
                template,
                inputFiles,
                new string[] {
                    "Facility Name",
                    "Facility Street Address",
                    "Facility City",
                    "Facility Zip",
                    "Visit Date"
                }
            );
            const string outputFileName = "Output.xlsx";
            if (File.Exists(outputFileName))
            {
                File.Delete(outputFileName);
            }
            File.WriteAllBytes(outputFileName, outputContents);
        }

        static byte[] GetTemplateContents()
        {
            return File.ReadAllBytes("./Template.xlsx");
        }

        static IEnumerable<byte[]> GetInputFilesContents()
        {
            var filesContents = new List<byte[]>();
            string fileName;
            for(var i = 1; File.Exists(fileName = $"./File{i}.xlsx"); i++)
            {
                filesContents.Add(
                    File.ReadAllBytes(fileName)
                );
            }
            return filesContents;
        }
    }
}
