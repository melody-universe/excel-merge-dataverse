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
            Console.WriteLine(
                Convert.ToBase64String(
                    template
                )
            );
            return;
            var inputFiles = GetInputFilesContents();
            var outputContents = Excel.Merge(
                template,
                inputFiles,
                new int[] {
                    5, 6, 7, 8
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
