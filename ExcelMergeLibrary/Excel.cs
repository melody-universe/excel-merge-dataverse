using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace ExcelMerge.Library
{
    public static class Excel
    {
        static Excel() {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        }

        public static byte[] Merge(
            byte[] template,
            IEnumerable<byte[]> inputFiles,
            IEnumerable<int> keyColumns)
        {
            var uniqueKeys = new HashSet<IList<string>>(
                new ListComprarer<string>()
            );
            return template.FromExcelPackage(package =>
            {
                var worksheet = package.Workbook.Worksheets[0];
                var rowIndex = 2;
                foreach(var inputFile in inputFiles)
                {
                    inputFile.WithExcelPackage(inputPackage =>
                    {
                        var inputWorksheet = inputPackage.Workbook.Worksheets[0];
                        for(var inputRowIndex = 2;
                            inputRowIndex <= inputWorksheet.Dimension.Rows;
                            inputRowIndex++)
                        {
                            var key = new List<string>();
                            foreach (var columnIndex in keyColumns)
                            {
                                var value = (string)inputWorksheet.Cells[
                                    inputRowIndex, columnIndex].Value;
                                key.Add(value);
                            }
                            if(!uniqueKeys.Contains(key))
                            {
                                uniqueKeys.Add(key);
                                for(var columnIndex = 1;
                                    columnIndex < inputWorksheet.Dimension.Columns;
                                    columnIndex++)
                                {
                                    worksheet.Cells[rowIndex, columnIndex].Value =
                                        inputWorksheet.Cells[inputRowIndex, columnIndex].Value;
                                }
                                rowIndex++;
                            }
                        }
                    });
                }
                return package.ToByteArray();
            });
        }

        private class ListComprarer<T> : EqualityComparer<IList<T>>
        {
            public override bool Equals(IList<T> x, IList<T> y)
            {
                if(x.Count != y.Count)
                {
                    return false;
                }
                for(var i = 0; i < x.Count; i++)
                {
                    if(!x[i].Equals(y[i]))
                    {
                        return false;
                    }
                }
                return true;
            }

            public override int GetHashCode(IList<T> obj)
            {
                return obj.Aggregate(
                    0,
                    (current, next) => current + next.GetHashCode()
                );
            }
        }

        private static void WithExcelPackage(this byte[] contents, Action<ExcelPackage> action)
        {
            using(var memoryStream = new MemoryStream(contents))
            {
                using (var package = new ExcelPackage(memoryStream))
                {
                    action(package);
                }
            }
        }

        private static TOutput FromExcelPackage<TOutput>(
            this byte[] contents, Func<ExcelPackage, TOutput> function)
        {
            TOutput output = default(TOutput);
            contents.WithExcelPackage(package =>
            {
                output = function(package);
            });
            return output;
        }

        private static byte[] ToByteArray(this ExcelPackage package) {
            using (var memoryStream = new MemoryStream())
            {
                package.SaveAs(memoryStream);
                return memoryStream.ToArray();
            }
        }
    }
}
