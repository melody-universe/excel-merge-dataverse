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
            IEnumerable<string> keyColumns)
        {
            var uniqueKeys = new HashSet<IList<object>>(
                new ListComprarer<object>()
            );
            return template.FromExcelPackage(package =>
            {
                var worksheet = package.Workbook.Worksheets[0];
                var rowIndex = 2;
                var templateColumnMap = CreateColumnMap(worksheet);
                foreach(var inputFile in inputFiles)
                {
                    inputFile.WithExcelPackage(inputPackage =>
                    {
                        var inputWorksheet = inputPackage.Workbook.Worksheets[0];
                        var inputColumnMap = CreateColumnMap(inputWorksheet);
                        for(var inputRowIndex = 2;
                            inputRowIndex <= inputWorksheet.Dimension.Rows;
                            inputRowIndex++)
                        {
                            var key = new List<object>();
                            bool atLeastOneNonNull = false;
                            foreach (var columnLabel in keyColumns)
                            {
                                var objValue = inputWorksheet.Cells[
                                    inputRowIndex,
                                    inputColumnMap.Forward[columnLabel]
                                ].Value;
                                object keyValue;
                                if(objValue == null)
                                {
                                    keyValue = null;
                                } else
                                {
                                    var type = objValue.GetType();
                                    if (type == typeof(string))
                                    {
                                        var strValue = ((string)objValue).Trim();
                                        if (string.IsNullOrEmpty(strValue))
                                        {
                                            keyValue = null;
                                        }
                                        else
                                        {
                                            keyValue = strValue;
                                        }
                                    }
                                    else
                                    {
                                        keyValue = objValue;
                                    }
                                }
                                key.Add(keyValue);
                                if(keyValue != null && !atLeastOneNonNull)
                                {
                                    atLeastOneNonNull = true;
                                }
                            }
                            if(!atLeastOneNonNull)
                            {
                                continue;
                            }
                            if(!uniqueKeys.Contains(key))
                            {
                                uniqueKeys.Add(key);
                                for(var sourceColumnIndex = 1;
                                    sourceColumnIndex < inputWorksheet.Dimension.Columns;
                                    sourceColumnIndex++)
                                {
                                    var destinationColumnIndex = templateColumnMap.Forward[
                                        inputColumnMap.Reverse[sourceColumnIndex]];
                                    worksheet.Cells[rowIndex, destinationColumnIndex].Value =
                                        inputWorksheet.Cells[inputRowIndex, sourceColumnIndex].Value;
                                }
                                rowIndex++;
                            }
                        }
                    });
                }
                return package.ToByteArray();
            });
        }

        private static Map<string, int> CreateColumnMap(ExcelWorksheet worksheet)
        {
            var map = new Map<string, int>();
            for(var column = 1; column <worksheet.Dimension.Columns; column++)
            {
                var label = ((string)worksheet.Cells[1, column].Value).Trim();
                map.Add(label, column);
            }
            return map;
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
                    if((x[i] == null && y[i] != null)
                        || (x[i] != null && y[i] == null)
                        || (x[i] != null && y[i] != null && !x[i].Equals(y[i])))
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
                    (current, next) => current + (next?.GetHashCode() ?? 0)
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
