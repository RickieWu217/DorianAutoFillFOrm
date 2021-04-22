﻿using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Linq;
using System.Reflection;

namespace SeleniumExample
{
    public static class MyServices
    {
        public static IEnumerable<T> ParseExcelDataToObject<T>(string path) where T : new()
        {
            IEnumerable<T> excelDto;
            using (FileStream fileStream = new FileStream(path, FileMode.Open))
            {
                ExcelPackage excel = new ExcelPackage(fileStream);
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                var workSheet = excel.Workbook.Worksheets.First();
                excelDto = workSheet.ConvertSheetToObjects<T>();
            }
            return excelDto;
        }


        public static IEnumerable<T> ConvertSheetToObjects<T>(this ExcelWorksheet worksheet) where T : new()
        {

            Func<CustomAttributeData, bool> columnOnly = y => y.AttributeType == typeof(ColumnIndex);

            var columns = typeof(T).GetProperties()
                .Where(x => x.CustomAttributes.Any(columnOnly))
                .Select(p => new
                {
                    Property = p,
                    Column = p.GetCustomAttributes<ColumnIndex>().First().Index //safe because if where above
                }).ToList();


            var rows = worksheet.Cells
                .Select(cell => cell.Start.Row)
                .Distinct()
                .OrderBy(x => x);


            //Create the collection container
            var collection = rows.Skip(1)
                .Select(row =>
                {
                    var tnew = new T();
                    columns.ForEach(col =>
                    {
                        //This is the real wrinkle to using reflection - Excel stores all numbers as double including int
                        var val = worksheet.Cells[row, col.Column];
                        //If it is numeric it is a double since that is how excel stores all numbers
                        if (val.Value == null)
                        {
                            col.Property.SetValue(tnew, null);
                            return;
                        }
                        if (col.Property.PropertyType == typeof(Int32))
                        {
                            col.Property.SetValue(tnew, val.GetValue<int>());
                            return;
                        }
                        if (col.Property.PropertyType == typeof(double))
                        {
                            col.Property.SetValue(tnew, val.GetValue<double>());
                            return;
                        }
                        if (col.Property.PropertyType == typeof(DateTime))
                        {
                            col.Property.SetValue(tnew, val.GetValue<DateTime>());
                            return;
                        }
                        //Its a string
                        col.Property.SetValue(tnew, val.GetValue<string>());
                    });

                    return tnew;
                });


            //Send it back
            return collection;
        }


    }
}


    [AttributeUsage(AttributeTargets.All)]
    public class ColumnIndex : System.Attribute
    {
        public int Index { get; set; }


        public ColumnIndex(int column)
        {
            Index = column;
        }
    }
}
