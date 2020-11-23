using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using JsonToXlsx.NETCore.Extensions;
using Newtonsoft.Json.Linq;
using OfficeOpenXml;

namespace JsonToXlsx.NETCore
{
    public static class JsonToXlsxConverter
    {
        public static void SaveAs(string json, string saveToPath)
        {
            JArray array = JArray.Parse(json);

            DataTable dataTable = CreateEmptyDataTable(array, Path.GetFileName(saveToPath));
            
            foreach (JObject jObject in array.Children<JObject>())
            {
                DataRow row = dataTable.NewRow();
                
                foreach (JProperty jProperty in jObject.Children<JProperty>())
                {
                    if (jProperty.Value.Type == JTokenType.Null)
                    {
                        continue;
                    }
                    
                    row[jProperty.Name] = jProperty.Value;
                }

                dataTable.Rows.Add(row);
                dataTable.AcceptChanges();
            }

            dataTable.WriteToXlsx(Path.Combine(Path.GetDirectoryName(saveToPath), Path.GetFileName(saveToPath)));
        }

        private static DataTable CreateEmptyDataTable(JArray jArray, string tableName)
        {
            DataTable dataTable = new DataTable(tableName);
            
            List<(string propertyName, int index)> notDefinedProperties = 
                new List<(string propertyName, int index)>();

            foreach (JObject jObject in jArray.Children<JObject>())
            {
                foreach ((JProperty jProperty, int index) in jObject
                    .Children<JProperty>()
                    .Select((prop, index) => (prop, index)))
                {
                    if (dataTable.Columns.Contains(jProperty.Name))
                    {
                        continue;
                    }
                    
                    Type type = GetSuitableType(jProperty.Value.Type);

                    if (type == null)
                    {
                        if (!notDefinedProperties.Exists(
                            pair => 
                                pair.propertyName == jProperty.Name &&
                                pair.index == index))
                        {
                            notDefinedProperties.Add((jProperty.Name, index));
                        }
                        continue;
                    }

                    DataColumn currentColumn = dataTable.Columns.Add(jProperty.Name, type);
                    if (index < dataTable.Columns.Count)
                    {
                        currentColumn.SetOrdinal(index);
                    }

                    if (notDefinedProperties.Exists(
                        pair => 
                            pair.propertyName == jProperty.Name &&
                            pair.index == index))
                    {
                        notDefinedProperties.Remove((jProperty.Name, index));
                    }
                }
            }

            if (notDefinedProperties.Any())
            {
                foreach ((string propertyName, int index) in notDefinedProperties)
                {
                    DataColumn currentColumn = dataTable.Columns.Add(propertyName, typeof(String));
                    if (index < dataTable.Columns.Count)
                    {
                        currentColumn.SetOrdinal(index);
                    }
                }
            }

            return dataTable;
        }

        private static Type GetSuitableType(JTokenType type)
        {
            return type switch
            {
                JTokenType.Integer => typeof(Int32),
                JTokenType.Float => typeof(Single),
                JTokenType.String => typeof(String),
                JTokenType.Boolean => typeof(Boolean),
                JTokenType.Date => typeof(DateTime),
                JTokenType.Null => null,
                JTokenType.Undefined => null,
                _ => throw new NotImplementedException($"Type '{type:G}' is not implemented yet!")
            };
        }

        public static JArray XlsxToJson(string filePath, bool isFirstRowHeaders = true)
        {
            if (String.IsNullOrEmpty(filePath))
            {
                return null;
            }
            
            FileInfo fileInfo = new FileInfo(filePath);
            
            JArray jArray = new JArray();
            
            using (ExcelPackage excelPackage = new ExcelPackage(fileInfo))
            {
                ExcelWorksheet excelWorksheet = excelPackage.Workbook.Worksheets.First();

                List<string> fieldNames;

                if (isFirstRowHeaders)
                {
                    fieldNames = new List<string>();
                    for (int i = 1; i <= excelWorksheet.Dimension.End.Column; i++)
                    {
                        fieldNames.Add(excelWorksheet.Cells[1, i].Value.ToString());
                    }
                }
                else
                {
                    fieldNames = Enumerable.Range(1, excelWorksheet.Dimension.End.Column + 1)
                        .Select(n => $"column_{n}")
                        .ToList();
                }

                for (int row = isFirstRowHeaders ? 2 : 1; row <= excelWorksheet.Dimension.End.Row; row++)
                {
                    JObject rowJObject = new JObject();

                    for (int column = 1; column <= excelWorksheet.Dimension.End.Column; column++)
                    {
                        object value = excelWorksheet.Cells[row, column].Value;

                        if (value is String stringValue &&
                            DateTime.TryParse(stringValue, out DateTime dateTimeValue))
                        {
                            rowJObject.Add(new JProperty(
                                fieldNames[column - 1],
                                dateTimeValue));
                        }
                        else
                        {
                            rowJObject.Add(new JProperty(
                                fieldNames[column - 1],
                                excelWorksheet.Cells[row, column].Value));
                        }
                    }
                    
                    jArray.Add(rowJObject);
                }
            }

            return jArray;
        }
    }
}