using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Excel;
using Newtonsoft.Json;
using Office = Microsoft.Office.Core;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using System;
using Newtonsoft.Json.Linq;
using System.Diagnostics;

namespace u7ha1ExcelAddIn
{
   public partial class ThisAddIn
   {
      private enum VocabularyProperties
      {
         Index = 1,
         Content,
         Furigana,
         Romanji,
         HanViet,
         Meaning,
         Group,
         Key
      }

      private const string FileName = "data.js";
      private const string ProjectName = "u7ha1";

      private void WorkbookAfterSaveEventHandler(Excel.Workbook Wb, bool Success)
      {
         this.Save(GetContent(Wb));
      }

      private string GetContent(Excel.Workbook Wb)
      {
         string result = string.Empty;

         Excel.Worksheet worksheet = this.GetProjectWorksheet(Wb);

         if (worksheet != null)
         {
            Excel.ListObject table = this.GetProjectTable(worksheet);

            if (table != null)
            {
               result += "var data = ";

               JObject json = this.GetJsonAndFormat(table);

               result += json.ToString();
            }
         }

         return result;
      }

      private Excel.Worksheet GetProjectWorksheet(Excel.Workbook Wb)
      {
         Excel.Worksheet result = null;

         foreach (Excel.Worksheet worksheet in Wb.Worksheets)
         {
            if (worksheet.Name == ProjectName)
            {
               result = worksheet;

               break;
            }
         }

         return result;
      }

      private Excel.ListObject GetProjectTable(Excel.Worksheet worksheet)
      {
         Excel.ListObject result = null;

         foreach (Excel.ListObject listObject in worksheet.ListObjects)
         {
            if (listObject.Name == ProjectName)
            {
               result = listObject;

               break;
            }
         }

         return result;
      }

      private JObject GetJsonAndFormat(Excel.ListObject table)
      {
         JObject result = new JObject();

         Excel.Range headerRow = table.HeaderRowRange;
         Excel.Range entries = table.Range.Rows;

         int columnsCount = headerRow.Columns.Count;
         int rowsCount = entries.Rows.Count;

         for (int i = 2; i <= rowsCount; i++)
         {
            Excel.Range entry = entries[i];

            Excel.Range cell = entry.Cells[1, VocabularyProperties.Key];

            string key = string.Empty;

            if (cell.Value2 == null)
            {
               continue;
            }
            else
            {
               key = cell.Value2.ToString();

               if (string.IsNullOrEmpty(key))
               {
                  continue;
               }
            }

            JObject vocabulary = new JObject();

            result.Add(key, vocabulary);

            for (int j = 1; j <= columnsCount; j++)
            {
               cell = entry.Cells[1, j];

               string value = string.Empty;

               if (cell.Value2 != null)
               {
                  value = cell.Value2.ToString();

                  value = value.Trim();
               }

               switch ((VocabularyProperties)j)
               {
                  case VocabularyProperties.Index:
                     {
                        vocabulary.Add("index", value);
                     }
                     break;

                  case VocabularyProperties.Content:
                     {
                        vocabulary.Add("content", value);
                     }
                     break;

                  case VocabularyProperties.Furigana:
                     {
                        vocabulary.Add("furigana", value);
                     }
                     break;

                  case VocabularyProperties.Romanji:
                     {
                        value = value.ToLower();

                        vocabulary.Add("romanji", value);
                     }
                     break;

                  case VocabularyProperties.HanViet:
                     {
                        value = value.ToUpper();

                        vocabulary.Add("han-viet", value);
                     }
                     break;

                  case VocabularyProperties.Meaning:
                     {
                        value = this.UppercaseFirst(value);

                        vocabulary.Add("meaning", value);
                     }
                     break;

                  case VocabularyProperties.Group:
                     {
                        vocabulary.Add("group", value);
                     }
                     break;

                  case VocabularyProperties.Key:
                     {
                        vocabulary.Add("key", value);
                     }
                     break;

                  default:
                     {
                     }
                     break;
               }

               cell.Value = value;
            }
         }

         entries.AutoFit();
         entries.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;

         return result;
      }

      private string UppercaseFirst(string content)
      {
         if (string.IsNullOrEmpty(content))
         {
            return string.Empty;
         }

         char[] array = content.ToCharArray();

         array[0] = char.ToUpper(array[0]);

         return new string(array);
      }

      private void Save(string Content)
      {
         string directory = this.Application.ActiveWorkbook.Path + @"\js\";

         if (!Directory.Exists(directory))
         {
            Directory.CreateDirectory(directory);
         }

         using (StreamWriter streamWriter = new StreamWriter(directory + FileName))
         {
            streamWriter.WriteLine(Content);
         }
      }
   }
}
