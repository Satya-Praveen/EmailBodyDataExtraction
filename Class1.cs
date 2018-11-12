using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Text.RegularExpressions;
using System.Net;
using System.IO;

using System.Activities;
using System.ComponentModel;



namespace TAODataextraction
{
    public class EmailBodyDataextraction: CodeActivity
    {
        [Category("Input")]
        [RequiredArgument]
        [Description("Enter Path of Html Input File")]
        public InArgument<string> HtmlFilePath { get; set; }

        [Category("Input")]
        public InArgument<System.Data.DataTable> DataTable { get; set; }

        [Category("Output")]
        public OutArgument<List<string>> XmlString { get; set; }

        protected override void Execute(CodeActivityContext context)
        {
            var htmlFilePath = HtmlFilePath.Get(context);   
            var dataTable = DataTable.Get(context);

            if (!(String.IsNullOrEmpty(htmlFilePath)) & (dataTable)==null)
            {
                XmlString.Set(context, GetValueFromHTMLTags(htmlFilePath));
            }
            //throw new NotImplementedException();

            if (!(String.IsNullOrEmpty(htmlFilePath)) & (dataTable) != null)
            {
                XmlString.Set(context, GetValueFromHTMLTags(dataTable, htmlFilePath));
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="InputFilePath"></param>
        /// <returns></returns>
        public static List<string> GetValueFromHTMLTags(string InputFilePath)
        {

            if (string.IsNullOrEmpty(InputFilePath))
                return null;


            string lHTMLString = System.IO.File.ReadAllText(InputFilePath);
            string TableExpression = "<table[^>]*>(.*?)</table>";
            string HeaderExpression = "<th[^>]*>(.*?)</th>";
            string RowExpression = "<tr[^>]*>(.*?)</tr>";
            string ColumnExpression = @"<td[^>]*>(.*?)</td>";

            DataTable lDataTbl = null;
            DataRow lDataR = null;
            DataColumn DataCol = null;
            List<string> lXmlData = new List<string>();
            int iCurrentColumn = 0;
            int iCurrentRow = 0;
            int iTableCount = 0;

            bool HeadersExist = false;

            foreach (Match Table in Regex.Matches(lHTMLString, TableExpression, RegexOptions.Multiline | RegexOptions.Singleline | RegexOptions.IgnoreCase))
            {
                iCurrentRow = 0;
                HeadersExist = false;
                lDataTbl = new DataTable("HtmlTables" + iTableCount);

                if (Table.Value.ToLower().Contains("<th"))
                {
                    HeadersExist = true;
                    foreach (Match Header in Regex.Matches(Table.Value, HeaderExpression, RegexOptions.Multiline | RegexOptions.Singleline | RegexOptions.IgnoreCase))
                    {
                        lDataTbl.Columns.Add(Header.Groups[1].ToString());
                    }
                }


                foreach (Match Row in Regex.Matches(Table.Value, RowExpression, RegexOptions.Multiline | RegexOptions.Singleline | RegexOptions.IgnoreCase))
                {
                    if (!(iCurrentRow >= 0 & HeadersExist == true)) // need to relook into this code / line since we are checking the "HeadersExist =true"
                    {
                        lDataR = lDataTbl.NewRow();

                        iCurrentColumn = 0;

                        foreach (Match Column in Regex.Matches(Row.Value, ColumnExpression, RegexOptions.Multiline | RegexOptions.Singleline | RegexOptions.IgnoreCase))
                        {
                            DataColumnCollection columns = lDataTbl.Columns;
                            if (!columns.Contains("Column " + iCurrentColumn))
                            {
                                lDataTbl.Columns.Add("Column " + iCurrentColumn);
                            }

                            string TagReplacePattern = "<[^>]*>";
                            string cellValue = Column.Groups[1].ToString();
                            cellValue = Regex.Replace(cellValue, TagReplacePattern, string.Empty);

                            if (!string.IsNullOrEmpty(cellValue.ToString().Trim()))
                            {
                                lDataR[iCurrentColumn] = (cellValue.Replace("&nbsp;", string.Empty)).Trim();
                            }

                            iCurrentColumn += 1;
                        }

                        lDataTbl.Rows.Add(lDataR);
                    }
                    iCurrentRow += 1;
                }


                iTableCount += 1;
                //Adds Table only if the number of rows are greater than one
                if (lDataTbl.Rows.Count > 1)
                {
                    lXmlData.Add(ConvertDatatableToXML(lDataTbl));
                }

            }
            return (lXmlData);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="dt"></param>
        /// <param name="InputFilePath"></param>
        /// <returns></returns>
        public static List<string> GetValueFromHTMLTags(DataTable dt, string InputFilePath)
        {
            List<string> OutXmlStr = new List<string>();
            string[] HtmlTextLines = System.IO.File.ReadAllLines(InputFilePath);
            DataTable VendorRegexDT = dt;

            //Building Current Datatable
            DataTable table = new DataTable("Current");
            table.Columns.Add("Keywords", typeof(string));
            table.Columns.Add("MatchedValue", typeof(string));

            foreach (string LineItem in HtmlTextLines)
            {
                foreach (DataRow row in VendorRegexDT.Rows)
                {
                    if (LineItem.Contains(row["Keywords"].ToString()))
                    {
                        MatchCollection RegexMatchedValues = Regex.Matches(LineItem, row["Pattern"].ToString());
                        foreach (Match MatchedValue in RegexMatchedValues)
                        {
                            String Value = MatchedValue.Groups[0].Value;
                            //row["MatchedValue"] = Value;
                            //Building DataRow
                            DataRow dataRow = table.NewRow();
                            dataRow["Keywords"] = row["Keywords"].ToString();
                            dataRow["MatchedValue"] = Value;
                            table.Rows.Add(dataRow);
                        }
                    }
                }
            }
            //VendorRegexDT.Columns.Remove("Pattern");
            string xmlstr = ConvertDatatableToXML(table);
            OutXmlStr.Add(xmlstr);
            return OutXmlStr;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="dt"></param>
        /// <returns></returns>
        private static string ConvertDatatableToXML(DataTable dt)
        {
            MemoryStream str = new MemoryStream();
            dt.WriteXml(str, true);
            str.Seek(0, SeekOrigin.Begin);
            StreamReader sr = new StreamReader(str);
            string xmlstr;
            xmlstr = sr.ReadToEnd();
            return (xmlstr);
        }
    }
}
