using System;
using System.Linq;
using System.Data.Linq.Mapping;
using System.Reflection;
using System.Data.Linq;
using System.IO;
using System.Data;
using System.Data.OleDb;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using Microsoft.Office.Interop.Excel;
using System.Text.RegularExpressions;

namespace MyWeb.Processor
{
    public class ExcelConverter
    {
        //개행문자제거
        private string Cleaning(string source, string deleteText = @"\W")   //"(?<!\r)\n"
        {
            try
            {
                return Regex.Replace(source, deleteText, "");
                //var result = Regex.Replace("colName", "(?<!\r)\n", ""); //개행문자제거
                //var cleanChars = "colName".Where(c => !"\n\r".Contains(c)).ToList().ToString(); //개행문자제거
            }
            catch (Exception)
            {

                throw;
            }
        }
        //Excel to DataSet변환(Office.Interop방식)
        public DataSet OfficeExcelTODataSet(string fileName)
        {
            var findedCount = 0;
            char spliter = ',';
            string headerList = "EID,Inven,No.,SG No.,TID";
            int headerCellCount = headerList.Split(spliter).Length;
            int headerIndex = 1;

            DataSet ds = new DataSet();
            System.Data.DataTable dataTable = new System.Data.DataTable("ExcelImportFileInfo");
            ds.Tables.Add(dataTable);

            try
            {
                var missing = System.Reflection.Missing.Value;

                Application app = new Microsoft.Office.Interop.Excel.Application();

                //파일읽기
                Workbook workbook = app.Workbooks.Open(fileName, false, true, missing, missing, missing, true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, '\t', false, false, 0, false, true, 0);
                Worksheet worksheet = workbook.Worksheets[1] as Microsoft.Office.Interop.Excel.Worksheet; worksheet = (Microsoft.Office.Interop.Excel.Worksheet)workbook.Worksheets.get_Item(1);

                Range xlRange = worksheet.UsedRange;
                Array myValues = (Array)xlRange.Cells.Value2;

                int vertical = myValues.GetLength(0);
                int horizontal = myValues.GetLength(1);
                string columnName = "";
                System.Data.DataTable dt = new System.Data.DataTable("ExcelImportedData");

                // must start with index = 1                
                // 여기서 Header를 찾는 Logic구현한다.                
                for (int a = 1; a <= vertical; a++)
                {
                    for (int i = 1; i <= horizontal; i++)
                    {
                        //일치하는 경우
                        if (myValues.GetValue(a, i) != null)
                        {
                            columnName = myValues.GetValue(a, i).ToString().Trim();
                            if (columnName.Length > 0 && headerList.IndexOf(columnName) > -1)
                            {
                                findedCount++;
                            }
                        }

                    }
                    if (findedCount >= 3) //기본 Header를 모두 찾으면
                    {
                        headerIndex = a;
                        break;
                    }
                }
                // get header information
                int extIndex = 1;
                for (int i = 1; i <= horizontal; i++)
                {
                    if (myValues.GetValue(headerIndex, i) != null)
                    {
                        columnName = myValues.GetValue(headerIndex, i).ToString().Trim();
                        if (columnName.Length > 0)
                        {
                            var column = Cleaning(columnName);
                            if (dt.Columns.Contains(column))  //동일한폴더존재시
                            {
                                if (column == "Comment")
                                {
                                    dt.Columns.Add(new DataColumn(column + "_" + i));
                                }
                                else {
                                    dt.Columns.Add(new DataColumn("Ext" + extIndex++));
                                }
                            }
                            else
                            {
                                dt.Columns.Add(new DataColumn(column));
                            }
                        }
                    }

                }

                // Get the row information
                //for (int a = (headerIndex + 1); a <= vertical; a++)
                //{
                //    object[] poop = new object[horizontal];
                //    for (int b = 1; b <= horizontal; b++)
                //    {
                //        poop[b - 1] = myValues.GetValue(a, b);
                //    }
                //    DataRow row = dt.NewRow();
                //    row.ItemArray = poop;
                //    dt.Rows.Add(row);
                //}

                for (int a = (headerIndex + 1); a <= vertical; a++)
                {
                    object[] poop = new object[horizontal];
                    for (int b = 1; b <= horizontal; b++)
                    {
                        poop[b - 1] = myValues.GetValue(a, b);
                    }
                    DataRow row = dt.NewRow();
                    row.ItemArray = poop;
                    dt.Rows.Add(row);
                }

                ds.Tables.Add(dt);

                //File info insert
                DataRow pRow = dataTable.NewRow();
                dataTable.Rows.Add(pRow);

                workbook.Close(true, missing, missing);
                app.Quit();

                releaseObject(worksheet);
                releaseObject(workbook);
                releaseObject(app);

            }
            catch (Exception ex)
            {
                throw ex;
            }

            return ds;
        }

        //DB저장()
        /// <summary>
        /// File의 경로를 읽고
        /// 한 file당 2개의 Table에 행추가
        /// 행만큼 Loop예외는 내용기록
        /// 중간에 문제 발생시 Report & Confirm => 처리/중단여부 결정처리
        /// 완료되면 완료결과 Message
        /// </summary>
        /// <param name="fileName"></param>
        /// <returns></returns>
        public DataSet ExcelToDB(string fileName)
        {

            string sqlDatabase = "Data Source=ISAAC-PC\\SQLEXPRESS;Initial Catalog=KSS.Local;Persist Security Info=True;User ID=sa;Password=1234";
            string classNamespace = "MyWeb.Models.Excel";

            var classList = Assembly.GetExecutingAssembly().GetTypes().Where(t => String.Equals(t.Namespace, classNamespace, StringComparison.Ordinal)).ToList();
            var filesToProcess = fileName;
            int extIndex = 1;
            int idx = 0;

            using (var dbContext = new DataContext(sqlDatabase))
            {
                var classQuery = (from tmpClass in classList
                                  where !tmpClass.Name.StartsWith("I")
                                  select tmpClass);

                foreach (var tmpClass in classQuery)
                {
                    var sqlTable = dbContext.GetTable(tmpClass);
                    var fileTable = dbContext.GetTable(typeof(MyWeb.Models.FileImport));

                    MyWeb.Models.FileImport fileInfo = new Models.FileImport()
                    {
                        Name = fileName,
                        Creator = "",
                        Extend = "",
                        Path = "",
                        Reason = "",
                        Remark = "",
                        ExtName = "",
                        Result = ""
                    };
                    fileTable.InsertOnSubmit(fileInfo);
                    dbContext.SubmitChanges();

                    try
                    {


                        // How many records exist so far?
                        var countQuery = (from object o in sqlTable select o);
                        var countQueryFile = (from object o in fileTable select o);

                        // Only process the table when no records existed yet?

                        if (1 == 1)   //if (!countQuery.Any())
                        {
                            var attributes = tmpClass.GetCustomAttributes(typeof(TableAttribute), true);    //Class의 속성을 가져온다

                            if (attributes.Any())
                            {
                                var tableName = ((TableAttribute)attributes[0]).Name;

                                using (var myDataSet = OfficeExcelTODataSet(fileName))
                                {
                                    //1. try catch해서 문제있으면 fileInfo에 저장후 마침


                                    // The data table will have the same name
                                    using (var dataTable = myDataSet.Tables[1])
                                    {
                                        //Copy Datatable
                                        var currDataTable = dataTable.Copy();
                                        if (!currDataTable.Columns.Contains("Reason"))
                                        {
                                            currDataTable.Columns.Add("Reason");
                                        }
                                        for (int i = 0; i < 30; i++)
                                        {
                                            if (currDataTable.Columns.Contains("Ext" + (i + 1)))
                                            {
                                                continue;
                                            }
                                            else
                                            {
                                                currDataTable.Columns.Add("Ext" + (i + 1));
                                            }                                            
                                        }

                                        // We need to create a new object of type tmpClass for each row and populate it.
                                        foreach (DataRow row in currDataTable.Rows)
                                        {
                                            extIndex = 1;
                                            if (row["Name"] == null || row["Name"].ToString().Trim() == "") continue;

                                            // Using Reflection to create this object and fill in properties.
                                            var instance = Activator.CreateInstance(tmpClass);
                                            var properties = tmpClass.GetProperties();

                                            var propertyQuery = (from property in properties
                                                                 where property.CanWrite
                                                                 select property);

                                            //Table mapping(Extended)...<!------------------------------
                                            //dataColumn만큼 loop, 속성과 일치하는 column이 없으면 Ext

                                            foreach (DataColumn col in dataTable.Columns)
                                            {
                                                //컬럼명이 엑셀에 존재하지 않으면
                                                if (propertyQuery.ToArray().Where(e => e.Name == col.ColumnName).Count() < 1)
                                                {
                                                    //Ext컬럼에 추가
                                                    string currentExtColumnName = "Ext" + (extIndex++);
                                                    row[currentExtColumnName] = row[col.ColumnName];
                                                    row["Reason"] += "|" + col.ColumnName;
                                                }
                                            }
                                            //---------------------------------------------------------->


                                            foreach (PropertyInfo property in propertyQuery)
                                            {
                                                // Grab the Linq to Sql data attributes.
                                                var dbProperty = property.GetCustomAttribute
                                                    (typeof(ColumnAttribute), false) as ColumnAttribute;

                                                if (dbProperty == null) continue;


                                                // Make sure that this column exists in the data we received from the XLS spreadsheet
                                                if (currDataTable.Columns.Contains(property.Name))
                                                {
                                                    // Grab the value.  We need to account for DBNull first.
                                                    var val = row[property.Name];
                                                    if (val == DBNull.Value)
                                                    {
                                                        val = null;
                                                    }

                                                    // We need a bunch of special processing for null.  Empty cells are returned
                                                    // instead of empty strings for example.
                                                    if (val == null)
                                                    {
                                                        // DateTime should get processed specially.
                                                        if ((property.PropertyType == typeof(DateTime)) ||
                                                            (property.PropertyType == typeof(DateTime?)))
                                                        {
                                                            DateTime? nullableDate = null;
                                                            property.SetValue(instance, nullableDate);
                                                        }
                                                        else if (!dbProperty.CanBeNull)
                                                        {
                                                            // If the value should not be null we need to create the default instance
                                                            // of that class. (e.g. int = 0, etc.)  Strings do not have a constructor
                                                            // that's usable this way so strings are a special check.
                                                            if (property.PropertyType == typeof(string))
                                                            {
                                                                property.SetValue(instance, string.Empty);
                                                            }
                                                            else {
                                                                var tmpVal = Activator.CreateInstance(property.PropertyType).GetType();
                                                                property.SetValue(instance, Activator.CreateInstance(tmpVal));
                                                            }
                                                        }
                                                        else {
                                                            // To here, we have a valid null value and it's not a DateTime.
                                                            property.SetValue(instance, null);
                                                        }
                                                    }
                                                    else if ((dbProperty.DbType.StartsWith("nvarchar") &&
                                                             (!string.IsNullOrEmpty(val.ToString()))))
                                                    {

                                                        // This block of code assumes that the DbType is specified.  If it is,
                                                        // we can account for string truncation here.
                                                        var sLength = dbProperty.DbType.Substring(("nvarchar(").Length);
                                                        sLength = sLength.Substring(0, sLength.Length - 1);
                                                        var iLength = Int32.Parse(sLength);
                                                        var newVal = val.ToString();
                                                        newVal = newVal.Substring(0, Math.Min(iLength, newVal.Length));

                                                        // We've truncated to here. If we are handling char type, a string
                                                        // cannot be converted to char.  We need to handle this now. Only
                                                        // handle for 1 length, otherwise we'll let the app throw an error.
                                                        if ((property.PropertyType == typeof(char)) &&
                                                            (newVal.Length == 1))
                                                        {
                                                            property.SetValue(instance, newVal[0]);
                                                        }
                                                        else {
                                                            // Set the truncated string
                                                            property.SetValue(instance, newVal);
                                                        }
                                                    }
                                                    else if (val.GetType() != property.PropertyType)
                                                    {

                                                        // To here, the resulting types are different somehow. We need to
                                                        // do some conversions on the data.  Checking for DateTime.
                                                        if ((val.GetType() == typeof(DateTime)) ||
                                                            (val.GetType() == typeof(DateTime?)))
                                                        {

                                                            // nullable fields don't convert otherwise.
                                                            DateTime? nullableDate = (DateTime)val;
                                                            property.SetValue(instance, nullableDate);
                                                        }
                                                        else if ((property.PropertyType == typeof(DateTime)) ||
                                                                 (property.PropertyType == typeof(DateTime?)))
                                                        {

                                                            // A number of times the record comes back as a string instead.
                                                            var newVal = val.ToString();
                                                            var nullableDate = (string.IsNullOrWhiteSpace
                                                               (newVal) ? (DateTime?)null : DateTime.Parse(newVal));
                                                            property.SetValue(instance, nullableDate);
                                                        }
                                                        else {
                                                            // To here we have a different type and need to convert. It's not
                                                            // a DateTime, and it's not a null value which was handled already.
                                                            var pType = property.PropertyType;

                                                            // We can't take "Int? 3" and
                                                            // put it into "Int" field. Must convert.
                                                            if ((property.PropertyType.IsGenericType) &&
                                                                (property.PropertyType.GetGenericTypeDefinition().
                                                                   Equals(typeof(Nullable<>))))
                                                            {
                                                                pType = Nullable.GetUnderlyingType(property.PropertyType);
                                                            }

                                                            // Finally change the type and set the value.
                                                            var newProp = Convert.ChangeType(val, pType);
                                                            property.SetValue(instance, newProp);
                                                        }
                                                    }
                                                    else {
                                                        // To here the types match and the value isn't null
                                                        property.SetValue(instance, val);
                                                    }

                                                } // dbColumn exists
                                                else
                                                {
                                                    if ("FileID".Contains(property.Name))   //파일ID가져온다
                                                    {
                                                        property.SetValue(instance, fileInfo.ID);
                                                    }
                                                    else {
                                                        //Ext add
                                                        try
                                                        {
                                                            var val = row[property.Name];
                                                            if (val == DBNull.Value)
                                                            {
                                                                val = null;
                                                            }
                                                            if (property.Name == "Ext" + extIndex++)
                                                            {
                                                                property.SetValue(instance, val != null ? val.ToString().Trim() : "");
                                                            }
                                                            property.SetValue(instance, val != null ? val.ToString().Trim() : "");

                                                        }
                                                        catch (Exception)
                                                        {

                                                        }
                                                    }
                                                }

                                            } // property loop

                                            // This instance can be inserted if needed.
                                            //var inst = (MyWeb.Models.Excel.ValuationModels)instance;
                                            //if (inst.Name != null)
                                            sqlTable.InsertOnSubmit(instance);


                                        } // DataRow loop

                                        // Submit changes.
                                        dbContext.SubmitChanges();

                                    } // using DataTable

                                    return myDataSet;
                                } // using DataSet

                            } // Attributes exist

                        } // No records were preexisting in the database table.
                    }
                    catch (Exception ex)
                    {
                        string msg = ex.Message;                       
                        
                        fileInfo.Result = "E";
                        fileInfo.Reason = msg;
                        dbContext.SubmitChanges();
                    }

                } // class loop

            } // using DataContext

            return null;
        }

        //Excel to DataSet변환(OleDB방식)
        public DataSet OleDBExcelToDataSet(string fileName)
        {
            DataSet ds = new DataSet();
            System.Data.DataTable dtFile = new System.Data.DataTable("ExcelImportFileInfo");
            System.Data.DataTable dt = new System.Data.DataTable("ExcelImportedData");
            ds.Tables.Add(dtFile);
            ds.Tables.Add(dt);

            string sqlDatabase = "Data Source=ISAAC-PC\\SQLEXPRESS;Initial Catalog=KSS.Local;Persist Security Info=True;User ID=sa;Password=1234";
            string xlsxFile = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0}; Extended Properties='Excel 12.0;HDR=YES';";   //엑셀 2007 
            string classNamespace = "MyWeb.Models.Excel";
            string connectionString = xlsxFile;

            var missing = System.Reflection.Missing.Value;
            Application app = new Microsoft.Office.Interop.Excel.Application();

            //1번째 Worksheet찾기
            Workbook workbook = app.Workbooks.Open(fileName, false, true, missing, missing, missing, true, XlPlatform.xlWindows, '\t', false, false, 0, false, true, 0);
            Worksheet worksheet = workbook.Worksheets[1] as Worksheet; worksheet = (Worksheet)workbook.Worksheets.get_Item(1);
            var classList = Assembly.GetExecutingAssembly().GetTypes().Where(t => String.Equals(t.Namespace, classNamespace, StringComparison.Ordinal)).ToList();

            var filesToProcess = fileName;

            using (var dbContext = new DataContext(sqlDatabase))
            {
                // Skip interfaces. There is likely a better way
                var classQuery = (from tmpClass in classList
                                  where !tmpClass.Name.StartsWith("I")
                                  select tmpClass);

                foreach (var tmpClass in classQuery)
                {
                    var sqlTable = dbContext.GetTable(tmpClass);

                    // How many records exist so far?
                    var countQuery = (from object o in sqlTable select o);

                    // Only process the table when no records existed yet?

                    if (!countQuery.Any())
                    {
                        var attributes = tmpClass.GetCustomAttributes(typeof(TableAttribute), true);    //Class의 속성을 가져온다

                        if (attributes.Any())
                        {
                            var tableName = ((TableAttribute)attributes[0]).Name;
                            string sheetName = worksheet.Name;

                            //Class명과 일치하는 파일명
                            var file = Path.GetFileNameWithoutExtension(filesToProcess).Equals(tableName,
                            StringComparison.CurrentCultureIgnoreCase);

                            //Table mapping...<!------------------------------
                            /*
                            List<string> cols = new List<string> { "", "" };

                            //존재하면 추가
                            cols.Add("");
                            */
                            //------------------------------------------------>

                            using (var dataAdapter = new OleDbDataAdapter
                            ("SELECT * FROM [" + sheetName + "]", string.Format(connectionString, file)))
                            {
                                using (var myDataSet = new DataSet())
                                {
                                    dataAdapter.Fill(myDataSet, tableName);

                                    // The data table will have the same name
                                    using (var dataTable = myDataSet.Tables[tableName])
                                    {
                                        // We need to create a new object of type tmpClass for each row and populate it.
                                        foreach (DataRow row in dataTable.Rows)
                                        {
                                            // Using Reflection to create this object and fill in properties.
                                            var instance = Activator.CreateInstance(tmpClass);
                                            var properties = tmpClass.GetProperties();

                                            var propertyQuery = (from property in properties
                                                                 where property.CanWrite
                                                                 select property);

                                            foreach (PropertyInfo property in propertyQuery)
                                            {
                                                // Grab the Linq to Sql data attributes.
                                                var dbProperty = property.GetCustomAttribute
                                                    (typeof(ColumnAttribute), false) as ColumnAttribute;

                                                if (dbProperty == null) continue;

                                                // Make sure that this column exists in the data we received from the XLS spreadsheet
                                                if (dataTable.Columns.Contains(dbProperty.Name))
                                                {

                                                    // Grab the value.  We need to account for DBNull first.
                                                    var val = row[dbProperty.Name];
                                                    if (val == DBNull.Value)
                                                    {
                                                        val = null;
                                                    }

                                                    // We need a bunch of special processing for null.  Empty cells are returned
                                                    // instead of empty strings for example.
                                                    if (val == null)
                                                    {

                                                        // DateTime should get processed specially.
                                                        if ((property.PropertyType == typeof(DateTime)) ||
                                                            (property.PropertyType == typeof(DateTime?)))
                                                        {
                                                            DateTime? nullableDate = null;
                                                            property.SetValue(instance, nullableDate);
                                                        }
                                                        else if (!dbProperty.CanBeNull)
                                                        {

                                                            // If the value should not be null we need to create the default instance
                                                            // of that class. (e.g. int = 0, etc.)  Strings do not have a constructor
                                                            // that's usable this way so strings are a special check.
                                                            if (property.PropertyType == typeof(string))
                                                            {
                                                                property.SetValue(instance, string.Empty);
                                                            }
                                                            else {
                                                                var tmpVal = Activator.CreateInstance(property.PropertyType).GetType();
                                                                property.SetValue(instance, Activator.CreateInstance(tmpVal));
                                                            }
                                                        }
                                                        else {
                                                            // To here, we have a valid null value and it's not a DateTime.
                                                            property.SetValue(instance, null);
                                                        }
                                                    }
                                                    else if ((dbProperty.DbType.StartsWith("nvarchar") &&
                                                             (!string.IsNullOrEmpty(val.ToString()))))
                                                    {

                                                        // This block of code assumes that the DbType is specified.  If it is,
                                                        // we can account for string truncation here.
                                                        var sLength = dbProperty.DbType.Substring(("nvarchar(").Length);
                                                        sLength = sLength.Substring(0, sLength.Length - 1);
                                                        var iLength = Int32.Parse(sLength);
                                                        var newVal = val.ToString();
                                                        newVal = newVal.Substring(0, Math.Min(iLength, newVal.Length));

                                                        // We've truncated to here. If we are handling char type, a string
                                                        // cannot be converted to char.  We need to handle this now. Only
                                                        // handle for 1 length, otherwise we'll let the app throw an error.
                                                        if ((property.PropertyType == typeof(char)) &&
                                                            (newVal.Length == 1))
                                                        {
                                                            property.SetValue(instance, newVal[0]);
                                                        }
                                                        else {
                                                            // Set the truncated string
                                                            property.SetValue(instance, newVal);
                                                        }
                                                    }
                                                    else if (val.GetType() != property.PropertyType)
                                                    {

                                                        // To here, the resulting types are different somehow. We need to
                                                        // do some conversions on the data.  Checking for DateTime.
                                                        if ((val.GetType() == typeof(DateTime)) ||
                                                            (val.GetType() == typeof(DateTime?)))
                                                        {

                                                            // nullable fields don't convert otherwise.
                                                            DateTime? nullableDate = (DateTime)val;
                                                            property.SetValue(instance, nullableDate);
                                                        }
                                                        else if ((property.PropertyType == typeof(DateTime)) ||
                                                                 (property.PropertyType == typeof(DateTime?)))
                                                        {

                                                            // A number of times the record comes back as a string instead.
                                                            var newVal = val.ToString();
                                                            var nullableDate = (string.IsNullOrWhiteSpace
                                                               (newVal) ? (DateTime?)null : DateTime.Parse(newVal));
                                                            property.SetValue(instance, nullableDate);
                                                        }
                                                        else {
                                                            // To here we have a different type and need to convert. It's not
                                                            // a DateTime, and it's not a null value which was handled already.
                                                            var pType = property.PropertyType;

                                                            // We can't take "Int? 3" and
                                                            // put it into "Int" field. Must convert.
                                                            if ((property.PropertyType.IsGenericType) &&
                                                                (property.PropertyType.GetGenericTypeDefinition().
                                                                   Equals(typeof(Nullable<>))))
                                                            {
                                                                pType = Nullable.GetUnderlyingType(property.PropertyType);
                                                            }

                                                            // Finally change the type and set the value.
                                                            var newProp = Convert.ChangeType(val, pType);
                                                            property.SetValue(instance, newProp);
                                                        }
                                                    }
                                                    else {
                                                        // To here the types match and the value isn't null
                                                        property.SetValue(instance, val);
                                                    }

                                                } // dbColumn exists

                                            } // property loop

                                            // This instance can be inserted if needed.
                                            sqlTable.InsertOnSubmit(instance);

                                        } // DataRow loop

                                        // Submit changes.
                                        dbContext.SubmitChanges();

                                    } // using DataTable

                                } // using DataSet

                            } // Using DataAdapter

                        } // Attributes exist

                    } // No records were preexisting in the database table.

                } // class loop

            } // using DataContext

            return ds;
        }

        private void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                //MessageBox.Show("Unable to release the Object " + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }
    }
}