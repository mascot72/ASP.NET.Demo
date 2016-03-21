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
using MyWeb.Models;
using System.Collections;

namespace MyWeb.Processor
{
    public class ExcelConverter
    {
        //Members declare
        string sqlDatabase = "Data Source=ISAAC-PC\\SQLEXPRESS;Initial Catalog=KSS.Local;Persist Security Info=True;User ID=sa;Password=1234";
        int extCount = 50;  //추가컬럼갯수

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
            string headerList = "EID,INVENNO,SGNO,TID,NAME";
            int headerCellCount = headerList.Split(spliter).Length;
            int headerIndex = 1;

            DataSet ds = new DataSet();
            System.Data.DataTable dataTable = new System.Data.DataTable("ExcelImportFileInfo");
            ds.Tables.Add(dataTable);

            var missing = System.Reflection.Missing.Value;

            Application app = new Microsoft.Office.Interop.Excel.Application();

            //파일읽기
            Workbook workbook = app.Workbooks.Open(fileName, false, true, missing, missing, missing, true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, '\t', false, false, 0, false, true, 0);
            Worksheet worksheet = workbook.Worksheets[1] as Microsoft.Office.Interop.Excel.Worksheet; worksheet = (Microsoft.Office.Interop.Excel.Worksheet)workbook.Worksheets.get_Item(1);

            Range xlRange = worksheet.UsedRange;
            Array myValues = (Array)xlRange.Cells.Value2;

            try
            {

                int vertical = myValues.GetLength(0);
                int horizontal = myValues.GetLength(1);
                string columnName = "";
                System.Data.DataTable dt = new System.Data.DataTable("ExcelImportedData");

                // must start with index = 1                
                // Header의 Row위치를 찾는다(Header가 처음에 오지 않는 예외가 있음으로)
                for (int a = 1; a <= vertical; a++)
                {
                    for (int i = 1; i <= horizontal; i++)
                    {
                        //일치하는 경우
                        if (myValues.GetValue(a, i) != null)
                        {
                            columnName = Cleaning(myValues.GetValue(a, i).ToString().Trim());
                            if (columnName.Length > 0 && headerList.IndexOf(columnName.ToUpper()) > -1)
                            {
                                findedCount++;
                            }
                        }

                    }
                    if (findedCount >= 3) //필수항목이 4개 이상이면 Header이다
                    {
                        headerIndex = a;
                        break;
                    }
                }

                //property declare
                string classNamespace = "MyWeb.Models.Excel";
                var classList = Assembly.GetExecutingAssembly().GetTypes().Where(t => String.Equals(t.Namespace, classNamespace, StringComparison.Ordinal)).ToList();
                var classQuery = (from tmpClass in classList
                                  where !tmpClass.Name.StartsWith("I")
                                  select tmpClass);
                foreach (var tmpClass in classQuery)
                {
                    var attributes = tmpClass.GetCustomAttributes(typeof(TableAttribute), true);    //Class의 속성을 가져온다
                    var properties = tmpClass.GetProperties();
                    var propertyQuery = (from property in properties
                                         where property.CanWrite
                                         select property);

                    // Generate DataColumn(Cleanning)
                    for (int i = 1; i <= horizontal; i++)   //Excel의 컬럼만큼 Loop
                    {
                        if (myValues.GetValue(headerIndex, i) != null)
                        {
                            columnName = myValues.GetValue(headerIndex, i).ToString().Trim();
                            if (columnName.Length > 0)
                            {
                                var column = Cleaning(columnName);

                                var property = propertyQuery.Where(e => e.Name.ToUpper() == column.ToUpper()).FirstOrDefault();

                                int idxColumn = dt.Columns.IndexOf(column);
                                if (idxColumn > -1 && dt.Columns[idxColumn].ColumnName == column)  //컬럼이 존재시(중복일 경우)
                                {
                                    if (column == "Comment")
                                    {
                                        if (dt.Columns.Contains("Comment_1"))  //컬럼이 존재시(중복일 경우)
                                            throw new ApplicationException(string.Format("중복 Column({0})이 발생하여서 처리하지 못합니다", columnName));
                                        else
                                            dt.Columns.Add(new DataColumn(column + "_1", property.PropertyType));
                                    }
                                    else if ("Profit%".ToUpper().Contains(columnName.ToUpper()))
                                    {
                                        if (dt.Columns.Contains("ProfitPercent"))  //컬럼이 존재시(중복일 경우)
                                            throw new ApplicationException(string.Format("중복 Column({0})이 발생하여서 처리하지 못합니다", columnName));
                                        else
                                            dt.Columns.Add(new DataColumn("ProfitPercent", typeof(float)));
                                    }
                                    else {
                                        throw new ApplicationException(string.Format("중복 Column({0})이 발생하여서 처리하지 못합니다", columnName));
                                    }
                                }
                                else
                                {
                                    if (property != null)   //속성과 일치하는 Excel컬럼 일 때
                                    {
                                        dt.Columns.Add(new DataColumn(column, property.PropertyType));
                                    }
                                    else
                                    {
                                        dt.Columns.Add(new DataColumn(column, typeof(object))); //속성에 없는 Excel컬럼이 나올때.
                                    }
                                }
                            }
                        }

                    }
                    //dt.Columns.Add(new DataColumn("Reason", typeof(string)));   //필수컴럼중 추가
                }

                if (dt.Columns.Count < horizontal)
                    throw new ApplicationException("Header에 공백Column이 나와서 처리 할 수 없습니다!");

                // Excel Data 복사
                for (int a = (headerIndex + 1); a <= vertical; a++) //Excel행만큼 Loop
                {
                    DataRow row = dt.NewRow();
                    string message = "";
                    for (int b = 1; b <= horizontal; b++)   //Excel컬럼만큼 Loop
                    {
                        if (myValues.GetValue(a, b) != null && myValues.GetValue(a, b).ToString().Trim() != string.Empty)
                        {
                            try
                            {
                                if (dt.Columns[b - 1].DataType == typeof(DateTime))
                                {
                                    row[b - 1] = DateTime.FromOADate((double)myValues.GetValue(a, b));
                                }
                                else
                                {
                                    row[b - 1] = myValues.GetValue(a, b).ToString().Trim();
                                }
                            }
                            catch (Exception ex)
                            {
                                message += (message != null && message.Length > 0 ? "|" : "") + "ERROR{" + ex.Message + ", index(" + a + ", " + b + ") orginalValue=" + myValues.GetValue(a, b) + "}";
                            }
                        }
                    }
                    if (message.Length > 0)
                        throw new ApplicationException(message);

                    dt.Rows.Add(row);
                }

                ds.Tables.Add(dt);

                //File info insert
                DataRow pRow = dataTable.NewRow();
                dataTable.Rows.Add(pRow);
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                workbook.Close(true, missing, missing);
                app.Quit();

                releaseObject(worksheet);
                releaseObject(workbook);
                releaseObject(app);
            }

            return ds;
        }

        //파일입력정보
        public IEnumerable<FileImport> GetFileTable()
        {
            try
            {
                IEnumerable<MyWeb.Models.FileImport> fileTable = null;
                using (var dbContext = new DataContext(sqlDatabase))
                {
                    fileTable = dbContext.GetTable<MyWeb.Models.FileImport>().ToList();
                    //.ExecuteQuery<MyWeb.Models.FileImport>("SELECT ID, Path, Name, Extname, Result, Reason, Remark, Extend, CreateDate, Creator, Size FROM FILE_IMPORT_INFO");

                }
                return fileTable;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        //확장정보 목록조회
        public IEnumerable<ModelExtendColumn> GetModelExtendList()
        {
            try
            {
                IEnumerable<ModelExtendColumn> fileTable = null;
                using (var dbContext = new DataContext(sqlDatabase))
                {
                    fileTable = dbContext.GetTable<ModelExtendColumn>().ToList();
                }
                return fileTable;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        //Upload후처리(DB Update)
        public int UpdateAfter()
        {
            try
            {
                int result = 0;
                using (var dbContext = new DataContext(sqlDatabase))
                {
                    result = dbContext.ExecuteCommand(@"
update VALU_EXCEL
set BuyDate = null
where convert(varchar(10), BuyDate, 126) = '1900-01-01'

update VALU_EXCEL
set SellDate = null
where convert(varchar(10), SellDate, 126) = '1900-01-01'

update VALU_EXCEL
set Date = null
where convert(varchar(10), Date, 126) = '1900-01-01'");
                }
                return result;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        //확장정보 등록
        public ModelExtendColumn AddModelExtend(ModelExtendColumn column)
        {
            try
            {
                ModelExtendColumn result = null;
                using (var dbContext = new DataContext(sqlDatabase))
                {
                    if (GetModelExtend(column.Name).Count() == 0)
                    {
                        dbContext.ExecuteCommand("INSERT INTO VALU_EXCEL_EXT(ID, Name) VALUES({0}, {1})", column.ID, column.Name);
                    }
                    result = dbContext.GetTable<ModelExtendColumn>().Where(e => e.ID == column.ID).First();
                }
                return result;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        //모델확장컬럼정보
        public IEnumerable<ModelExtendColumn> GetModelExtend(string name)
        {
            try
            {
                IEnumerable<MyWeb.Models.ModelExtendColumn> fileTable = null;
                using (var dbContext = new DataContext(sqlDatabase))
                {
                    if (name != null)
                        fileTable = dbContext.GetTable<MyWeb.Models.ModelExtendColumn>().Where(e => e.Name == name).ToList();
                    else
                        fileTable = dbContext.GetTable<MyWeb.Models.ModelExtendColumn>().ToList();
                }
                return fileTable;
            }
            catch (Exception ex)
            {
                throw ex;
            }
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
        public DataSet ExcelToDB(string fileName, int[] success)
        {
            string classNamespace = "MyWeb.Models.Excel";
            var extList = GetModelExtendList();

            if (success == null)
            {
                int[] aa = { 0, 0 };
                success = aa;
            }

            var classList = Assembly.GetExecutingAssembly().GetTypes().Where(t => String.Equals(t.Namespace, classNamespace, StringComparison.Ordinal)).ToList();
            var filesToProcess = fileName;
            int idx = 0;
            int successFileCount = 0;

            using (var dbContext = new DataContext(sqlDatabase))
            {
                var classQuery = (from tmpClass in classList
                                  where !tmpClass.Name.StartsWith("I")
                                  select tmpClass);

                foreach (var tmpClass in classQuery)
                {
                    var sqlTable = dbContext.GetTable(tmpClass);
                    var fileTable = dbContext.GetTable(typeof(MyWeb.Models.FileImport));

                    FileInfo files = new System.IO.FileInfo(fileName);

                    if (!files.Exists)
                    {
                        break;
                    }

                    FileImport fileInfo = new FileImport()
                    {
                        Name = files.Name,
                        Creator = "",
                        CreateDate = DateTime.Now,
                        Extend = "",
                        Path = files.DirectoryName,
                        Reason = "",
                        Result = "P",
                        Remark = "file Last write time at(" + files.LastWriteTime.ToString() + ")",
                        ExtName = Cleaning(files.Extension),
                        Size = files.Length
                    };
                    fileTable.InsertOnSubmit(fileInfo);
                    dbContext.SubmitChanges();

                    try
                    {
                        var countQuery = (from object o in sqlTable select o);
                        var countQueryFile = (from object o in fileTable select o);

                        if (1 == 1)   //if (!countQuery.Any())
                        {
                            var attributes = tmpClass.GetCustomAttributes(typeof(TableAttribute), true);    //Class의 속성을 가져온다

                            if (attributes.Any())
                            {

                                var instanceMst = Activator.CreateInstance(tmpClass);
                                var propertiesMst = tmpClass.GetProperties();
                                var propertyQueryMst = (from property in propertiesMst
                                                        where property.CanWrite
                                                        select property);

                                var tableName = ((TableAttribute)attributes[0]).Name;

                                using (var myDataSet = OfficeExcelTODataSet(fileName))
                                {
                                    //1. try catch해서 문제있으면 fileInfo에 저장후 마침

                                    // The data table will have the same name
                                    using (var dataTable = myDataSet.Tables[1])
                                    {
                                        //<==================================================================
                                        //컬럼 정의
                                        //1. 추가컬럼 등록
                                        //Copy Datatable
                                        var currDataTable = dataTable.Copy();
                                        if (!currDataTable.Columns.Contains("Reason"))
                                        {
                                            currDataTable.Columns.Add("Reason");
                                        }
                                        
                                        for (int i = 0; i < extCount; i++)
                                        {
                                            string extId = "Ext" + (i + 1);
                                            int idxColumn = currDataTable.Columns.IndexOf(extId);
                                            if (idxColumn > -1 && currDataTable.Columns[idxColumn].ColumnName == extId)  //컬럼이 존재시(중복일 경우)
                                            {
                                                continue;
                                            }
                                            else
                                            {
                                                currDataTable.Columns.Add(extId);
                                            }
                                        }

                                        //2. 컬럼 예외처리------------------------------------------------------------------------------ 
                                        foreach (DataColumn col in currDataTable.Columns)
                                        {
                                            //엑셀컬럼중 Model Property에 존재하지 않으면
                                            if (propertyQueryMst.Count(e => e.Name.ToUpper() == col.ColumnName.ToUpper()) < 1)
                                            {
                                                if ("Profit%".ToUpper().Contains(col.ColumnName.ToUpper()) == false)
                                                {
                                                    //확장정보가 기존에 존재하는지 확인                                                    
                                                    var entity = extList.Where(e => e.Name.ToUpper() == col.ColumnName.ToUpper()).SingleOrDefault();
                                                    if (entity == null && !col.ColumnName.StartsWith("Ext"))
                                                    {
                                                        if (extList.Count() >= extCount) throw new ApplicationException("Ext컬럼이 " + extCount + "개가 넘어서 현재파일은 Rollback합니다");

                                                        //확장정보추가
                                                        var addExt = AddModelExtend(new ModelExtendColumn()
                                                        {
                                                            ID = "Ext" + (extList.Count() + 1),
                                                            Name = col.ColumnName
                                                        });

                                                        fileInfo.Extend += (fileInfo.Extend != "" ? "|" : "") + col.ColumnName + "(" + addExt.ID + ")";
                                                        extList = GetModelExtendList(); //Reload from DB
                                                    }

                                                }
                                            }
                                        }//------------------------------------------------------------------------------------------------                               
                                        //==================================================================>

                                        foreach (DataRow row in currDataTable.Rows)
                                        {
                                            //each row reset Ext? Columns
                                            if (row["Name"] == null || row["Name"].ToString().Trim() == "")
                                            {
                                                //throw new ApplicationException("Name컬럼이 존재하지 않습니다");
                                                continue;   //병합된 컬럼이 존재시 이Row만 제외하고 통과해야 한다
                                            }

                                            //속성
                                            var instance = Activator.CreateInstance(tmpClass);
                                            var properties = tmpClass.GetProperties();
                                            var propertyQuery = (from property in properties
                                                                 where property.CanWrite
                                                                 select property);

                                            //Table mapping(Extended)...<!------------------------------                                            
                                            //dataColumn만큼 loop, if property not matched column then Ext++
                                            foreach (DataColumn col in currDataTable.Columns)
                                            {
                                                //엑셀컬럼중 Model Property에 존재하지 않으면
                                                if (propertyQuery.Count(e => e.Name.ToUpper() == col.ColumnName.ToUpper()) < 1)
                                                {
                                                    if ("Profit%".ToUpper().Contains(col.ColumnName.ToUpper()))
                                                    {
                                                        row["ProfitPercent"] = row[col.ColumnName];
                                                    }
                                                    else
                                                    {
                                                        //확장정보가 기존에 존재하는지 확인
                                                        var entity = extList.Where(e => e.Name.ToUpper() == col.ColumnName.ToUpper()).SingleOrDefault();
                                                        if (entity != null)
                                                        {
                                                            row[entity.ID] = row[col.ColumnName];
                                                        }

                                                    }
                                                }
                                            }
                                            //---------------------------------------------------------->

                                            foreach (PropertyInfo property in propertyQuery)
                                            {
                                                // Grab the Linq to Sql data attributes.
                                                var dbProperty = property.GetCustomAttribute
                                                    (typeof(ColumnAttribute), false) as ColumnAttribute;

                                                if (dbProperty == null) continue;

                                                int idxColumn = currDataTable.Columns.IndexOf(property.Name);
                                                if (idxColumn > -1 && currDataTable.Columns[idxColumn].ColumnName == property.Name)  //컬럼이 존재시
                                                {
                                                    var val = row[property.Name];
                                                    if (val == DBNull.Value)
                                                    {
                                                        val = null;
                                                    }

                                                    if (val == null)
                                                    {
                                                        if ((property.PropertyType == typeof(DateTime)) ||
                                                            (property.PropertyType == typeof(DateTime?)))
                                                        {
                                                            //DateTime? nullableDate = null;

                                                            //min DateTime
                                                            DateTime nullableDate = new DateTime(1900, 1, 1);

                                                            property.SetValue(instance, nullableDate);
                                                        }
                                                        else if (!dbProperty.CanBeNull)
                                                        {
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
                                                            property.SetValue(instance, null);
                                                        }
                                                    }
                                                    else if ((dbProperty.DbType.StartsWith("nvarchar") &&
                                                             (!string.IsNullOrEmpty(val.ToString()))))
                                                    {

                                                        var sLength = dbProperty.DbType.Substring(("nvarchar(").Length);
                                                        sLength = sLength.Substring(0, sLength.Length - 1);
                                                        var iLength = Int32.Parse(sLength);
                                                        var newVal = val.ToString();
                                                        newVal = newVal.Substring(0, Math.Min(iLength, newVal.Length));

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

                                                        if ((val.GetType() == typeof(DateTime)) ||
                                                            (val.GetType() == typeof(DateTime?)))
                                                        {
                                                            //nullable DateTime
                                                            //DateTime? nullableDate = (DateTime)val;                                                            

                                                            //min DateTime
                                                            DateTime nullableDate = new DateTime(1900, 1, 1);

                                                            property.SetValue(instance, nullableDate);
                                                        }
                                                        else if ((property.PropertyType == typeof(DateTime)) ||
                                                                 (property.PropertyType == typeof(DateTime?)))
                                                        {

                                                            var newVal = val.ToString();
                                                            var nullableDate = (string.IsNullOrWhiteSpace
                                                               (newVal) ? (DateTime?)null : DateTime.Parse(newVal));
                                                            property.SetValue(instance, nullableDate);
                                                        }
                                                        else {
                                                            var pType = property.PropertyType;

                                                            if ((property.PropertyType.IsGenericType) &&
                                                                (property.PropertyType.GetGenericTypeDefinition().
                                                                   Equals(typeof(Nullable<>))))
                                                            {
                                                                pType = Nullable.GetUnderlyingType(property.PropertyType);
                                                            }

                                                            var newProp = Convert.ChangeType(val, pType);
                                                            property.SetValue(instance, newProp);
                                                        }
                                                    }
                                                    else {
                                                        property.SetValue(instance, val);
                                                    }

                                                } // dbColumn exists
                                                else
                                                {
                                                    //if ("ID|Creator".Contains(property.Name)) continue;

                                                    if ("FileID" == (property.Name))   //파일ID Setting
                                                    {
                                                        property.SetValue(instance, fileInfo.ID);
                                                    }
                                                    /*
                                                    else {
                                                        //Ext add
                                                        try
                                                        {
                                                            //if (property.Name.ToUpper() == "WAREHOUSECOST")
                                                            //{
                                                            //    break;
                                                            //}
                                                            var val = row[property.Name];
                                                            if (val == DBNull.Value)
                                                            {
                                                                val = null;
                                                            }
                                                            //if (property.Name == "Ext" + extIndex++)
                                                            //{
                                                            //    property.SetValue(instance, val != null ? val.ToString().Trim() : "");
                                                            //}
                                                            property.SetValue(instance, val != null ? val.ToString().Trim() : "");

                                                        }
                                                        catch (Exception ex)
                                                        {
                                                            throw ex;
                                                        }                                                        
                                                    }
                                                    */
                                                }

                                            } // property loop


                                            //if (inst.Name != null)
                                            sqlTable.InsertOnSubmit(instance);
                                            idx++;  //processed row Index

                                        } // DataRow loop

                                        // Submit changes.
                                        fileInfo.Result = "S";
                                        fileInfo.Reason = (idx).ToString();
                                        dbContext.SubmitChanges();
                                        successFileCount++;

                                    } // using DataTable

                                    success[0] = successFileCount;
                                    success[1] = idx;

                                    return myDataSet;
                                } // using DataSet

                            } // Attributes exist

                        } // No records were preexisting in the database table.
                    }
                    catch (Exception ex)
                    {
                        string msg = ex.Message;

                        fileInfo.Result = "E";
                        fileInfo.Reason = msg + "\r\nwork row index is (" + idx + ")";
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