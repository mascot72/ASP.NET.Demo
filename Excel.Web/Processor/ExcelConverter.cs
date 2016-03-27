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
using System.Collections;
using Excel.Domain.Entites;
using Excel.Domain.Concrete;
using System.Transactions;

namespace Excel.Web.Processor
{
    public class ExcelConverter
    {
        //Members declare
        int extCount = 100;  //추가컬럼갯수
        string classNamespace = "Excel.Domain.Entites.Valuation";
        //min DateTime
        DateTime nullableDate = new DateTime(1900, 1, 1);

        //개행문자제거
        private string Cleaning(string source, string deleteText = @"\W")   //"(?<!\r)\n"
        {
            try
            {
                return Regex.Replace(source, deleteText, "");
            }
            catch (Exception)
            {
                throw;
            }
        }

        //Excel to DataSet변환(Office.Interop방식)
        public DataSet ExcelToDataSet(string fileName)
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
                var classList = Assembly.GetExecutingAssembly().GetTypes().Where(t => String.Equals(t.FullName, classNamespace, StringComparison.Ordinal)).ToList();
                classList = new List<Type>() { typeof(Excel.Domain.Entites.Valuation) };

                var classQuery = (from tmpClass in classList
                                  where !tmpClass.Name.StartsWith("I")
                                  select tmpClass);
                foreach (var tmpClass in classQuery)
                {
                    var attributes = tmpClass.GetCustomAttributes(typeof(System.Data.Linq.Mapping.TableAttribute), true);    //Class의 속성을 가져온다
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
                                        if ((property.PropertyType == typeof(DateTime)) ||
                                            (property.PropertyType == typeof(DateTime?)))   //DateTime이면
                                        {
                                            dt.Columns.Add(new DataColumn(column, typeof(DateTime)));
                                        }
                                        else
                                        {
                                            dt.Columns.Add(new DataColumn(column, property.PropertyType));
                                        }
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
                IEnumerable<Excel.Domain.Entites.FileImport> fileTable = null;
                using (var context = new EFDbContext())
                {
                    fileTable = context.FileImports.ToList();

                }
                return fileTable;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        /*
        //확장정보 목록조회
        public IEnumerable<ExtendDefine> GetModelExtendList()
        {
            try
            {
                IEnumerable<ExtendDefine> fileTable = null;
                using (var dbContext = new ExtendDefineRepository())
                {
                    fileTable = dbContext.ExtendDefines;
                }
                return fileTable;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

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
        public bool AddModelExtend(ExtendDefine column)
        {
            try
            {
                bool result = default(bool);
                using (var dbContext = new ExtendDefineRepository())
                {
                    result = dbContext.AddModel(column);
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
                IEnumerable<Excel.Domain.Entites.ModelExtendColumn> fileTable = null;
                using (var dbContext = new DataContext(sqlDatabase))
                {
                    if (name != null)
                        fileTable = dbContext.GetTable<Excel.Domain.Entites.ModelExtendColumn>().Where(e => e.Name == name).ToList();
                    else
                        fileTable = dbContext.GetTable<Excel.Domain.Entites.ModelExtendColumn>().ToList();
                }
                return fileTable;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        */

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
        public DataSet ExcelToDB(string fileName, int[] success, string processState = "")
        {
            //var extList = GetModelExtendList();

            if (success == null)
            {
                int[] aa = { 0, 0, 0 };
                success = aa;
            }

            //Namespace에 속한 Model조회
            var classList = Assembly.GetExecutingAssembly().GetTypes().Where(t => String.Equals(t.Namespace, classNamespace, StringComparison.Ordinal)).ToList();
            classList = new List<Type>() { typeof(Excel.Domain.Entites.Valuation) };

            var filesToProcess = fileName;
            int idx = 0;
            int successFileCount = 0;   //완료된 파일갯수 반환용
            int targetRowCount = 0; //대상 Excel행수 반환용
            FileImport fileInfo = null;

            //0. Validation
            FileInfo files = new System.IO.FileInfo(fileName);  //파일 존재시에만 처리한다.
            if (!files.Exists)
            {
                return null;
            }
            using (EFDbContext context = new EFDbContext())
            {
                //1. 파일정보등록
                fileInfo = new FileImport()
                {
                    Name = files.Name,
                    Creator = "Isaac",
                    CreateDate = DateTime.Now,
                    Extend = "",
                    Path = files.DirectoryName,
                    Reason = "",
                    Result = "P",
                    Remark = "file Last write time at(" + files.LastWriteTime.ToString() + ")",
                    ExtName = Cleaning(files.Extension),
                    Size = files.Length
                };
                context.FileImports.Add(fileInfo);
                context.SaveChanges();
            }

            try
            {
                //2. 데이타등록
                using (TransactionScope scope = new TransactionScope(TransactionScopeOption.Required, new TransactionOptions() { IsolationLevel = System.Transactions.IsolationLevel.ReadUncommitted }))
                {
                    using (EFDbContext context = new EFDbContext())
                    {
                        var classQuery = (from tmpClass in classList
                                          where !tmpClass.Name.StartsWith("I")
                                          select tmpClass);

                        foreach (var tmpClass in classQuery)
                        {   
                            try
                            {
                                using (var myDataSet = ExcelToDataSet(fileName))
                                {
                                    // The data table will have the same name
                                    using (var dataTable = myDataSet.Tables[1])
                                    {
                                        //Copy Datatable
                                        var currDataTable = dataTable.Copy();
                                        if (!currDataTable.Columns.Contains("Reason"))
                                        {
                                            currDataTable.Columns.Add("Reason");
                                        }

                                        using (ValuationRepository mstContext = new ValuationRepository())
                                        using (ExtendDefineRepository extDefContext = new ExtendDefineRepository())    //using ExtDefine context
                                        using (ExtendContentRepository extContContext = new ExtendContentRepository())
                                        {
                                            //2-1. 확장컬럼 나오면 등록(Propery에 없다면 확장컬럼)                                            
                                            foreach (DataColumn col in currDataTable.Columns)
                                            {
                                                if ("Profit%".ToUpper().Contains(col.ColumnName.ToUpper())) continue;   //예외

                                                //엑셀컬럼중 Model Property에 존재하지 않으면
                                                if (tmpClass.GetProperties().Count(e => e.Name.ToUpper() == col.ColumnName.ToUpper()) < 1)
                                                {
                                                    var newExt = new ExtendDefine()
                                                    {
                                                        Name = col.ColumnName.Trim()
                                                    };
                                                    bool exists = extDefContext.AddModel(newExt);   //확장정의Table에 등록
                                                    
                                                    if (exists) //새로 추가되면 Remark에 기록
                                                    {
                                                        fileInfo.Remark += (fileInfo.Remark != "" ? "|" : ", add extend column:") + col.ColumnName + "(" + newExt.ID + ")";
                                                        fileInfo.Extend += (fileInfo.Extend != "" ? "|" : "") + col.ColumnName + "(" + newExt.ID + ")"; //확장정보에 기록
                                                    }
                                                    else
                                                    {
                                                        var curExt = extDefContext.FindName(col.ColumnName.Trim());
                                                        if (curExt != null)
                                                            fileInfo.Extend += (fileInfo.Extend != "" ? "|" : "") + col.ColumnName + "(" + curExt.ID + ")"; //확장정보에 기록
                                                    }
                                                }

                                            }
                                            //2-2. 저장(행만큼
                                            foreach (DataRow row in currDataTable.Rows)
                                            {
                                                targetRowCount++;   //대상 컬럼 증가
                                                //each row reset Ext? Columns
                                                if (row["Name"] == null || row["Name"].ToString().Trim() == "")
                                                {
                                                    //throw new ApplicationException("Name컬럼이 존재하지 않습니다");
                                                    continue;   //병합된 컬럼이 존재시 이Row만 제외하고 통과해야 한다
                                                }

                                                var instance = Activator.CreateInstance(tmpClass);
                                                var properties = tmpClass.GetProperties();
                                                var propertyQuery = (from property in properties
                                                                     where property.CanWrite
                                                                     select property);

                                                var valuationRow = instance as Valuation;

                                                //1 컴럼 매핑 및 확장속성 값 적용
                                                #region Table mapping(Extended)
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
                                                            //확장컬럼위치에 넣기
                                                            var curExt = extDefContext.FindName(col.ColumnName.Trim());   //확장정의명과 컬럼명이 일치한것 가져오기
                                                            if (curExt != null && row[col.ColumnName] != null && row[col.ColumnName].ToString().Trim() != "")  //내용이 있을때만
                                                            {
                                                                //import ModelExtendContent
                                                                var extData = new ExtendContent()
                                                                {
                                                                    //ImportID = valuationRow.ID,
                                                                    EID = curExt.ID,
                                                                    Name = curExt.Name,
                                                                    Content = row[col.ColumnName].ToString().Trim(),
                                                                    Ref1 = idx
                                                                };
                                                                valuationRow.ExtendContent.Add(extData);
                                                                //extContContext.AddModel(extData);

                                                                row["Reason"] += (row["Reason"] != null && row["Reason"].ToString() != "" ? "|" : "") + col.ColumnName + "(" + curExt.ID + ")";
                                                            }

                                                        }
                                                    }
                                                }
                                                #endregion

                                                //2 기본컬럼정보 값 적용
                                                #region SetValue(column loop)
                                                foreach (PropertyInfo property in propertyQuery)
                                                {
                                                    if (property.Name.ToLower() == "buyer")
                                                    {
                                                        var val1 = row[property.Name];
                                                    }

                                                    int idxColumn = currDataTable.Columns.IndexOf(property.Name);
                                                    if (idxColumn > -1 && currDataTable.Columns[idxColumn].ColumnName.ToUpper() == property.Name.ToUpper())  //컬럼이 존재시
                                                    {
                                                        var val = row[property.Name];
                                                        if (val == DBNull.Value)
                                                        {
                                                            val = null;
                                                        }

                                                        if (val == null)    //값이 Null이면
                                                        {
                                                            if ((property.PropertyType == typeof(DateTime)) ||
                                                                (property.PropertyType == typeof(DateTime?)))   //DateTime이면
                                                            {
                                                                property.SetValue(instance, nullableDate);
                                                            }
                                                        }
                                                        else if (val.GetType() != property.PropertyType)    //실제 값과 속성의 Type이 다르면
                                                        {
                                                            if ((property.PropertyType == typeof(DateTime)) ||
                                                                     (property.PropertyType == typeof(DateTime?)))
                                                            {
                                                                var newVal = val.ToString();
                                                                var nullableDate = (string.IsNullOrWhiteSpace
                                                                   (newVal) ? (DateTime?)null : DateTime.Parse(newVal));
                                                                property.SetValue(instance, nullableDate);
                                                            }
                                                            else if ((val.GetType() == typeof(DateTime)) ||
                                                                (val.GetType() == typeof(DateTime?)))
                                                            {
                                                                property.SetValue(instance, val.ToString());
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
                                                            //값넣기
                                                            property.SetValue(instance, val);
                                                        }

                                                    } // dbColumn exists
                                                    else
                                                    {
                                                        //속성에 없는 Row(column)이면(Ext1~Ext50, ID, CreateDate, Creator, Ref2)

                                                        if ("FileID" == (property.Name))   //파일ID Setting
                                                        {
                                                            property.SetValue(instance, fileInfo.ID);   //FileID저장
                                                        }
                                                        else if ("Ref1" == property.Name)    //확장컬럼 Update용
                                                        {
                                                            property.SetValue(instance, idx);   //순서 index저장
                                                        }
                                                    }

                                                } // property loop
                                                #endregion

                                                mstContext.AddModel(valuationRow);                                                
                                                idx++;  //processed row Index

                                            } // DataRow loop
                                        }// using ExtDefine context

                                        // Submit changes.
                                        using (FileImportRepository fileContext = new FileImportRepository())
                                        {
                                            var fileEntity = fileContext.GetModel(fileInfo.ID);
                                            fileEntity.Result = "S";
                                            fileEntity.Reason = (idx).ToString();
                                            fileEntity.Remark = fileInfo.Remark;
                                            fileEntity.Extend = fileInfo.Extend;
                                            fileEntity.UpdateDate = DateTime.Now;
                                            fileEntity.Updater = "Isaac";

                                            fileContext.ModifyModel(fileEntity);
                                        }
                                        context.SaveChanges();
                                        successFileCount++; //proceded file count

                                        //extContentTable

                                    } // using DataTable

                                    success[0] = successFileCount;
                                    success[1] = idx;
                                    success[2] = targetRowCount;

                                    scope.Complete();   //Transaction Complete!!!

                                    return myDataSet;
                                } // using DataSet


                            }
                            catch (Exception ex)
                            {
                                throw ex;
                            }

                        } // class loop
                    }

                } // using DataContext
            }
            catch (Exception ex)
            {
                string msg = GetException(ex).Message;
                //파일내용 갱신
                using (FileImportRepository fileContext = new FileImportRepository())
                {
                    var fileEntity = fileContext.GetModel(fileInfo.ID);
                    fileEntity.Result = "E";
                    fileEntity.Reason = msg + "\r\nwork row index is (" + idx + ")";
                    fileEntity.UpdateDate = DateTime.Now;
                    fileEntity.Updater = "Isaac";
                    fileContext.ModifyModel(fileEntity);                    
                }
            }
            return null;
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

        //내부예외가져오기
        public Exception GetException(Exception exception)
        {
            if (exception.InnerException != null)
                return GetException(exception.InnerException);
            else
                return exception;
        }

    }
}