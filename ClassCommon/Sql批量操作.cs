using System;
using System.Data;
using System.Data.SqlClient;

namespace ClassCommon
{
    public class Sql批量操作
    {
        /// <summary>
        /// SqlBulkCopy批量插入数据
        /// </summary>
        /// <param name="connectionString">连接字符串</param>
        /// <param name="dt">源数据：！注意 tableName与对应列要与数据库对应</param>
        /// <param name="sqlBulkCopyOptions">按位标识</param>
        public static void SqlBulkCopyByDatatable(string connectionString, DataTable dt, SqlBulkCopyOptions sqlBulkCopyOptions)
        {
            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                using (SqlBulkCopy sqlbulkcopy = new SqlBulkCopy(connectionString, sqlBulkCopyOptions))
                {
                    try
                    {
                        sqlbulkcopy.DestinationTableName = dt.TableName;
                        for (int i = 0; i < dt.Columns.Count; i++)
                        {
                            sqlbulkcopy.ColumnMappings.Add(dt.Columns[i].ColumnName, dt.Columns[i].ColumnName);
                        }
                        sqlbulkcopy.WriteToServer(dt);
                    }
                    catch (System.Exception ex)
                    {
                        throw ex;
                    }
                }
            }
        }
        /// <summary>
        /// SqlBulkCopy批量更新数据 ---有待优化
        /// </summary>
        /// <param name="connectionString">连接字符串</param>
        /// <param name="dt">源数据：！注意 tableName与对应列要与数据库对应</param>
        /// <param name="sqlBulkCopyOptions">按位标识</param>
        /// <param name="Column">更新字段名--如根据ID更新就传入ID名称</param>
        public static void SqlBulkCopyByDatatableUpdate(string connectionString, DataTable dt, SqlBulkCopyOptions sqlBulkCopyOptions, string Column)
        {
            if (!dt.Columns.Contains(Column))
            {
                new EntryPointNotFoundException("DataTable中不存在此列");
            }
            else
            {
                dt.Columns[Column].ColumnName = Column + "_T";
            }
            try
            {
                using (SqlConnection conn = new SqlConnection(connectionString))
                {
                    using (SqlCommand command = new SqlCommand("", conn))
                    {
                        var Name = DateTime.Now.ToFileTime();
                        conn.Open();
                        var inset = "";
                        for (int i = 0; i < dt.Columns.Count; i++)
                        {
                            if (dt.Columns[i].ColumnName.ToString() != Column + "_T")
                            {
                                if (dt.Columns[i].ColumnName == "Order")
                                    inset += "[Order],";
                                else if (dt.Columns[i].ColumnName == "DUMMY-CHECK")
                                    inset += "[DUMMY-CHECK],";
                                else
                                    inset += dt.Columns[i].ColumnName + ",";
                            }
                            else
                            {
                                var type = dt.Columns[i].DataType;
                                string d = "0";
                                if (type == typeof(string))
                                    d = "'''";
                                else if (type == typeof(int))
                                    d = "0";
                                else if (type == typeof(decimal))
                                    d = "0.00";
                                else if (type == typeof(DateTime))
                                    d = "''";
                                inset += d + " as " + dt.Columns[i].ColumnName + ",";
                            }
                        }
                        inset = inset.Substring(0, inset.Length - 1);
                        var sql1 = "SELECT " + inset + " INTO ##" + Name + " FROM " + dt.TableName + " WHERE 1<>1";
                        command.CommandText = sql1;
                        command.ExecuteNonQuery();
                        using (SqlBulkCopy sqlbulkcopy = new SqlBulkCopy(connectionString, sqlBulkCopyOptions))
                        {
                            try
                            {
                                sqlbulkcopy.DestinationTableName = "dbo.##" + Name;
                                for (int i = 0; i < dt.Columns.Count; i++)
                                {
                                    sqlbulkcopy.ColumnMappings.Add(dt.Columns[i].ColumnName, dt.Columns[i].ColumnName);
                                }
                                sqlbulkcopy.WriteToServer(dt);
                            }
                            catch (System.Exception ex)
                            {
                                throw ex;
                            }
                        }
                        var upset = "";
                        var where = " T." + Column + "=Temp." + Column + "_T";
                        foreach (var item in dt.Columns)
                        {
                            if (item.ToString() != Column + "_T")
                            {
                                if (item.ToString() == "Order")
                                    upset += " T.[Order]=Temp.[Order],";
                                else if (item.ToString() == "DUMMY-CHECK")
                                    upset += " T.[DUMMY-CHECK]=Temp.[DUMMY-CHECK],";
                                else
                                    upset += " T." + item.ToString() + "=Temp." + item.ToString() + ",";
                            }
                        }
                        upset = upset.Substring(0, upset.Length - 1);
                        var sql2 = "UPDATE T SET " + upset + " FROM " + dt.TableName + " T INNER JOIN ##" + Name + " Temp ON " + where + "; DROP TABLE ##" + Name + ";";
                        command.CommandText = sql2;
                        command.ExecuteNonQuery();
                        conn.Close();
                    }
                }
            }
            catch (Exception ex)
            {

                throw;
            }

        }
    }
}
