using NPOI.HSSF.UserModel;
using NPOI.HSSF.Util;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace ConvertData
{
    public class CData
    {
        /// <summary>
        /// 将DataTable 转换为 List,不转换表头
        /// </summary>
        /// <param name="dt"></param>
        /// <param name="start_index"></param>
        /// <returns></returns>
        public List<string[]> ConvertDataTableToList_Nohead(DataTable dt, int start_index)
        {
            List<string[]> ll = new List<string[]>();
            for (int a = start_index; a < dt.Rows.Count; a++)
            {
                DataRow row = dt.Rows[a];
                string rowStr = string.Empty;
                for (int b = 0; b < row.ItemArray.Count(); b++)
                {
                    if (b == row.ItemArray.Count() - 1)
                    {
                        rowStr += row.ItemArray[b];
                    }
                    else
                    {
                        rowStr += row.ItemArray[b] + "`";
                    }
                }
                ll.Add(rowStr.Split('`'));
            }
            return ll;
        }

        /// <summary>
        /// 将DataTable 转换为 List,包括转换表头
        /// </summary>
        /// <param name="dt"></param>
        /// <param name="start_index"></param>
        /// <returns></returns>
        public List<string[]> ConvertDataTableToList(DataTable dt, int start_index)
        {
            List<string[]> ll = new List<string[]>();
            string cols = string.Empty;

            for (int i = 0; i < dt.Columns.Count; i++)
            {
                cols += (dt.Columns[i].ColumnName.ToString() + ",");
            }
            cols = cols.TrimEnd(',');
            string[] cols_temp = cols.Split(',');
            ll.Add(cols_temp);
            for (int a = start_index; a < dt.Rows.Count; a++)
            {
                DataRow row = dt.Rows[a];
                string rowStr = string.Empty;

                for (int b = 0; b < row.ItemArray.Count(); b++)
                {
                    if (b == row.ItemArray.Count() - 1)
                    {
                        rowStr += row.ItemArray[b].ToString();
                    }
                    else
                    {
                        rowStr += row.ItemArray[b].ToString() + "`";
                    }
                }
                ll.Add(rowStr.Split('`'));
            }
            return ll;
        }

        /// <summary>
        /// 将Array转换成DataTable
        /// </summary>
        /// <param name="arr"></param>
        /// <returns></returns>
        public static DataTable ConvertArrayToDataTable(string[,] arr)
        {
            DataTable dataSouce = new DataTable();
            for (int i = 0; i < arr.GetLength(1); i++)
            {
                DataColumn newColumn = new DataColumn(i.ToString(), arr[0, 0].GetType());
                dataSouce.Columns.Add(newColumn);
            }
            for (int i = 0; i < arr.GetLength(0); i++)
            {
                DataRow newRow = dataSouce.NewRow();
                for (int j = 0; j < arr.GetLength(1); j++)
                {
                    newRow[j.ToString()] = arr[i, j];
                }
                dataSouce.Rows.Add(newRow);
            }
            return dataSouce;
        }

        /// <summary>
        /// 将List转换为DataTable
        /// </summary>
        /// <param name="list"></param>
        /// <returns></returns>
        public static DataTable ListToDataTable(IList list)
        {
            DataTable result = new DataTable();
            if (list.Count > 0)
            {
                PropertyInfo[] propertys = list[0].GetType().GetProperties();
                foreach (PropertyInfo pi in propertys)
                {
                    //获取类型
                    Type colType = pi.PropertyType;
                    //当类型为Nullable<>时
                    if ((colType.IsGenericType) && (colType.GetGenericTypeDefinition() == typeof(Nullable<>)))
                    {
                        colType = colType.GetGenericArguments()[0];
                    }
                    result.Columns.Add(pi.Name, colType);
                }
                for (int i = 0; i < list.Count; i++)
                {
                    ArrayList tempList = new ArrayList();
                    foreach (PropertyInfo pi in propertys)
                    {
                        object obj = pi.GetValue(list[i], null);
                        tempList.Add(obj);
                    }
                    object[] array = tempList.ToArray();
                    result.LoadDataRow(array, true);
                }
            }
            return result;
        }

        /// <summary>
        /// 将List 导出成Excel
        /// </summary>
        /// <param name="strFileName"></param>
        /// <param name="ls_values"></param>
        public static void Export(string strFileName, List<string[]> ls_values)
        {
            string ext_ = strFileName.Substring(strFileName.LastIndexOf("."));

            bool xls = true;

            if (ext_.ToLower() != ".xlsx")
            {
                xls = false;
            }
            IWorkbook wb = xls ? new XSSFWorkbook() as IWorkbook : new HSSFWorkbook() as IWorkbook;

            ISheet sheet1 = wb.CreateSheet("Sheet1");

            if (!xls)
            {
                sheet1.TabColorIndex = HSSFColor.Red.Index;
            }

            for (int i = 0; i < ls_values.Count; i++)
            {
                IRow irow = sheet1.CreateRow(i);

                for (int j = 0; j < ls_values[i].Length; j++)
                {
                    ls_values[i][j] = ls_values[i][j].Replace("ListViewSubItem: {", "");
                    ls_values[i][j] = ls_values[i][j].Replace("}", "");
                    irow.CreateCell(j).SetCellValue(ls_values[i][j]);
                }
            }
            FileStream file = new FileStream(strFileName, FileMode.Create);
            wb.Write(file);
            file.Close();
        }
    }
}
