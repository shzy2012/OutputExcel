using System;
using System.Collections.Generic;
using System.Data;
using System.Reflection;

/// <summary>
/// Summary description for OutputExcel
/// </summary>
public class OutputExcel
{
    public static void ResponseExcel<T>(System.Web.HttpResponse response, List<T> items)
    {
        try
        {
            var dt = ToDataTable(items);
            string attachment = "attachment; filename=vauExcel.xls";
            response.ClearContent();
            response.AddHeader("content-disposition", attachment);
            response.ContentType = "application/vnd.ms-excel";
            string tab = "";
            foreach (DataColumn dc in dt.Columns)
            {
                response.Write(tab + dc.ColumnName);
                tab = "\t";
            }
            response.Write("\n");
            int i;
            foreach (DataRow dr in dt.Rows)
            {
                tab = "";
                for (i = 0; i < dt.Columns.Count; i++)
                {
                    response.Write(tab + dr[i].ToString());
                    tab = "\t";
                }
                response.Write("\n");
            }
            response.Flush();
            response.Close();
        }
        catch (Exception ex)
        {
            throw ex;
        }
    }

    public static DataTable ToDataTable<T>(List<T> items)
    {
        DataTable dataTable = new DataTable(typeof(T).Name);

        //Get all the properties
        PropertyInfo[] Props = typeof(T).GetProperties(BindingFlags.Public | BindingFlags.Instance);
        foreach (PropertyInfo prop in Props)
        {
            //Setting column names as Property names
            dataTable.Columns.Add(prop.Name);
        }
        foreach (T item in items)
        {
            var values = new object[Props.Length];
            for (int i = 0; i < Props.Length; i++)
            {
                //inserting property values to datatable rows
                values[i] = Props[i].GetValue(item, null);
            }
            dataTable.Rows.Add(values);
        }
        //put a breakpoint here and check datatable
        return dataTable;
    }
    
        public static void NewResponseExcel<T>(System.Web.HttpResponse response, List<T> items)
    {
        try
        {
            string attachment = "attachment; filename=vauExcel.xls";
            response.ClearContent();
            response.AddHeader("content-disposition", attachment);
            response.ContentType = "application/vnd.ms-excel";
            string tab = "";

            //Get all the properties
            PropertyInfo[] Props = typeof(T).GetProperties(BindingFlags.Public | BindingFlags.Instance);
            foreach (PropertyInfo prop in Props)
            {
                response.Write(tab + prop.Name);
                tab = "\t";
            }
            response.Write("\n");
            foreach (T item in items)
            {
                var values = new object[Props.Length];
                for (int i = 0; i < Props.Length; i++)
                {
                    values[i] = Props[i].GetValue(item, null);
                    if (values[i] != null)
                        response.Write(values[i].ToString().Trim() + "\t");
                    else
                        response.Write("\t");

                }
                response.Write("\n");
            }
            response.Flush();
            response.Close();
        }
        catch (Exception ex)
        {
            throw ex;
        }
    }
    
}
