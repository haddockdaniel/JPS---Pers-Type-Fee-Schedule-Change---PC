using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using JDataEngine;
using JurisAuthenticator;
using JurisUtilityBase.Properties;
using System.Data;
using System.Reflection;

namespace JurisUtilityBase
{
    public class Error
    {
        public bool doesPTExist(string PT, JurisUtility _jurisUtility)
        {
            bool ret = false;
            String sql = "Select * from PersonnelType where PrsTypCode = '" + PT + "'";
            DataSet fd = _jurisUtility.RecordsetFromSQL(sql);
            if (fd != null && fd.Tables.Count > 0 && fd.Tables[0].Rows.Count > 0)
            {
                ret = true;
            }
                return ret;
        }

        public bool doesCliExist(string cli, JurisUtility _jurisUtility)
        {
            bool ret = false;
            String sql = "select * from client where dbo.jfn_FormatClientCode(clicode) = '" + cli + "'";
            DataSet fd = _jurisUtility.RecordsetFromSQL(sql);
            if (fd != null && fd.Tables.Count > 0 && fd.Tables[0].Rows.Count > 0)
            {
                ret = true;
            }
            return ret;
        }

        public bool isRateNumeric(string rate, JurisUtility _jurisUtility)
        {
            bool ret = false;
            try
            {
                double aa = Convert.ToDouble(rate);
                ret = true;
            }
            catch (Exception vvt)
            { }
            return ret;
        }

        public DataTable ToDataTable<T>(IEnumerable<T> collection)
        {
            DataTable dt = new DataTable("DataTable");
            Type t = typeof(T);
            PropertyInfo[] pia = t.GetProperties();

            //Inspect the properties and create the columns in the DataTable
            foreach (PropertyInfo pi in pia)
            {
                Type ColumnType = pi.PropertyType;
                if ((ColumnType.IsGenericType))
                {
                    ColumnType = ColumnType.GetGenericArguments()[0];
                }
                dt.Columns.Add(pi.Name, ColumnType);
            }

            //Populate the data table
            foreach (T item in collection)
            {
                DataRow dr = dt.NewRow();
                dr.BeginEdit();
                foreach (PropertyInfo pi in pia)
                {
                    if (pi.GetValue(item, null) != null)
                    {
                        dr[pi.Name] = pi.GetValue(item, null);
                    }
                }
                dr.EndEdit();
                dt.Rows.Add(dr);
            }
            return dt;
        }



    }
}
