using Microsoft.Office.Interop.Access.Dao;
using System.Collections;
using System.Reflection;
using DAO = Microsoft.Office.Interop.Access.Dao;

namespace MSAccessLib
{
    public static class MSAccessPrintExtnesions
    {

        const int TBL_ATTR_LOCAL = 0;
        const int TBL_ATTR_LINK = 1073741824;

        public static void Print(this DBEngine eng)
        {
            Console.WriteLine("Engine:");
            //TryGetValue(eng, "DefaultType", out var dt); // WorkspaceTypeEnum
            //TryGetValue(eng, "Errors", out var errs);
            //TryGetValue(eng, "IniPath", out var pt);
            //TryGetValue(eng, "LoginTimeout", out var to);
            TryGetValue(eng, "Properties", out var props);
            //TryGetValue(eng, "SystemDB", out var s);
            //TryGetValue(eng, "Version", out var v);
            //TryGetValue(eng, "Workspaces", out var ws);
            Console.WriteLine();

            Console.WriteLine("Engine properties:");
            if (props != null)
                (props as Properties).Print();
        }

        #region  Database

        /// <summary>
        /// Show information about the database including tables and queries. 
        /// </summary>
        public static void Print(this Database db, bool hideEmptyProperty = false, Predicate<TableDef>? tblFilter = null)
        {
            Console.WriteLine("Database:\n");

            // Show database properties
            db.Properties.Print("\t", hideEmptyProperty);

            // Show database containers
            Console.WriteLine("Containers:\n");
            db.Containers.Print("", hideEmptyProperty);

            // Show tables
            Console.WriteLine("Tables:\n");
            db.TableDefs.Print("", hideEmptyProperty);

            // Show queries
            Console.WriteLine("Queries:\n");
            db.QueryDefs.Print("", hideEmptyProperty);
        }

        /// <summary>
        /// Show information about the database Containers
        /// </summary>
        public static void Print(this Containers containers, string tabs = "", bool hideEmptyProperty = false)
        {
            if (containers.Count < 1) return;
            var tb = tabs + "\t";
            var tb2 = tb + "\t";
            foreach (Container? c in containers)
            {
                if (c == null) break;
                Console.WriteLine(string.Format(tabs + "{0}  - C", c.Name));

                // Show properties of container
                c.Properties.Print(tb, hideEmptyProperty);

                // Show properties of documents
                c.Documents.Print(tb, hideEmptyProperty);
            }
            Console.WriteLine();
        }

        /// <summary>
        /// Show information about the Documents in a Container
        /// </summary>
        public static void Print(this Documents documents, string tabs = "", bool hideEmptyProperty = false)
        {
            if (documents.Count < 1) return;
            Console.WriteLine(tabs + "Documents:\n");
            string tb = tabs + "\t";
            foreach (Document d in documents)
            {
                if (d == null) break;

                //TryGetValue(d, "AllPermissions", out var ap);
                //TryGetValue(d, "Container", out var c);
                //TryGetValue(d, "DateCreated", out var dc);
                //TryGetValue(d, "LastUpdated", out var du);
                //TryGetValue(d, "Name", out var n);
                //TryGetValue(d, "Owner", out var o);
                //TryGetValue(d, "Permissions", out var pm);
                //TryGetValue(d, "Properties", out var p);
                //TryGetValue(d, "UserName", out var u);
                Console.WriteLine(string.Format(tabs + "{0} - D", d.Name));
                d.Properties.Print(tb, hideEmptyProperty);
            }
        }

        #endregion

        #region TableDef

        public static void Print(this TableDefs tables, string tabs = "", 
            bool hideEmptyProperty = false, Predicate<TableDef>? tblFilter = null)
        {
            var f = tblFilter ??
                (t => t.Attributes == TBL_ATTR_LOCAL || t.Attributes == TBL_ATTR_LINK);
            foreach (TableDef t in tables)
            {
                if (!f(t)) continue;
                t.Print(tabs, hideEmptyProperty);
            }
        }

        public static void Print(this TableDef table, string tabs = "", bool hideEmptyProperty = false)
        {
            var tb = tabs + "\t";
            Console.WriteLine(string.Format(tabs + "{0} - T", table.Name));
            table.Properties.Print(tb, hideEmptyProperty);
            Console.WriteLine(tb + "Table Fields:\n");
            table.Fields.Print(" - TF", tb, hideEmptyProperty);
            Console.WriteLine(tb + "Table Indexes:\n");
            table.Indexes.Print(tb, hideEmptyProperty);
        }

        #endregion

        #region Index property

        public static void Print(this Indexes indexes, string tabs = "", bool hideEmptyProperty = false)
        {
            foreach (DAO.Index i in indexes)
            {
                i.Print(tabs, hideEmptyProperty);
            }
        }

        public static void Print(this DAO.Index index, string tabs = "", bool hideEmptyProperty = false)
        {
            Console.WriteLine(string.Format(tabs + "{0} - I", index.Name));
            string tb = tabs + "\t";
            index.Properties.Print(tb, hideEmptyProperty);

            // Show fields in index
            var fir = index.IndexFields();
            if (fir == null) return;
            while (fir.MoveNext())
            {
                Field f = (fir.Current as Field)!;
                Console.WriteLine(string.Format(tb + "{0} - IF", f.Name));
            }
            Console.WriteLine();
        }

        public static IEnumerator? IndexFields(this DAO.Index index)
        {
            TryGetValue(index.Fields, "GetEnumerator", out var c);
            return c as System.Collections.IEnumerator;
        }

        #endregion

        #region Querydef

        public static void Print(this QueryDefs qdefs, string tabs = "", bool hideEmptyProperty = false)
        {
            if (qdefs.Count < 1) return;
            foreach (QueryDef q in qdefs)
            {
                q.Print(tabs, hideEmptyProperty);
            };
            Console.WriteLine();
        }

        public static void Print(this QueryDef qdef, string tabs = "", bool hideEmptyProperty = false)
        {
            //TryGetValue(q, "CacheSize", out var cs);
            //TryGetValue(q, "Connect", out var cn);
            //TryGetValue(q, "DateCreated", out var dc);
            //TryGetValue(q, "Fields", out var qfs);
            //TryGetValue(q, "hStmt", out var st);
            //TryGetValue(q, "LastUpdated", out var du);
            //TryGetValue(q, "MaxRecords", out var mr);
            //TryGetValue(q, "Name", out var n);
            //TryGetValue(q, "ODBCTimeout", out var to);
            //TryGetValue(q, "Parameters", out var pms);
            //TryGetValue(q, "Prepare", out var pr);
            //TryGetValue(q, "Properties", out var props);
            //TryGetValue(q, "RecordsAffected", out var ra);
            //TryGetValue(q, "ReturnsRecords", out var rr);
            //TryGetValue(q, "SQL", out var sql);
            //TryGetValue(q, "StillExecuting", out var se);
            //TryGetValue(q, "Type", out var tp);
            //TryGetValue(q, "Updatable", out var up);
            Console.WriteLine(string.Format(tabs + "{0} - Q", qdef.Name));
            qdef.Properties.Print(tabs + "\t", hideEmptyProperty);
            var tb = tabs + "\t";
            Console.WriteLine(tb + "Query Fields:\n");
            foreach (Field f in qdef.Fields)
            {
                if (string.IsNullOrEmpty(f.SourceTable))
                    Console.WriteLine(string.Format(tb + "{0}", f.Name));
                else
                    Console.WriteLine(string.Format(tb + "{0} - [{1}].[{2}]", f.Name, f.SourceTable, f.SourceField));
            }
            Console.WriteLine();
        }

        #endregion

        #region Field property

        public static void Print(this Fields fields, string suffix = "", string tabs = "", bool hideEmptyProperty = false)
        {
            if (fields.Count == 0) return;
            foreach (Field f in fields)
            {
                f.Print(suffix, tabs, hideEmptyProperty);
            }
            Console.WriteLine();
        }

        public static void Print(this Field field, string suffix = "", string tabs = "", bool hideEmptyProperty = false)
        {
            Console.WriteLine(string.Format(tabs + "{0}{1}", field.Name, suffix));
            string tb = tabs + "\t";
            //PrintProperty(f, "AllowZeroLength", "", tb);
            //PrintProperty(f, "Attributes", "", tb);
            //PrintProperty(f, "CollatingOrder", "", tb);
            //PrintProperty(f, "CollectionIndex", "", tb);
            //PrintProperty(f, "DataUpdatable", "", tb);
            //PrintProperty(f, "DefaultValue", "", tb);
            //PrintProperty(f, "FieldSize", "", tb);
            //PrintProperty(f, "ForeignName", "", tb);
            //PrintProperty(f, "OrdinalPosition", "", tb);
            //PrintProperty(f, "OriginalValue", "", tb);
            //PrintProperty(f, "Required", "", tb);
            //PrintProperty(f, "Size", "", tb);
            //PrintProperty(f, "SourceField", "", tb);
            //PrintProperty(f, "SourceTable", "", tb);
            //PrintProperty(f, "Type", "", tb);
            //PrintProperty(f, "ValidateOnSet", "", tb);
            //PrintProperty(f, "ValidationRule", "", tb);
            //PrintProperty(f, "ValidationText", "", tb);
            //PrintProperty(f, "Value", "", tb);
            //PrintProperty(f, "VisibleValue", "", tb);
            //Console.WriteLine();
            field.Properties.Print(tb, hideEmptyProperty);
        }

        #endregion

        #region Property

        public static void Print(this Properties props, string tabs = "\t", bool hideEmptyProperty = false)
        {
            if (props == null || props?.Count < 1) return;
            foreach (Property prop in props!)
            {
                PrintProperty(prop, prop.Name, "Value", tabs, hideEmptyProperty);
            };
            Console.WriteLine();
        }

        public static void PrintProperty(object obj, string propName, string propValue = "",
            string tabs = "\t", bool hideEmptyProperty = false)
        {
            var pv = string.IsNullOrWhiteSpace(propValue) ? propName : propValue;
            bool s = TryGetValue(obj, pv, out var v);
            if (!hideEmptyProperty || (s && v != null && !string.IsNullOrWhiteSpace(v.ToString())))
                Console.WriteLine(string.Format(tabs + "{0} : {1}", propName, v));
        }

        public static bool TryGetValue(object obj, string propName, out dynamic? value, object[]? paramValues = null)
        {
            value = null;
            try
            {
                value = obj.GetType().InvokeMember(propName,
                    BindingFlags.Instance | BindingFlags.Public | BindingFlags.GetProperty | BindingFlags.InvokeMethod,
                    Type.DefaultBinder, obj, paramValues);
                return true;
            }
            catch (Exception err)
            {
                var msg = err.Message;
                return false;
            }
        }

        #endregion

    }
}
