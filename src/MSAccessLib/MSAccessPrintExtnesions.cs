using Microsoft.Extensions.Logging;
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

        public static void Print(this DBEngine eng, Context ctx)
        {
            var wr = ctx.Writer;
            wr.WriteLine("Engine:");
            //TryGetValue(eng, "DefaultType", out var dt); // WorkspaceTypeEnum
            //TryGetValue(eng, "Errors", out var errs);
            //TryGetValue(eng, "IniPath", out var pt);
            //TryGetValue(eng, "LoginTimeout", out var to);
            TryGetValue(eng, "Properties", out var props);
            //TryGetValue(eng, "SystemDB", out var s);
            //TryGetValue(eng, "Version", out var v);
            //TryGetValue(eng, "Workspaces", out var ws);
            wr.WriteLine();

            wr.WriteLine("Engine properties:");
            if (props != null)
                (props as Properties).Print(ctx);
        }

        #region  Database

        /// <summary>
        /// Show information about the database including tables and queries. 
        /// </summary>
        public static void PrintAll(this Database db, Context ctx)
        {
            var wr = ctx.Writer;
            var logger = ctx.Logger;

            wr.WriteLine("Database:\n");

            // Show database properties
            logger.LogInformation("Print database properties");
            db.Properties.Print(ctx, "\t");

            if (ctx.IsMSAccessDB)
            {
                // Show database containers
                wr.WriteLine("Containers:\n");
                logger.LogInformation("Print database containers");
                db.Containers.Print(ctx, "");
            }

            if (db.TableDefs.Count > 0)
            {
                // Show tables
                wr.WriteLine("Tables:\n");
                logger.LogInformation("Print database tables");
                db.TableDefs.Print(ctx, "");
            }

            if (db.QueryDefs.Count > 0)
            {
                // Show queries
                wr.WriteLine("Queries:\n");
                logger.LogInformation("Print database queries");
                db.QueryDefs.Print(ctx, "");
            }
        }

        public static void Print(this Database db, Context ctx, string tabs = "")
        {
            var wr = ctx.Writer;
            var logger = ctx.Logger;

            wr.WriteLine(tabs + "Database:\n");

            // Show database properties
            logger?.LogInformation("Print database properties");
            db.Properties.Print(ctx, tabs + "\t");
        }

        /// <summary>
        /// Show information about the database Containers
        /// </summary>
        public static void Print(this Containers containers, Context ctx, string tabs = "")
        {
            var wr = ctx.Writer;
            var logger = ctx.Logger;

            int total = containers.Count;
            if (total < 1) return;
            int cnt = 0;
            var tb = tabs + "\t";
            foreach (Container? c in containers)
            {
                if (c == null) break;
                logger.LogInformation(string.Format("Print {0} of {1} container: {2}", ++cnt, total, c.Name));
                wr.WriteLine(string.Format(tabs + "{0}  - C", c.Name));

                // Show properties of container
                c.Properties.Print(ctx, tb);

                // Show properties of documents
                c.Documents.Print(ctx, tb);
            }
            wr.WriteLine();
        }

        /// <summary>
        /// Show information about the Documents in a Container
        /// </summary>
        public static void Print(this Documents documents, Context ctx, string tabs = "")
        {
            var wr = ctx.Writer;
            var logger = ctx.Logger;

            int total = documents.Count;
            if (total < 1) return;
            logger.LogInformation(string.Format("Print {0} documents", total));
            wr.WriteLine(tabs + "Documents:\n");
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
                wr.WriteLine(string.Format(tabs + "{0} - D", d.Name));
                d.Properties.Print(ctx, tb);
            }
        }

        #endregion

        #region TableDef

        public static void Print(this TableDefs tables, Context ctx, string tabs = "")
        {
            var wr = ctx.Writer;
            var logger = ctx.Logger;
            var f = ctx.TableFilter;

            int cnt = 0;
            int total = tables.Count;
            foreach (TableDef t in tables)
            {
                var att = t.Attributes;
                var n = t.Name;
                if (!f(t))
                {
                    logger.LogInformation(string.Format("Skip {0} of {1} table: {2}",
                        ++cnt, total, n));
                    continue;
                }
                logger.LogInformation(string.Format("Print {0} of {1} table: {2}", ++cnt, total, n));
                t.Print(ctx, tabs);
            }
        }

        public static void Print(this TableDef table, Context ctx, string tabs = "")
        {
            var wr = ctx.Writer;
            var logger = ctx.Logger;

            var tb = tabs + "\t";
            wr.WriteLine(string.Format(tabs + "{0} - T", table.Name));
            table.Properties.Print(ctx, tb);

            wr.WriteLine(tb + "Table Fields:\n");
            logger.LogInformation(string.Format("Print {0} table fields", table.Fields.Count));
            table.Fields.Print(ctx, " - TF", tb);

            TryGetValue(table.Indexes, "Count", out var cnt);
            if (cnt != null && cnt > 0)
            {
                wr.WriteLine(tb + "Table Indexes:\n");
                logger.LogInformation(string.Format("Print {0} table indexes", cnt as object));
                table.Indexes.Print(ctx, tb);
            }
        }

        #endregion

        #region Index property

        public static void Print(this Indexes indexes, Context ctx, string tabs = "")
        {
            foreach (DAO.Index i in indexes)
            {
                i.Print(ctx, tabs);
            }
        }

        public static void Print(this DAO.Index index, Context ctx, string tabs = "")
        {
            var wr = ctx.Writer;

            wr.WriteLine(string.Format(tabs + "{0} - I", index.Name));
            string tb = tabs + "\t";
            index.Properties.Print(ctx, tb);

            // Show fields in index
            var fir = index.IndexFields();
            if (fir == null) return;
            while (fir.MoveNext())
            {
                Field f = (fir.Current as Field)!;
                wr.WriteLine(string.Format(tb + "{0} - IF", f.Name));
            }
            wr.WriteLine();
        }

        public static IEnumerator? IndexFields(this DAO.Index index)
        {
            TryGetValue(index.Fields, "GetEnumerator", out var c);
            return c as System.Collections.IEnumerator;
        }

        #endregion

        #region Querydef

        public static void Print(this QueryDefs qdefs, Context ctx, string tabs = "")
        {
            var wr = ctx.Writer;
            var logger = ctx.Logger;
            var f = ctx.QueryFilter;

            int total = qdefs.Count;
            if (total < 1) return;
            int cnt = 0;
            foreach (QueryDef q in qdefs)
            {
                if (!f(q))
                {
                    logger.LogInformation(string.Format("Skip {0} of {1} queries: {2}",
                        ++cnt, total, q.Name));
                    continue;
                }
                q.Print(ctx, tabs);
            };
            wr.WriteLine();
        }

        public static void Print(this QueryDef qdef, Context ctx, string tabs = "")
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
            var wr = ctx.Writer;
            wr.WriteLine(string.Format(tabs + "{0} - Q", qdef.Name));
            qdef.Properties.Print(ctx, tabs + "\t");
            var tb = tabs + "\t";
            wr.WriteLine(tb + "Query Fields:\n");
            foreach (Field f in qdef.Fields)
            {
                if (string.IsNullOrEmpty(f.SourceTable))
                    wr.WriteLine(string.Format(tb + "{0}", f.Name));
                else
                    wr.WriteLine(string.Format(tb + "{0} - [{1}].[{2}]", f.Name, f.SourceTable, f.SourceField));
            }
            wr.WriteLine();
        }

        #endregion

        #region Field property

        public static void Print(this Fields fields, Context ctx, 
            string suffix = "", string tabs = "")
        {
            var wr = ctx.Writer;
            if (fields.Count == 0) return;
            foreach (Field f in fields)
            {
                f.Print(ctx, suffix, tabs);
            }
            wr.WriteLine();
        }

        public static void Print(this Field field, Context ctx, 
            string suffix = "", string tabs = "")
        {
            var wr = ctx.Writer;
            wr.WriteLine(string.Format(tabs + "{0}{1}", field.Name, suffix));
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
            //wr.WriteLine();
            if (!ctx.HideFieldProperty)
                field.Properties.Print(ctx, tb);
        }

        #endregion

        #region Property

        public static void Print(this Properties props, Context ctx, string tabs = "\t")
        {
            var wr = ctx.Writer;
            var hideEmptyProperty = ctx.HideEmptyProperty;

            if (props == null || props?.Count < 1) return;
            foreach (Property prop in props!)
            {
                PrintProperty(prop, prop.Name, wr, "Value", tabs, hideEmptyProperty);
            };
            wr.WriteLine();
        }

        /// <summary>
        /// A convenient method to print a property of a COM object
        /// </summary>
        /// <param name="obj">COM object</param>
        /// <param name="propName">Property name of the COM object</param>
        /// <param name="wr">Text writer</param>
        /// <param name="propValue">Property name of the COM object</param>
        /// <param name="tabs">Tab indentation</param>
        /// <param name="hideEmptyProperty">If true then don't print when value is empty</param>
        public static void PrintProperty(object obj, string propName, TextWriter wr, string propValue = "", 
            string tabs = "\t", bool hideEmptyProperty = false)
        {
            var pv = string.IsNullOrWhiteSpace(propValue) ? propName : propValue;
            bool s = TryGetValue(obj, pv, out var v);
            if (!hideEmptyProperty || (s && v != null && !string.IsNullOrWhiteSpace(v.ToString())))
                wr.WriteLine(string.Format(tabs + "{0} : {1}", propName, v));
        }

        /// <summary>
        /// Attempt to call a COM get property or a COM method
        /// </summary>
        /// <param name="obj">COM object</param>
        /// <param name="propName">Property or method name of the COM object</param>
        /// <param name="value">Return value</param>
        /// <param name="paramValues">Arguments for a COM method</param>
        /// <returns></returns>
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

        public static bool TrySetValue(object obj, string propName, out dynamic? value, object[]? paramValues = null)
        {
            value = null;
            try
            {
                value = obj.GetType().InvokeMember(propName,
                    BindingFlags.Instance | BindingFlags.Public | BindingFlags.SetProperty,
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
