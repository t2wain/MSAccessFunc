using Microsoft.Office.Interop.Access.Dao;
using System.Collections;
using System.Reflection;
using DAO = Microsoft.Office.Interop.Access.Dao;

namespace MSAccessLib
{
    public class DB : IDisposable
    {
        #region Other

        DBEngine _eng = null!;
        Workspace _wk = null!;

        const int TBL_ATTR_LOCAL = 0;
        const int TBL_ATTR_LINK = 1073741824;

        public DB()
        {
            _eng = new DBEngineClass();
            _wk = _eng.Workspaces[0];
        }
        public void Dispose()
        {
            _wk?.Close();
            _wk = null;
            _eng = null;
        }

        #endregion

        #region Data Provider

        public Database? OpenExcel(string fileName) => 
            OpenDB(fileName, Connect.Excel(fileName));

        public Database? OpenAccessDB(string fileName, string pwd = "") => 
            OpenDB(fileName, Connect.MSAccess(fileName, pwd));

        public Database? OpenODBC(string connString, bool openreadonly = false) =>
            OpenDB("", connString, openreadonly);

        public Database? OpenDB(string name, string connectString, bool openreadonly = false) => 
            _wk.OpenDatabase(name, false, openreadonly, connectString);

        public Database? CreateMSAccess(string fileName) =>
            File.Exists(fileName) ? null :
                _wk.CreateDatabase(
                    fileName,
                    string.Format("{0}", LanguageConstants.dbLangGeneral),
                    DatabaseTypeEnum.dbVersion150
                );

        #endregion

        #region Table properties

        public void PrintDatabaseTables(Database? db, 
            bool hideEmptyProperty = false, Predicate<TableDef>? tblFilter = null)
        {
            if (db == null || db.TableDefs.Count == 0) return;

            Console.WriteLine(string.Format("Database Name: {0}\n", db.Name));

            var f = tblFilter ?? 
                (t => t.Attributes == TBL_ATTR_LOCAL || t.Attributes == TBL_ATTR_LINK);
            foreach(TableDef t in db.TableDefs)
            {
                if (!f(t)) continue;
                Console.WriteLine(String.Format("{0}", t.Name));
                PrintProperties(t.Properties);
                PrintTableFields(t.Fields, "\t", hideEmptyProperty);
                PrintIndexes(t.Indexes, "\t", hideEmptyProperty);
            }
        }

        protected void PrintTableFields(Fields fields, string tabs = "\t", bool hideEmptyProperty = false)
        {
            if (fields.Count == 0) return;
            foreach(Field f in fields)
            {
                Console.WriteLine(string.Format(tabs + "{0} - F", f.Name));
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

                PrintProperties(f.Properties, tb, hideEmptyProperty);
            }
            Console.WriteLine();
        }

        protected void PrintIndexes(Indexes indexes, string tabs = "\t", bool hideEmptyProperty = false)
        {
            string tb = tabs + "\t";
            foreach (DAO.Index i in indexes)
            {
                Console.WriteLine(string.Format(tabs + "{0} - I", i.Name));
                PrintProperties(i.Properties, tb, hideEmptyProperty);

                // Show fields in index
                TryGetValue(i.Fields, "GetEnumerator", out var c);
                var fir = c as System.Collections.IEnumerator;
                if (fir == null) continue;
                while (fir.MoveNext())
                {
                    Field f = (fir.Current as Field)!;
                    Console.WriteLine(string.Format(tb + "{0} - IF", f.Name));
                }
                Console.WriteLine();
            }
        }


        #endregion

        #region All database properties

        public void PrintDBEngine()
        {
            Console.WriteLine("Engine:");
            TryGetValue(_eng, "DefaultType", out var dt); // WorkspaceTypeEnum
            TryGetValue(_eng, "Errors", out var errs);
            TryGetValue(_eng, "IniPath", out var pt);
            TryGetValue(_eng, "LoginTimeout", out var to);
            TryGetValue(_eng, "Properties", out var props);
            TryGetValue(_eng, "SystemDB", out var s);
            TryGetValue(_eng, "Version", out var v);
            TryGetValue(_eng, "Workspaces", out var ws);
            Console.WriteLine();

            Console.WriteLine("Engine properties:");
            if (props != null)
                PrintProperties(props as Properties);
        }

        /// <summary>
        /// Show information about the database including tables and queries. 
        /// </summary>
        public void PrintDatabase(Database? db, bool hideEmptyProperty = false)
        {
            if (db != null)
            {
                Console.WriteLine(string.Format("Database Name: {0}", db.Name));

                // Show database properties
                PrintProperties(db.Properties, hideEmptyProperty: hideEmptyProperty);

                // Show database containers
                PrintContainers(db.Containers, hideEmptyProperty);

                // Show tables
                PrintTableDefs(db.TableDefs, hideEmptyProperty);

                // Show queries
                PrintQueryDefs(db.QueryDefs, hideEmptyProperty);
            }
        }


        /// <summary>
        /// Show information about the database Containers
        /// </summary>
        protected void PrintContainers(Containers containers, bool hideEmptyProperty = false)
        {
            if (containers.Count < 1) return;
            Console.WriteLine("Containers:");
            foreach (Container? c in containers)
            {
                if (c == null) break;
                Console.WriteLine(string.Format("\t{0}  - C", c.Name));

                // Show properties of container
                PrintProperties(c.Properties, "\t\t", hideEmptyProperty);

                // Show properties of documents
                PrintDocuments(c.Documents, "\t\t", hideEmptyProperty);
            }
            Console.WriteLine();
        }

        /// <summary>
        /// Show information about the Documents in a Container
        /// </summary>
        protected void PrintDocuments(Documents documents, string tabs = "\t\t", bool hideEmptyProperty = false)
        {
            if (documents.Count < 1) return;
            Console.WriteLine(tabs + "Documents:");
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
                Console.WriteLine(string.Format(tabs + "\t{0} - D", d.Name));
                PrintProperties(d.Properties, tabs + "\t\t", hideEmptyProperty);
            }
        }

        protected void PrintQueryDefs(QueryDefs qdefs, bool hideEmptyProperty = false)
        {
            if (qdefs.Count < 1) return;
            Console.WriteLine("Database Queries:");
            foreach (QueryDef q in qdefs)
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
                Console.WriteLine(string.Format("\t{0} - Q", q.Name));
                PrintProperties(q.Properties, "\t\t", hideEmptyProperty);
            };
            Console.WriteLine();
        }

        protected void PrintTableDefs(TableDefs tdefs, bool hideEmptyProperty = false)
        {
            if (tdefs.Count < 1) return;
            Console.WriteLine("Database Tables:");
            foreach (TableDef t in tdefs)
            {
                switch(t.Attributes)
                {
                    case 0: // base table
                    case 1073741824: // link table
                        Console.WriteLine(string.Format("\t{0} - T", t.Name));
                        var tb = "\t\t";
                        //PrintProperty(t, "Attributes", "", tb, hideEmptyProperty);
                        //PrintProperty(t, "ConflictTable", "", tb, hideEmptyProperty);
                        //PrintProperty(t, "Connect", "", tb, hideEmptyProperty);
                        //PrintProperty(t, "DateCreated", "", tb, hideEmptyProperty);
                        //PrintProperty(t, "Fields", "", tb, hideEmptyProperty);
                        //PrintProperty(t, "Indexes", "", tb, hideEmptyProperty);
                        //PrintProperty(t, "LastUpdated", "", tb, hideEmptyProperty);
                        //PrintProperty(t, "Name", "", tb, hideEmptyProperty);
                        //PrintProperty(t, "Properties", "", tb, hideEmptyProperty);
                        //PrintProperty(t, "RecordCount", "", tb, hideEmptyProperty);
                        //PrintProperty(t, "ReplicaFilter", "", tb, hideEmptyProperty);
                        //PrintProperty(t, "SourceTableName", "", tb, hideEmptyProperty);
                        //PrintProperty(t, "Updatable", "", tb, hideEmptyProperty);
                        //PrintProperty(t, "ValidationRule", "", tb, hideEmptyProperty);
                        //PrintProperty(t, "ValidationText", "", tb, hideEmptyProperty);
                        //Console.WriteLine();
                        PrintProperties(t.Properties, tb, hideEmptyProperty);
                        break;
                }
            };
            Console.WriteLine();
        }

        protected void PrintProperties(Properties props, string tabs = "\t", bool hideEmptyProperty = false)
        {
            if (props?.Count < 1) return;
            foreach (Property prop in props)
            {
                PrintProperty(prop, prop.Name, "Value", tabs, hideEmptyProperty);
            };
            Console.WriteLine();
        }

        #endregion

        #region Utility

        protected void PrintProperty(object obj, string propName, string propValue = "", 
            string tabs = "\t", bool hideEmptyProperty = false)
        {
            var pv = string.IsNullOrWhiteSpace(propValue) ? propName : propValue;
            bool s = TryGetValue(obj, pv, out var v);
            if (!hideEmptyProperty || (s && v != null && !string.IsNullOrWhiteSpace(v.ToString())))
                Console.WriteLine(string.Format(tabs + "{0} : {1}", propName, v));
        }

        protected bool TryGetValue(object obj, string propName, out dynamic? value, object[]? paramValues = null)
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

        #region Link tables

        public void LinkTables(Database destDb, Database srcDb, Predicate<TableDef> filter)
        {
            var cnnstr = GetConnectString(srcDb);
            foreach(TableDef st in srcDb.TableDefs)
            {
                if (!filter(st)) continue;
                var dt = destDb.CreateTableDef();
                dt.Name = st.Name;
                dt.SourceTableName = 
                    string.IsNullOrEmpty(st.SourceTableName) ? st.Name : st.SourceTableName;
                if (st.Attributes == TBL_ATTR_LOCAL)
                    dt.Connect = cnnstr;
                else if (st.Attributes == TBL_ATTR_LINK)
                    dt.Connect = st.Connect;
                destDb.TableDefs.Append(dt);
                dt.RefreshLink();
            }
        }

        protected string GetConnectString(Database db)
        {
            var cnnstr = "";
            if (db.Name.EndsWith(".accdb") || db.Name.EndsWith(".mdb"))
            {
                cnnstr = db.Connect;
                cnnstr = string.IsNullOrWhiteSpace(cnnstr) ? Connect.MSAccess(db.Name, "") : cnnstr;
                cnnstr = cnnstr.StartsWith("MS Access") ? cnnstr : "MS Access" + cnnstr;
            }
            else cnnstr = db.Connect;

            return cnnstr;
        }

        #endregion

        #region Import tables

        public void ImportTables(Database destDb, Database srcDb, Predicate<TableDef> filter)
        {
            var srcCnnString = GetConnectString(srcDb);
            foreach (TableDef st in srcDb.TableDefs)
            {
                if (!filter(st)) continue;

                // create table and duplicate data
                var sql = string.Format("select * into [{0}] from [{1}].[{0}]", st.Name, srcCnnString);
                destDb.Execute(sql, RecordsetOptionEnum.dbFailOnError);

                // duplicate indexes
                destDb.TableDefs.Refresh();
                CopyIndexes(destDb.TableDefs[st.Name], st);
            }
        }
        protected void CopyIndexes(TableDef dt, TableDef st)
        {
            foreach(DAO.Index si in st.Indexes)
            {
                TryGetValue(si.Fields, "GetEnumerator", out var c);
                var fir = c as IEnumerator;
                if (fir == null) continue;

                var di = dt.CreateIndex(si.Name);
                dynamic fields = di.Fields;
                while (fir.MoveNext())
                {
                    var sf = (fir.Current as Field)!;
                    var df = di.CreateField(sf.Name);
                    fields.Append(df);
                }

                di.Primary = si.Primary;
                di.Unique = si.Unique;
                di.IgnoreNulls = si.IgnoreNulls;
                di.Required = si.Required;
                di.Clustered = si.Clustered;
                dt.Indexes.Append(di);
            }
        }

        protected void CopyIndexes2(Database destDb, TableDef dt, TableDef st)
        {
            //var dt = destDb.TableDefs[st.Name];
            foreach (DAO.Index si in st.Indexes)
            {
                string sqlIdx = "";
                string sqlWith = "";
                string colNames = string.Join(",", GetIndexColumns(si).ToArray());
                if (si.Primary)
                    sqlWith = " WITH PRIMARY";
                else if (si.Required)
                    sqlWith = " WITH DISALLOW NULL";
                else if (si.IgnoreNulls)
                    sqlWith = " WITH IGNORE NULL";

                sqlIdx = string.Format("CREATE{0} INDEX {1} on {2} ({3}){4};",
                si.Unique ? " UNIQUE" : "",
                si.Name,
                st.Name,
                colNames,
                sqlWith);

                destDb.Execute(sqlIdx, RecordsetOptionEnum.dbFailOnError);
            }
        }

        protected List<string> GetIndexColumns(DAO.Index si)
        {
            var lst = new List<string>();

            //// Indexes iterator
            TryGetValue(si.Fields, "GetEnumerator", out var ir);
            var sfir = ir as System.Collections.IEnumerator;
            if (sfir == null) return lst;

            while (sfir.MoveNext())
            {
                var sf = (sfir.Current as Field)!;
                lst.Add(sf.Name);
            }
            return lst;
        }

        protected void CopyRelations(TableDef dt, TableDef st)
        {

        }

        #endregion

    }
}