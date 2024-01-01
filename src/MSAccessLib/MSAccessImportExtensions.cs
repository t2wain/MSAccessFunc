using Microsoft.Office.Interop.Access.Dao;
using System.Collections;
using DAO = Microsoft.Office.Interop.Access.Dao;

namespace MSAccessLib
{
    public static class MSAccessImportExtensions
    {
        const int TBL_ATTR_LOCAL = 0;
        const int TBL_ATTR_LINK = 1073741824;

        #region Link tables

        public static void LinkToTables(this Database db, Database externalDb, Predicate<TableDef> filter)
        {
            var cnnstr = externalDb.GetConnectString();
            foreach (TableDef st in externalDb.TableDefs)
            {
                if (!filter(st)) continue;
                db.LinkToTable(st, externalDb);
            }
        }

        public static void LinkToTable(this Database db, TableDef externalTable, Database externalDb)
        {
            var cnnstr = externalDb.GetConnectString();
            var dt = db.CreateTableDef();
            dt.Name = externalTable.Name;
            dt.SourceTableName =
                string.IsNullOrEmpty(externalTable.SourceTableName) ? externalTable.Name : externalTable.SourceTableName;
            if (externalTable.Attributes == TBL_ATTR_LOCAL)
                dt.Connect = cnnstr;
            else if (externalTable.Attributes == TBL_ATTR_LINK)
                dt.Connect = externalTable.Connect;
            db.TableDefs.Append(dt);
            dt.RefreshLink();
        }

        #endregion

        #region Import tables

        public static void ImportTables(this Database db, Database externalDb, Predicate<TableDef> filter)
        {
            var srcCnnString = externalDb.GetConnectString();
            foreach (TableDef st in externalDb.TableDefs)
            {
                if (!filter(st)) continue;

                // create table and duplicate data
                var sql = string.Format("select * into [{0}] from [{1}].[{0}]", st.Name, srcCnnString);
                db.Execute(sql, RecordsetOptionEnum.dbFailOnError);

                // duplicate indexes
                db.TableDefs.Refresh();
                CopyIndexes(db.TableDefs[st.Name], st);
            }
        }

        public static void ImportTables(this Database db, TableDef externalTable, Database externalDb)
        {
            var srcCnnString = externalDb.GetConnectString();
            // create table and duplicate data
            var sql = string.Format("select * into [{0}] from [{1}].[{0}]", externalTable.Name, srcCnnString);
            db.Execute(sql, RecordsetOptionEnum.dbFailOnError);

            // duplicate indexes
            db.TableDefs.Refresh();
            var dt = db.TableDefs[externalTable.Name];
            dt.CopyIndexes(externalTable);
        }

        public static void CopyIndexes(this TableDef dt, TableDef externalTable)
        {
            foreach (DAO.Index si in externalTable.Indexes)
            {
                var fir = si.IndexFields();
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

        #endregion

    }
}
