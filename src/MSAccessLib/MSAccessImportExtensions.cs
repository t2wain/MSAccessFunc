using Microsoft.Extensions.Logging;
using Microsoft.Office.Interop.Access.Dao;
using DAO = Microsoft.Office.Interop.Access.Dao;

namespace MSAccessLib
{
    public static class MSAccessImportExtensions
    {
        const int TBL_ATTR_LOCAL = 0;
        const int TBL_ATTR_LINK = 1073741824;

        #region Link tables

        public static void LinkToTables(this Database db, Database externalDb, Context ctx)
        {
            var logger = ctx.Logger;
            var filter = ctx.TableFilter;

            var srcTables = externalDb.TableDefs;
            var destTblNames = db.TableDefs.GetTableNames();
            int cnt = 0;
            int total = srcTables.Count;
            logger.LogInformation(string.Format("Linking total number of tables: {0}", total));
            var cnnstr = externalDb.GetConnectString();
            foreach (TableDef st in srcTables)
            {
                if (!filter(st) || destTblNames.Contains(ctx.GetNewTableName(st)))
                {
                    logger.LogInformation(string.Format("Skipping {0} of {1} table: {2}",
                        ++cnt, total, st.Name));
                    continue;
                }
                logger.LogInformation(string.Format("Linking {0} of {1} table", ++cnt, total));
                var nt = db.LinkToTable(st, externalDb, ctx);
                destTblNames.Add(nt.Name);
            }
        }

        public static TableDef LinkToTable(this Database db, TableDef externalTable, Database externalDb, Context ctx)
        {
            var logger = ctx.Logger;
            var gn = ctx.GetNewTableName;

            logger.LogInformation(string.Format("Linking table: {0}", externalTable.Name));
            var cnnstr = externalDb.GetConnectString();
            var dt = db.CreateTableDef();
            dt.Name = gn(externalTable);
            dt.SourceTableName =
                string.IsNullOrEmpty(externalTable.SourceTableName) ? externalTable.Name : externalTable.SourceTableName;

            if (ctx.IsSavePwdWithLinkTable && cnnstr.Contains("PWD=") && externalTable.Attributes == TBL_ATTR_LOCAL)
                dt.Attributes = (int)TableDefAttributeEnum.dbAttachSavePWD;

            if (externalTable.Attributes == TBL_ATTR_LOCAL)
                dt.Connect = cnnstr;
            else if (externalTable.Attributes == TBL_ATTR_LINK)
                dt.Connect = externalTable.Connect;
            
            db.TableDefs.Append(dt);
            dt.RefreshLink();
            return dt;
        }

        #endregion

        #region Import tables

        public static void ImportTables(this Database db, Database externalDb, Context ctx)
        {
            var logger = ctx.Logger;
            var filter = ctx.TableFilter;

            var srcTables = externalDb.TableDefs;
            var destTblNames = db.TableDefs.GetTableNames();
            int cnt = 0;
            int total = srcTables.Count;
            logger.LogInformation(string.Format("Importing total number of tables: {0}", total));
            foreach (TableDef st in srcTables)
            {
                if (!filter(st) || destTblNames.Contains(ctx.GetNewTableName(st)))
                {
                    logger.LogInformation(string.Format("Skipping {0} of {1} table: {2}", 
                        ++cnt, total, st.Name));
                    continue;
                }
                logger.LogInformation(string.Format("Importing {0} of {1} table", ++cnt, total));
                var nt = db.ImportTable(st, externalDb, ctx);
                destTblNames.Add(nt.Name);
            }
        }

        public static TableDef ImportTable(this Database db, TableDef externalTable, Database externalDb, Context ctx)
        {
            var logger = ctx.Logger;

            var srcCnnString = externalDb.GetConnectString();
            // create table and duplicate data
            var sql = string.Format("select * into [{0}] from [{1}].[{0}]", externalTable.Name, srcCnnString);
            logger?.LogInformation(string.Format("Importing table: {0}", externalTable.Name));
            db.Execute(sql, RecordsetOptionEnum.dbFailOnError);

            db.TableDefs.Refresh();
            var dt = db.TableDefs[externalTable.Name];
            logger?.LogInformation(string.Format("Number of imported records: {0}", dt.RecordCount));

            // duplicate indexes
            logger?.LogInformation(string.Format("Copying indexes for table: {0}", externalTable.Name));
            dt.CopyIndexes(externalTable, ctx);
            return dt;
        }

        public static void CopyIndexes(this TableDef dt, TableDef externalTable, Context ctx)
        {
            var logger = ctx.Logger;

            foreach (DAO.Index si in externalTable.Indexes)
            {
                var fir = si.IndexFields();
                if (fir == null) continue;

                var di = dt.CreateIndex(si.Name);
                logger.LogTrace("Create index: {0}", si.Name);
                dynamic fields = di.Fields;
                while (fir.MoveNext())
                {
                    var sf = (fir.Current as Field)!;
                    var df = di.CreateField(sf.Name);
                    fields.Append(df);
                    logger.LogTrace(string.Format("Append index column: {0}", df.Name));
                }

                di.Primary = si.Primary;
                di.Unique = si.Unique;
                di.IgnoreNulls = si.IgnoreNulls;
                di.Required = si.Required;
                di.Clustered = si.Clustered;
                dt.Indexes.Append(di);
                if (di.Primary)
                    logger.LogInformation(string.Format("Append primary index: {0}", di.Name));
                else
                    logger.LogTrace(string.Format("Append table index: {0}", di.Name));
            }
        }

        #endregion

        public static HashSet<string> GetTableNames(this TableDefs tables)
        {
            var l = new HashSet<string>();
            foreach (TableDef t in tables)
                l.Add(t.Name);
            return l;
        }
    }
}
