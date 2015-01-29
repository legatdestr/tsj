using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data.OleDb;


namespace souz.db.AccessImport
    {

    public static class ConnectionStringBuilder
        {
           public static string GetConnectionString(string pathToDbFile)
           {
            if (System.IO.File.Exists(pathToDbFile))
                {
                return 
                @"Provider=Microsoft.Jet.OLEDB.4.0;"
                + @"Data Source="
                + pathToDbFile;
                }
            else
                throw new System.ArgumentException("Указанный файл базы данных не существует: " + pathToDbFile,
                    "pathToDbFile");
            }
        }

    // класс утилита,  содержит методы для импорта
    static class DbImport
        {
        private static DateTime _dt;
        private static string _dbSourcePath;
        private static string _dbTargetPath;

        private static string _spDbFileName = "KVPLS.mdb";
        public static string spDbFileName
            {
            get { return _spDbFileName ; }
            set { _spDbFileName = value;}
            }
        private static string _archiveDbFileName = "archive.mdb";
        public static string archiveDbFileName
            {
            get { return _archiveDbFileName; }
            set { _archiveDbFileName = value; }
            }

        // копирует таблицу из Access в другую базу Access
        public static void ClearAndCopyTableToArchive(string targetTableName, string sourceTablename, OleDbConnection conArchive)
            {
            // SQL для удаления
            string sqlDel = @"DELETE FROM " + targetTableName;
            // SQL для копирования
            string sqlCopy =
            "INSERT INTO [" + targetTableName  + "] SELECT * FROM [MS Access;DATABASE=" +
           _dbSourcePath + @"\" + spDbFileName + ";].[" + sourceTablename + "]";
            try
                {
                if (conArchive.State == System.Data.ConnectionState.Closed)
                    {
                    conArchive.Open();
                    }
                // сначала очищаем
                OleDbCommand command = new OleDbCommand(sqlDel, conArchive);
                command.ExecuteNonQuery();
                // затем копируем
                command.CommandText = sqlCopy;
                command.ExecuteNonQuery();
                }
            catch (System.InvalidOperationException e)
                {
                conArchive.Close();
                throw new Exception("Ошибка загрузки данных в архив! " + e.Message);
                }
            }

        // проверяет существует ли искомая таблица в базе данных
        public static bool IsTableExists(string tableName, OleDbConnection con)
            {
            // Get list of user tables
            if (con.State == System.Data.ConnectionState.Closed)
                {
                con.Open();
                }
            System.Data.DataTable userTables = con.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
            foreach (System.Data.DataRow row in userTables.Rows)
                {
                string strSheetTableName = row["TABLE_NAME"].ToString();
                if (row["TABLE_TYPE"].ToString() == "TABLE")
                    if (strSheetTableName == tableName)
                        {
                        return true;
                        }
                }
            return false;
            }

        public static void DoImport(DateTime dt, string dbSourceDir,
            string dbTargetDir)
            {
            _dt = dt;
            _dbSourcePath = System.IO.Path.GetDirectoryName(dbSourceDir);
            _dbTargetPath = System.IO.Path.GetDirectoryName(dbTargetDir);  

            // --
            // строим объект connection для архива:
            OleDbConnection conArchive = new OleDbConnection(ConnectionStringBuilder.GetConnectionString(_dbTargetPath + @"\" + archiveDbFileName));
            // копируем в архив базу данных справочника организаций (сначала очищается):            
            ClearAndCopyTableToArchive("spZeu", "spZEU", conArchive);
            // копируем в архив базу данных справочника Улиц (сначала очищается)
            ClearAndCopyTableToArchive("spUL", "spUL", conArchive);

          
            // получаем список компаний для того, чотбы сформировать 
            // наименования директорий в которых искать инф-ию по лицевым
            // счетам
            string sqlSelect = "SELECT spZEU.* FROM spZEU";
            OleDbDataAdapter adapterArchiveZeu = new OleDbDataAdapter(sqlSelect, conArchive);
            System.Data.DataTable spZeu = new System.Data.DataTable("spZEU");
            // массив результатов
            System.Data.DataRow[] spZeuRows;
            try
                {
                    adapterArchiveZeu.Fill(spZeu);
                    spZeuRows = spZeu.Select();
                }
            catch (Exception)
                {
                    adapterArchiveZeu.Dispose();
                    spZeu.Clear();
                    spZeu.Dispose();
                    GC.Collect();
                    throw new Exception("Невозможно получить доступ к справочнику компаний.");
                }

            // -- конец получения списка компаний
            
            // Проверяем, есть ли вообще данные по компаниям
            if (spZeuRows.Length < 1)
                {
                conArchive.Close();
                return;
                }

            // Теперь формируем пути к файлам баз данных, проверяем
            List <string> dbPathes = new List<string>();
            // получаем год
            int year = _dt.Year;
            for (int i = 0; i < spZeuRows.Length; i++ )
                {
                // id - идентификатор компании
                string id = spZeuRows[i]["ZEU"].ToString();
                string path = _dbSourcePath + @"\" + @"KVPD\" + "Z" + id + @"\Z" + id + @"_" + year + ".mdb";
                if (System.IO.File.Exists(path))
                    {
                    dbPathes.Add(path);
                    }
                }
            // пути к базе данных сформированы  в dbPathes

            // теперь работаем с архивом Лицевых счетов, 
            // чтобы сделать очистку архива и импорт в него Лицевых счетов 
            
            // сначала необходимо выполнить очистку
            string clearLsSQL = "DELETE FROM LS";
            OleDbCommand clearLsCommand = new OleDbCommand(clearLsSQL,conArchive);
            if (conArchive.State == System.Data.ConnectionState.Closed) { conArchive.Open(); }
            clearLsCommand.ExecuteNonQuery();

            // перебираем все пути
            for (int i = 0; i < dbPathes.Count; i++)
                {
                // строим connection к файлу с базой источником 
                OleDbConnection conSource = new OleDbConnection(ConnectionStringBuilder.GetConnectionString(dbPathes[i]));
                conSource.Open();
                // Теперь необходимо построить имя таблички Лицевых счетов
                string month = _dt.Month.ToString();
                if (month.Length == 1) { month = "0" + month.ToString();}
                string tableName = "LS" + _dt.Year + month;
                // проверить что табличка есть в базе данных
                if (IsTableExists(tableName, conSource))
                    {
                    // табличка существует, значит ее копируем в архив
                    // SQL для копирования
                    string sqlCopy =
                    "INSERT INTO [LS] SELECT * FROM [MS Access;DATABASE=" + dbPathes[i] + ";].[" + tableName + "]";
                    try
                        {
                        if (conArchive.State == System.Data.ConnectionState.Closed)
                            {
                            conArchive.Open();
                            }
                        // копируем
                        OleDbCommand command = new OleDbCommand(sqlCopy, conArchive);
                        command.ExecuteNonQuery();
                        }
                    catch (System.InvalidOperationException e)
                        {
                        conArchive.Close();
                        throw new Exception("Ошибка загрузки данных в архив! " + e.Message);
                        }
                    }
                else // табличка не существует, надо переходить на след. шаг цикла
                    { continue; }
                if (conSource.State == System.Data.ConnectionState.Open)
                    { 
                        conSource.Close(); 
                    }
                }
            if (conArchive.State == System.Data.ConnectionState.Open)
                {
                conArchive.Close();
                }
            } 
        


        }




















    // код устарел в связи с тем, что найден способ копирования таблиц запросом в Access
    //  структура полей
    public static class Utils
        {
        // поле таблицы
        public struct TableField
            {
            public string name;
            public OleDbType type;
            }

        }

    // работа с таблицами - для импорта
    public abstract class AccessTable
        {
        protected Utils.TableField[] tableFields;

        // constructor
        public AccessTable(OleDbConnection con, string tableName)
            {
            this._tableName = tableName;
            this._con = con;
            }

        // получаем ссылку на массив с названиями столбцов
        public Utils.TableField[] getColNames()
            {
            return this.tableFields;
            }
        // получаем названия столбцов в виде строки готовой для употребления в SQL запросе
        public string getColNamesToString()
            {
            string result = null;
            Utils.TableField[] cN = getColNames();
            foreach (Utils.TableField s in cN)
                {
                if (result == null)
                    {
                    result = s.name;
                    }
                else
                    result += "," + s.name;
                }
            return result;

            }
        // получаем поля в виде списка со знаком @ перед каждым названием столбца
        public string getColNamesToStringForBinding()
            {
            string result = null;
            Utils.TableField[] cN = getColNames();
            foreach (Utils.TableField s in cN)
                {
                if (result == null)
                    {
                    result = "@" + s.name;
                    }
                else
                    result += "," + "@" + s.name;
                }
            return result;
            }



        // наименование таблицы с которой связан текущий лицевой счет
        private string _tableName;
        // свойство наименование таблицы
        public string tableName
            {
            set { this._tableName = value; }
            get { return this._tableName; }
            }
        // connection Object
        private OleDbConnection _con;
        // connection property
        public OleDbConnection connection
            {
            get { return this._con; }
            }

        // копируем базу лицевых счетов в общий архив счетов
        // сливаем из подключенной базы табличку со счетами в архив
        public bool CopyMeToArchive(AccessTable targetTable)
            {

            string select = "SELECT " + getColNamesToString() + " FROM " + tableName;
            OleDbCommand selCommand = new OleDbCommand(select, connection);

            try
                {
                try
                    {
                    connection.Open();
                    }
                catch (InvalidOperationException)
                    {

                    }
                // запрос на выборку данных
                OleDbDataReader reader = selCommand.ExecuteReader();
                if (reader.HasRows == false) { return false; }

                string insert = @"INSERT INTO " + targetTable.tableName +
                    "(" + getColNamesToString() + ") VALUES ("
                    + getColNamesToStringForBinding() + ")";
                // создаем командный объект для вставки данных
                OleDbCommand insCommand = new OleDbCommand(insert, targetTable.connection);
                // а теперь добавляем параметры в запрос
                for (int i = 0; i < getColNames().Length; i++)
                    {
                    insCommand.Parameters.Add(@"@" + getColNames()[i].name, getColNames()[i].type);
                    }
                // а теперь вставляем данные
                targetTable.connection.Open();
                while (reader.Read())
                    {
                    // опять перебираем все параметры\столбцы
                    for (int i = 0; i < getColNames().Length; i++)
                        {
                        insCommand.Parameters["@" + getColNames()[i].name].Value = reader[getColNames()[i].name];
                        //    insCommand.Parameters["@" + getColNames()[i].name].Size = reader[getColNames()[i].name].ToString().Length;
                        }

                    insCommand.ExecuteNonQuery();
                    }
                targetTable.connection.Close();
                }
            catch (Exception e)
                {
                connection.Close();
                targetTable.connection.Close();
                throw new Exception(@"Ошибка чтения\записи данных из источника данных в архив. " + e.Message);

                }


            connection.Close();
            return true;
            }

        // очищает таблицу (удаляет все записи)
        public bool ClearTable()
            {
            OleDbCommand deleteCommand = new OleDbCommand("DELETE FROM " + this.tableName, this.connection);

            try
                {
                this.connection.Open();
                deleteCommand.ExecuteNonQuery();
                }
            catch (Exception)
                {
                this.connection.Close();
                throw new Exception("Невозможно выполнить очищение архивной таблицы лицевых счетов");
                }

            this.connection.Close();
            return false;
            }

        }

    // Лицевые счета -  для импорта
    public class Ls : AccessTable
        {
        public Ls(OleDbConnection con, string tableName)
            : base(con, tableName)
            {
            tableFields = new Utils.TableField[]
        {
            new Utils.TableField(){name = "FIO", type= OleDbType.VarChar},
            new Utils.TableField(){name = "UL", type= OleDbType.Integer},
            new Utils.TableField(){name = "DOM", type= OleDbType.Integer},
            new Utils.TableField(){name = "DOML", type= OleDbType.VarChar},
            new Utils.TableField(){name = "DOMP", type= OleDbType.Integer},
            new Utils.TableField(){name = "KV", type= OleDbType.Integer},
            new Utils.TableField(){name = "KVL", type= OleDbType.VarChar},
            new Utils.TableField(){name = "LS", type= OleDbType.Integer},
            new Utils.TableField(){name = "NZEU", type= OleDbType.Integer},

            new Utils.TableField(){name = "LSX", type= OleDbType.Integer},
            new Utils.TableField(){name = "TIP", type= OleDbType.Integer},
            new Utils.TableField(){name = "OTYPE", type= OleDbType.Integer},
            new Utils.TableField(){name = "KOLL", type= OleDbType.Integer},
            new Utils.TableField(){name = "KOLP", type= OleDbType.Integer},
            new Utils.TableField(){name = "KOLW", type= OleDbType.Integer},
            new Utils.TableField(){name = "USLUG", type= OleDbType.VarChar},

            new Utils.TableField(){name = "PLO", type= OleDbType.Integer},
            new Utils.TableField(){name = "PLZ",type= OleDbType.Integer},
            new Utils.TableField(){name = "PLD", type= OleDbType.Integer},
            new Utils.TableField(){name = "REZ",type= OleDbType.Integer},
            new Utils.TableField(){name = "ITGN1", type= OleDbType.Currency},
            new Utils.TableField(){name = "SAL1W", type= OleDbType.Currency},
            new Utils.TableField(){name = "STE1", type= OleDbType.Integer},
            new Utils.TableField(){name = "STE2", type= OleDbType.Integer},

            new Utils.TableField(){name = "STX1", type= OleDbType.Integer},
            new Utils.TableField(){name = "STX2",type= OleDbType.Integer},
            new Utils.TableField(){name = "STX3", type= OleDbType.Integer},
            new Utils.TableField(){name = "STX4", type= OleDbType.Integer},
            new Utils.TableField(){name = "STG1", type= OleDbType.Integer},
            new Utils.TableField(){name = "STG2", type= OleDbType.Integer},
            new Utils.TableField(){name = "STG3", type= OleDbType.Integer},
            new Utils.TableField(){name = "STG4", type= OleDbType.Integer},

            new Utils.TableField(){name = "KVS", type= OleDbType.Currency},
            new Utils.TableField(){name = "UBP", type= OleDbType.Currency},
            new Utils.TableField(){name = "OPU", type= OleDbType.Currency},
            new Utils.TableField(){name = "TRM", type= OleDbType.Currency},
            new Utils.TableField(){name = "RAD", type= OleDbType.Currency},
            new Utils.TableField(){name = "SLI", type= OleDbType.Currency},
            new Utils.TableField(){name = "KAN", type= OleDbType.Currency},
            new Utils.TableField(){name = "ANT", type= OleDbType.Currency},
            new Utils.TableField(){name = "WMU", type= OleDbType.Currency},

            new Utils.TableField(){name = "LIF", type= OleDbType.Currency},
            new Utils.TableField(){name = "KRM", type= OleDbType.Currency},
            new Utils.TableField(){name = "NAI", type= OleDbType.Currency},
            new Utils.TableField(){name = "ELE", type= OleDbType.Currency},
            new Utils.TableField(){name = "OTO", type= OleDbType.Currency},
            new Utils.TableField(){name = "XWO", type= OleDbType.Currency},
            new Utils.TableField(){name = "GWO", type= OleDbType.Currency},
            new Utils.TableField(){name = "XGW", type= OleDbType.Currency},
            new Utils.TableField(){name = "ODN", type= OleDbType.Currency},

            new Utils.TableField(){name = "APP", type= OleDbType.Currency},
            new Utils.TableField(){name = "VCU", type= OleDbType.Currency},
            new Utils.TableField(){name = "PRC", type= OleDbType.Currency},
            new Utils.TableField(){name = "CSB", type= OleDbType.Currency},
            new Utils.TableField(){name = "SOO", type= OleDbType.Currency},
            new Utils.TableField(){name = "ITGNW", type= OleDbType.Currency},
            new Utils.TableField(){name = "KVS_S1", type= OleDbType.Currency},
            new Utils.TableField(){name = "UBP_S1", type= OleDbType.Currency},

            new Utils.TableField(){name = "OPU_S1",  type= OleDbType.Currency},
            new Utils.TableField(){name = "TRM_S1", type= OleDbType.Currency},
            new Utils.TableField(){name = "RAD_S1", type= OleDbType.Currency},
            new Utils.TableField(){name = "SLI_S1", type= OleDbType.Currency},
            new Utils.TableField(){name = "KAN_S1", type= OleDbType.Currency},
            new Utils.TableField(){name = "ANT_S1", type= OleDbType.Currency},

            new Utils.TableField(){name = "WMU_S1",  type= OleDbType.Currency},
            new Utils.TableField(){name = "LIF_S1", type= OleDbType.Currency},
            new Utils.TableField(){name = "KRM_S1", type= OleDbType.Currency},
            new Utils.TableField(){name = "NAI_S1", type= OleDbType.Currency},
            new Utils.TableField(){name = "ELE_S1", type= OleDbType.Currency},
            new Utils.TableField(){name = "OTO_S1", type= OleDbType.Currency},

            new Utils.TableField(){name = "XWO_S1",  type= OleDbType.Currency},
            new Utils.TableField(){name = "GWO_S1", type= OleDbType.Currency},
            new Utils.TableField(){name = "XGW_S1", type= OleDbType.Currency},
            new Utils.TableField(){name = "ODN_S1", type= OleDbType.Currency},
            new Utils.TableField(){name = "APP_S1", type= OleDbType.Currency},
            new Utils.TableField(){name = "VCU_S1", type= OleDbType.Currency},

            new Utils.TableField(){name = "PRC_S1",  type= OleDbType.Currency},
            new Utils.TableField(){name = "CSB_S1", type= OleDbType.Currency},
            new Utils.TableField(){name = "SOO_S1", type= OleDbType.Currency},
            new Utils.TableField(){name = "PEN_S1", type= OleDbType.Currency},
            new Utils.TableField(){name = "ITGLW", type= OleDbType.Currency},
            new Utils.TableField(){name = "ITGLF", type= OleDbType.Currency},

            new Utils.TableField(){name = "KVS_P",  type= OleDbType.Currency},
            new Utils.TableField(){name = "UBP_P", type= OleDbType.Currency},
            new Utils.TableField(){name = "OPU_P", type= OleDbType.Currency},
            new Utils.TableField(){name = "TRM_P", type= OleDbType.Currency},
            new Utils.TableField(){name = "RAD_P", type= OleDbType.Currency},
            new Utils.TableField(){name = "SLI_P", type= OleDbType.Currency},
            new Utils.TableField(){name = "KAN_P", type= OleDbType.Currency},

            new Utils.TableField(){name = "ANT_P",  type= OleDbType.Currency},
            new Utils.TableField(){name = "WMU_P", type= OleDbType.Currency},
            new Utils.TableField(){name = "LIF_P", type= OleDbType.Currency},
            new Utils.TableField(){name = "KRM_P", type= OleDbType.Currency},
            new Utils.TableField(){name = "NAI_P", type= OleDbType.Currency},
            new Utils.TableField(){name = "ELE_P", type= OleDbType.Currency},
            new Utils.TableField(){name = "OTO_P", type= OleDbType.Currency},

            new Utils.TableField(){name = "XWO_P",  type= OleDbType.Currency},
            new Utils.TableField(){name = "GWO_P", type= OleDbType.Currency},
            new Utils.TableField(){name = "XGW_P", type= OleDbType.Currency},
            new Utils.TableField(){name = "ODN_P", type= OleDbType.Currency},
            new Utils.TableField(){name = "APP_P", type= OleDbType.Currency},
            new Utils.TableField(){name = "VCU_P", type= OleDbType.Currency},

            new Utils.TableField(){name = "PRC_P",  type= OleDbType.Currency},
            new Utils.TableField(){name = "CSB_P", type= OleDbType.Currency},
            new Utils.TableField(){name = "SOO_P", type= OleDbType.Currency},
            new Utils.TableField(){name = "ITGNP", type= OleDbType.Currency},
            new Utils.TableField(){name = "ITGPW", type= OleDbType.Currency},
            new Utils.TableField(){name = "ITGSW", type= OleDbType.Currency},
            new Utils.TableField(){name = "CSB_SA", type= OleDbType.Currency},

            new Utils.TableField(){name = "SAL2W",  type= OleDbType.Currency},
            new Utils.TableField(){name = "DOLGW", type= OleDbType.Currency},


            new Utils.TableField(){name = "KDOL", type= OleDbType.Integer},
            new Utils.TableField(){name = "DOG", type= OleDbType.Integer},
            new Utils.TableField(){name = "DATDOG1", type= OleDbType.DBTimeStamp},
            new Utils.TableField(){name = "DATDOG2", type= OleDbType.DBTimeStamp},
            new Utils.TableField(){name = "PEN", type= OleDbType.Currency},

            new Utils.TableField(){name = "PENPW",  type= OleDbType.Currency},
            new Utils.TableField(){name = "PENSW", type= OleDbType.Currency},
            new Utils.TableField(){name = "ETAZ", type= OleDbType.Integer},
            new Utils.TableField(){name = "NPOD", type= OleDbType.Integer},
            new Utils.TableField(){name = "ROOMS", type= OleDbType.Integer},
            new Utils.TableField(){name = "TIPK", type= OleDbType.Integer},
            new Utils.TableField(){name = "MAST", type= OleDbType.Integer},
            new Utils.TableField(){name = "TELD", type= OleDbType.Integer},

            new Utils.TableField(){name = "TELS",  type= OleDbType.VarChar},
            new Utils.TableField(){name = "MES", type= OleDbType.Integer},
            new Utils.TableField(){name = "DNI", type= OleDbType.Integer},
            new Utils.TableField(){name = "FROM_DC", type= OleDbType.VarChar},
            new Utils.TableField(){name = "DATPK", type= OleDbType.DBTimeStamp},
            new Utils.TableField(){name = "KVS_LGn", type= OleDbType.Integer},

            new Utils.TableField(){name = "KVU", type= OleDbType.Currency},
            new Utils.TableField(){name = "UB1", type= OleDbType.Currency},

            new Utils.TableField(){name = "OP1",  type= OleDbType.Currency},
            new Utils.TableField(){name = "TR1", type= OleDbType.Currency},
            new Utils.TableField(){name = "RA", type= OleDbType.Currency},
            new Utils.TableField(){name = "SL", type= OleDbType.Currency},
            new Utils.TableField(){name = "AN", type= OleDbType.Currency},
            new Utils.TableField(){name = "WM", type= OleDbType.Currency},
            new Utils.TableField(){name = "LI", type= OleDbType.Currency},
            new Utils.TableField(){name = "KR", type= OleDbType.Currency},
            new Utils.TableField(){name = "NA", type= OleDbType.Currency},
            new Utils.TableField(){name = "EL", type= OleDbType.Currency},
            new Utils.TableField(){name = "OT", type= OleDbType.Currency},
            new Utils.TableField(){name = "XW", type= OleDbType.Currency},
            new Utils.TableField(){name = "GW", type= OleDbType.Currency},

            new Utils.TableField(){name = "XG",  type= OleDbType.Currency},
            new Utils.TableField(){name = "KA", type= OleDbType.Currency},
            new Utils.TableField(){name = "OD", type= OleDbType.Currency},
            new Utils.TableField(){name = "AP", type= OleDbType.Currency},
            new Utils.TableField(){name = "VC", type= OleDbType.Currency},
            new Utils.TableField(){name = "PR", type= OleDbType.Currency},
            new Utils.TableField(){name = "CS", type= OleDbType.Currency},
            new Utils.TableField(){name = "SO", type= OleDbType.Currency},

            new Utils.TableField(){name = "NDOM", type= OleDbType.Integer},
            new Utils.TableField(){name = "NSCH", type= OleDbType.Integer},
            new Utils.TableField(){name = "FZ", type= OleDbType.Integer}

        };
            }

        }

    // справочник улиц - для импорта
    public class SpUl : AccessTable
        {
        // constructor. Calls the base constructor
        public SpUl(OleDbConnection con, string tableName)
            : base(con, tableName)
            {
            tableFields = new Utils.TableField[]
        {
            new Utils.TableField(){name = "UL", type= OleDbType.Integer},
            new Utils.TableField(){name = "NAIM", type= OleDbType.VarChar},
            new Utils.TableField(){name = "Kod", type= OleDbType.Integer}

        };
            }
        }

    // справочник организаций - для импорта
    public class SpZeu : AccessTable
        {
        // constructor. Calls the base constructor
        public SpZeu(OleDbConnection con, string tableName)
            : base(con, tableName)
            {

            tableFields = new Utils.TableField[]
        {
            new Utils.TableField(){name = "ZEU", type= OleDbType.Integer},
            new Utils.TableField(){name = "NAIM", type= OleDbType.Integer},
            new Utils.TableField(){name = "DIRZ", type= OleDbType.Integer},
            new Utils.TableField(){name = "DIRZ1", type= OleDbType.Integer},
            new Utils.TableField(){name = "OTW", type= OleDbType.Integer},
            new Utils.TableField(){name = "TEL_ZEU", type= OleDbType.Integer},
            new Utils.TableField(){name = "РсчГод", type= OleDbType.Integer},
            new Utils.TableField(){name = "РсчМес", type= OleDbType.Integer},
            new Utils.TableField(){name = "РабГод", type= OleDbType.Integer},
            new Utils.TableField(){name = "РабМес", type= OleDbType.Integer},
            new Utils.TableField(){name = "DM_all", type= OleDbType.Integer},
            new Utils.TableField(){name = "DM_on", type= OleDbType.Integer},
            new Utils.TableField(){name = "DM_prv", type= OleDbType.Integer},
            new Utils.TableField(){name = "DM_mun", type= OleDbType.Integer},
            new Utils.TableField(){name = "DM_zil", type= OleDbType.Integer},
            new Utils.TableField(){name = "LS_all", type= OleDbType.Integer},
            new Utils.TableField(){name = "LS_on", type= OleDbType.Integer},
            new Utils.TableField(){name = "LS_prv", type= OleDbType.Integer},
            new Utils.TableField(){name = "LS_mun", type= OleDbType.Integer},
            new Utils.TableField(){name = "LS_zil", type= OleDbType.Integer},
            new Utils.TableField(){name = "PLO_all", type= OleDbType.Integer},
            new Utils.TableField(){name = "PLO_on", type= OleDbType.Integer},
            new Utils.TableField(){name = "PLO_prv", type= OleDbType.Integer},
            new Utils.TableField(){name = "PLO_mun", type= OleDbType.Integer},
            new Utils.TableField(){name = "PLO_zil", type= OleDbType.Integer},
            new Utils.TableField(){name = "KOLP_all", type= OleDbType.Integer},
            new Utils.TableField(){name = "KOLP_on", type= OleDbType.Integer},
            new Utils.TableField(){name = "KOLP_prv", type= OleDbType.Integer},
            new Utils.TableField(){name = "KOLP_mun", type= OleDbType.Integer},
            new Utils.TableField(){name = "KOLP_zil", type= OleDbType.Integer},
            new Utils.TableField(){name = "SAL1W_all", type= OleDbType.Integer},
            new Utils.TableField(){name = "SAL1W_on", type= OleDbType.Integer},
            new Utils.TableField(){name = "SAL1W_prv", type= OleDbType.Integer},
            new Utils.TableField(){name = "SAL1W_mun", type= OleDbType.Integer},
            new Utils.TableField(){name = "SAL1W_zil", type= OleDbType.Integer},
            new Utils.TableField(){name = "ITGNP_all", type= OleDbType.Integer},
            new Utils.TableField(){name = "ITGNP_on", type= OleDbType.Integer},
            new Utils.TableField(){name = "ITGNP_prv", type= OleDbType.Integer},
            new Utils.TableField(){name = "ITGNP_mun", type= OleDbType.Integer},
            new Utils.TableField(){name = "ITGNP_zil", type= OleDbType.Integer},
            new Utils.TableField(){name = "ITGSW_kol", type= OleDbType.Integer},
            new Utils.TableField(){name = "ITGSW_all", type= OleDbType.Integer},
            new Utils.TableField(){name = "ITGSW_on", type= OleDbType.Integer},
            new Utils.TableField(){name = "ITGSW_prv", type= OleDbType.Integer},
            new Utils.TableField(){name = "ITGSW_mun", type= OleDbType.Integer},
            new Utils.TableField(){name = "ITGSW_zil", type= OleDbType.Integer},
            new Utils.TableField(){name = "ITGPW_kol", type= OleDbType.Integer},
            new Utils.TableField(){name = "ITGPW_all", type= OleDbType.Integer},
            new Utils.TableField(){name = "ITGPW_on", type= OleDbType.Integer},
            new Utils.TableField(){name = "ITGPW_prv", type= OleDbType.Integer},
            new Utils.TableField(){name = "ITGPW_mun", type= OleDbType.Integer},
            new Utils.TableField(){name = "ITGPW_zil", type= OleDbType.Integer},
            new Utils.TableField(){name = "PEN_S1_kol", type= OleDbType.Integer},
            new Utils.TableField(){name = "PEN_S1_all", type= OleDbType.Integer},
            new Utils.TableField(){name = "PEN_S1_on", type= OleDbType.Integer},
            new Utils.TableField(){name = "PEN_S1_prv", type= OleDbType.Integer},
            new Utils.TableField(){name = "PEN_S1_mun", type= OleDbType.Integer},
            new Utils.TableField(){name = "PEN_S1_zil", type= OleDbType.Integer},
            new Utils.TableField(){name = "PEN_kol", type= OleDbType.Integer},
            new Utils.TableField(){name = "PEN_all", type= OleDbType.Integer},
            new Utils.TableField(){name = "PEN_on", type= OleDbType.Integer},
            new Utils.TableField(){name = "PEN_prv", type= OleDbType.Integer},
            new Utils.TableField(){name = "PEN_mun", type= OleDbType.Integer},
            new Utils.TableField(){name = "PEN_zil", type= OleDbType.Integer},
            new Utils.TableField(){name = "DOLGmF", type= OleDbType.Integer},
            new Utils.TableField(){name = "DOLGmK", type= OleDbType.Integer},
            new Utils.TableField(){name = "DOLGsF", type= OleDbType.Integer},
            new Utils.TableField(){name = "DOLGsK", type= OleDbType.Integer},
            new Utils.TableField(){name = "PENsF", type= OleDbType.Integer},
            new Utils.TableField(){name = "PENsK", type= OleDbType.Integer},
            new Utils.TableField(){name = "PEN_S1sF", type= OleDbType.Integer},
            new Utils.TableField(){name = "PEN_S1sK", type= OleDbType.Integer},
            new Utils.TableField(){name = "PLOsF", type= OleDbType.Integer},
            new Utils.TableField(){name = "PLOsK", type= OleDbType.Integer},
            new Utils.TableField(){name = "KOLPsF", type= OleDbType.Integer},
            new Utils.TableField(){name = "KOLPsK", type= OleDbType.Integer},
            new Utils.TableField(){name = "DATS", type= OleDbType.Integer},
            new Utils.TableField(){name = "DATP", type= OleDbType.Integer},
            new Utils.TableField(){name = "DATPK", type= OleDbType.Integer},
            new Utils.TableField(){name = "UK", type= OleDbType.Integer},
            new Utils.TableField(){name = "BILL_SB", type= OleDbType.Integer},
            new Utils.TableField(){name = "BILL_Sp", type= OleDbType.Integer},
            new Utils.TableField(){name = "RS", type= OleDbType.Integer},
            new Utils.TableField(){name = "KviS", type= OleDbType.Integer},
            new Utils.TableField(){name = "Дисп", type= OleDbType.Integer},
            new Utils.TableField(){name = "БухПасп", type= OleDbType.Integer},
            new Utils.TableField(){name = "Адрес", type= OleDbType.Integer},
            new Utils.TableField(){name = "ОсобОтм", type= OleDbType.Integer},
            new Utils.TableField(){name = "USLUG", type= OleDbType.Integer},
            new Utils.TableField(){name = "DOG_SubsN", type= OleDbType.Integer},
            new Utils.TableField(){name = "DOG_SubsD", type= OleDbType.Integer},
            new Utils.TableField(){name = "DAT_Subs", type= OleDbType.Integer},
            new Utils.TableField(){name = "DAT_TWK", type= OleDbType.Integer},
            new Utils.TableField(){name = "DAT_KVIT", type= OleDbType.Integer},
            new Utils.TableField(){name = "DAT_Post", type= OleDbType.Integer},
            new Utils.TableField(){name = "DAT_Bill", type= OleDbType.Integer},
            new Utils.TableField(){name = "DAT_BEG", type= OleDbType.Integer},
            new Utils.TableField(){name = "DAT_END", type= OleDbType.Integer},
            new Utils.TableField(){name = "DAT_Open", type= OleDbType.Integer},
            new Utils.TableField(){name = "s_sfilt", type= OleDbType.Integer},
            new Utils.TableField(){name = "p_sfilt", type= OleDbType.Integer},
            new Utils.TableField(){name = "COMM", type= OleDbType.Integer},
            new Utils.TableField(){name = "ZEU_SoC", type= OleDbType.Integer},
            new Utils.TableField(){name = "VC_P", type= OleDbType.Integer},
            new Utils.TableField(){name = "DATsp", type= OleDbType.Integer}
        };

            }

        } // end of class




    }
