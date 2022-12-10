using Dapper;
using Invoicing.model;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Linq;
using System.Threading.Tasks;

namespace Invoicing
{
    public class SQLConnect
    {
        private OleDbConnection oledb = new OleDbConnection(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=../Data.Mdb");
        public SQLConnect()
        {
        }
        public DataTable Find(string sql)
        {
            oledb.Open();
            OleDbDataAdapter db = new OleDbDataAdapter(sql, oledb);
            DataTable dt = new DataTable();
            db.Fill(dt);
            oledb.Close();
            return dt;
        }
        /*public async Task<bool> Add(string sql)
        {
            oledb.Open();
            OleDbCommand db = new OleDbCommand(sql, oledb);
            int i = await db.ExecuteNonQueryAsync();
            oledb.Close();
            return i > 0;
        }
        public async Task<bool> Del(string sql)
        {
            oledb.Open();
            OleDbCommand db = new OleDbCommand(sql, oledb);
            int i = await db.ExecuteNonQueryAsync();
            oledb.Close();
            return i > 0;
        }
        public async Task<bool> Change(string sql)
        {
            oledb.Open();
            OleDbCommand db = new OleDbCommand(sql, oledb);
            int i = await db.ExecuteNonQueryAsync();
            oledb.Close();
            return i > 0;
        }*/
        public async Task InsertList(DataTable DT)
        {
            List<save_list> listName = DT.AsEnumerable().Select(all => new save_list()  //儲存資料庫的欄位要照底下，不然會亂掉
            {
                bili = all.Field<string>("單子編號"),
                cost = all.Field<string>("貨品編號"),
                name_total = all.Field<string>("品名"),
                remark = all.Field<string>("數量"),
                number = all.Field<string>("基本單位"),
                store = all.Field<string>("單價"),
                unit = all.Field<string>("金額"),
                unit_cost = all.Field<string>("備註"),
            }).ToList();
            using (IDbConnection connection = new OleDbConnection(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=../Data.Mdb"))
            {
                string SQL = "insert into 整張儲存 values (@bili, @store, @name_total, @number, @unit, @unit_cost, @cost, @remark)";
                await connection.ExecuteAsync(SQL, listName);
            }
        }
        public async Task InsertPurchase(DataTable DT)
        {
            List<purchase_save> listName = DT.AsEnumerable().Select(all => new purchase_save()  //儲存資料庫的欄位要照底下，不然會亂掉
            {
                bili = all.Field<string>("單子編號"),
                name_total = all.Field<string>("貨品編號"),
                number = all.Field<string>("品名"),
                remark = all.Field<string>("數量"),
                store = all.Field<string>("基本單位"),
                unit = all.Field<string>("備註"),
            }).ToList();
            using (IDbConnection connection = new OleDbConnection(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=../Data.Mdb"))
            {
                string SQL = "insert into 整張儲存 (單子編號, 貨品編號, 品名, 數量, 基本單位, 備註) values (@bili, @store, @name_total, @number, @unit, @remark)";
                await connection.ExecuteAsync(SQL, listName);
            }
        }
        public async Task<DataTable> searchDataTable(string SQL, object corres)
        {
            DataTable DT = new DataTable();
            using (IDbConnection connection = new OleDbConnection(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=../Data.Mdb"))
            {
                var temp = await connection.ExecuteReaderAsync(SQL, corres);
                DT.Load(temp);
            }
            return DT;
        }

        public async Task<bool> execute(string SQL, object corres)
        {
            int i;
            using (IDbConnection connection = new OleDbConnection(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=../Data.Mdb"))
                i = await connection.ExecuteAsync(SQL, corres);
            return i > 0;
        }
    }

}