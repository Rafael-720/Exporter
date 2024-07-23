using Npgsql;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Exporter
{
    public class Conexao : IDisposable
    {
        private string connectionString = "";
        private NpgsqlConnection connection;
        private NpgsqlCommand command, cmd2, cmd3;
        private NpgsqlDataReader reader;
        string query, query2, query3;
        private NpgsqlDataAdapter adaptador, adaptador2, adaptador3;



        public Conexao(string query, string query2, string query3)
        {
            this.connection = new NpgsqlConnection(connectionString);
            this.query = query;
            this.query2 = query2;
            this.query3 = query3;
            this.command = new NpgsqlCommand(this.query, this.connection);
            this.cmd2 = new NpgsqlCommand(this.query2, this.connection);
            this.cmd3 = new NpgsqlCommand(this.query3, this.connection);
            this.adaptador = new NpgsqlDataAdapter(command);
            this.adaptador2 = new NpgsqlDataAdapter(cmd2);
            this.adaptador3 = new NpgsqlDataAdapter(cmd3);
        }

        public void Abrir()
        {
            connection.Open();
            //MessageBox.Show("conexao aberta");
        }

        public NpgsqlDataReader GetDataReader()
        {
            return command.ExecuteReader();
        }

        public NpgsqlDataAdapter GetAdapter()
        {
            return adaptador;
        }

        public NpgsqlDataAdapter GetAdapter2()
        {
            return adaptador2;
        }

        public NpgsqlDataAdapter GetAdapter3()
        {
            return adaptador3;
        }

        public void Dispose()
        {
            if (reader != null && !reader.IsClosed)
            {
                reader.Close();
            }

            if (connection != null && connection.State != System.Data.ConnectionState.Closed)
            {
                connection.Close();
            }
            //MessageBox.Show("conexao fechada");
        }


    }
}
