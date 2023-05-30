using Npgsql;
using System.Data;
using System.Data.Common;
using System.Numerics;
using MiniExcelLibs;
using System.Windows.Forms;

namespace Exporter
{
    public partial class Form1 : Form
    {
        private System.Windows.Forms.Timer timer;
        private string query, query2, query3;
        string diretorioOneDrive;
        private DataSet tabelaAnterior;
        string filePath;

        public Form1()
        {
            InitializeComponent();
            InitializeTimer();

            this.MaximizeBox = false;
            Exportar();

            // Minimizar para a bandeja do sistema ao iniciar o programa
            this.WindowState = FormWindowState.Minimized;
            this.ShowInTaskbar = false;
            notifyIcon.Visible = true;
        }

        private void Form1_Resize(object sender, EventArgs e)
        {
            if (FormWindowState.Minimized == this.WindowState)
            {
                //this.Hide();
                this.WindowState = FormWindowState.Minimized;
                this.ShowInTaskbar = false;
                notifyIcon.Visible = true;
            }
        }

        //private void notifyIcon_DoubleClick(object sender, EventArgs e)
        //{
        //    this.Show();
        //    this.WindowState = FormWindowState.Normal;
        //    notifyIcon.Visible = false;
        //    //this.ShowInTaskbar = true;
        //}

        private void fecharToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Close();
            //Application.Exit();
        }

        private void notifyIcon_MouseClick(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                contextMenuStrip.Show(Control.MousePosition);
            }
        }

        private void InitializeTimer()
        {
            timer = new System.Windows.Forms.Timer();
            timer.Interval = 30000;
            timer.Tick += Timer_Tick;
        }

        private void Timer_Tick(object? sender, EventArgs e)
        {
            Exportar();
        }

        private bool IsDataSetEqual(DataSet dataSet, DataSet dataSetAnterior)
        {
            // Verificar se os DataSets têm a mesma estrutura
            if (!AreDataSetsStructureEqual(dataSet, dataSetAnterior))
                return false;

            // Verificar se os dados são iguais
            foreach (DataTable tabela in dataSet.Tables)
            {
                if (!dataSetAnterior.Tables.Contains(tabela.TableName))
                    return false;

                DataTable tabelaAnterior = dataSetAnterior.Tables[tabela.TableName];

                if (!AreDataTablesEqual(tabela, tabelaAnterior))
                    return false;
            }

            return true;
        }

        private bool AreDataSetsStructureEqual(DataSet dataSet, DataSet dataSetAnterior)
        {
            if (dataSet.Tables.Count != dataSetAnterior.Tables.Count)
                return false;

            for (int i = 0; i < dataSet.Tables.Count; i++)
            {
                if (dataSet.Tables[i].TableName != dataSetAnterior.Tables[i].TableName)
                    return false;

                if (!AreDataTablesStructureEqual(dataSet.Tables[i], dataSetAnterior.Tables[i]))
                    return false;
            }

            return true;
        }

        private bool AreDataTablesEqual(DataTable tabela, DataTable tabelaAnterior)
        {
            // Verificar se as tabelas têm a mesma estrutura
            if (!AreDataTablesStructureEqual(tabela, tabelaAnterior))
                return false;

            // Verificar se os dados são iguais
            int rowCount = Math.Min(tabela.Rows.Count, tabelaAnterior.Rows.Count); // Usar o menor número de linhas entre as tabelas
            for (int i = 0; i < rowCount; i++)
            {
                for (int j = 0; j < tabela.Columns.Count; j++)
                {
                    if (!tabela.Rows[i][j].Equals(tabelaAnterior.Rows[i][j]))
                        return false;
                }
            }

            // Verificar se o número total de linhas é igual
            if (tabela.Rows.Count != tabelaAnterior.Rows.Count)
                return false;

            return true;
        }

        private bool AreDataTablesStructureEqual(DataTable tabela, DataTable tabelaAnterior)
        {
            if (tabela.Columns.Count != tabelaAnterior.Columns.Count)
                return false;

            for (int i = 0; i < tabela.Columns.Count; i++)
            {
                if (tabela.Columns[i].DataType != tabelaAnterior.Columns[i].DataType ||
                    tabela.Columns[i].ColumnName != tabelaAnterior.Columns[i].ColumnName)
                {
                    return false;
                }
            }

            return true;
        }


        private void btExportar_Click(object sender, EventArgs e)
        {
            Exportar();
            //timer.Start();

        }

        private void Exportar()
        {
            //parametro_bi  query  conexao
            query = "select * from parametro_bi;";
            query2 = "select * from conexao;";
            query3 = "select * from query_producao qp;";

            diretorioOneDrive = GetOneDriveFolderPath();

            if (!string.IsNullOrEmpty(diretorioOneDrive))
            {
                filePath = Path.Combine(diretorioOneDrive, "Arquivo_Controle.xlsx");

                Conexao conexao = new Conexao(query, query2, query3);

                try
                {

                    conexao.Abrir();

                    //DataTable tabela = new DataTable();
                    DataSet tabela = new DataSet();
                    conexao.GetAdapter().Fill(tabela, "Parametro_BI");
                    conexao.GetAdapter2().Fill(tabela, "Conexao");
                    conexao.GetAdapter3().Fill(tabela, "Query");


                    if (File.Exists(filePath))
                    {
                        if (tabelaAnterior != null)
                        {
                            // Verificar diferenças entre os resultados
                            if (!IsDataSetEqual(tabela, tabelaAnterior))
                            {


                                File.Delete(filePath);


                                // Salvar os dados em um arquivo Excel
                                MiniExcel.SaveAs(filePath, tabela);
                                tabelaAnterior = tabela;
                            }
                        }
                    }
                    else
                    {
                        MiniExcel.SaveAs(filePath, tabela);
                        tabelaAnterior = tabela;
                    }




                    //MessageBox.Show("Consulta concluída e resultado salvo no arquivo Excel: " + diretorioOneDrive);

                    timer.Start();

                }
                catch (Exception ex)
                {
                    MessageBox.Show("Erro ao conectar, buscar dados ou salvar no arquivo Excel: " + ex.Message);
                }
                finally
                {
                    conexao.Dispose();
                }
            }
            else
            {
                MessageBox.Show("Diretorio do OneDrive não encontrado. Verifique se o OneDrive está instalado e configurado.");
                return;
            }
        }

        private string GetOneDriveFolderPath()
        {
            string diretorioOneDrive = Environment.GetEnvironmentVariable("OneDrive");

            if (!string.IsNullOrEmpty(diretorioOneDrive))
            {
                // Se a variável de ambiente existir, substituir as ocorrências de "%OneDrive%" pelo caminho real
                diretorioOneDrive = diretorioOneDrive.Replace("%OneDrive%", Environment.GetFolderPath(Environment.SpecialFolder.UserProfile));
            }

            return diretorioOneDrive;
        }

        private void notifyIcon_DoubleClick(object sender, MouseEventArgs e)
        {
            this.Show();
            this.WindowState = FormWindowState.Normal;
            notifyIcon.Visible = false;
            this.ShowInTaskbar = true;
        }

    }
}
