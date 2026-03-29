using ClosedXML.Excel;
using MySql.Data.MySqlClient;
using System.Data;
using System.Security.Cryptography.X509Certificates;
public class Program
{
    public static void Main()
    {
        int opcao =0;
        Connection conn = new Connection();
        var connection = conn.GetConnection();
        if (connection != null)
        {
            Console.WriteLine("------------------------------------");
            Console.WriteLine("ESTOQUE ITENS DE IMPRESSORAS - HGG");
            Console.WriteLine("------------------------------------");
        }
        else
        {
            Console.WriteLine("Erro");
        }
        
        while (opcao != 4) {
            Console.WriteLine("---------------------------");
            Console.WriteLine("Escolha a opção desejada");
            Console.WriteLine("1 - Listar estoque");
            Console.WriteLine("2 - Alterar quantidade");
            Console.WriteLine("3 - Gerar relatório");
            Console.WriteLine("4 - Sair");
            Console.WriteLine("---------------------------");
            Console.Write("opção:");
            opcao = int.Parse(Console.ReadLine());
            if (opcao == 1)
            {
                Console.WriteLine($"Opção escolhida {opcao}");
                Select(connection);
            }
            else if (opcao == 2)
            {
                Console.WriteLine($"Opção escolhida {opcao}");
                Update(connection);
            }
            else if (opcao == 3)
            {
                Console.WriteLine($"Opção escolhida {opcao}");
                Relatory(connection);
            }
            else
            {
                Console.WriteLine("Operação Finalizada !");
            }
        }
    }
    public static void Select(MySqlConnection connection)
    {
        MySqlCommand cmd = new MySqlCommand("SELECT * FROM itens", connection);
        MySqlDataReader reader = cmd.ExecuteReader();
        while (reader.Read())
        {
            var id = reader["id"];
            var tipo_item = reader["tipo_item"];
            var modelo_impressora = reader["modelo_impressora"];
            int quantidade = Convert.ToInt32(reader["quantidade"]);
            if (quantidade == 5)
            {
                Console.ForegroundColor = ConsoleColor.DarkRed;
            }
            else if(quantidade <= 5)
            {
                Console.ForegroundColor = ConsoleColor.Red;
            }
            else
            {
                Console.ForegroundColor = ConsoleColor.Green;
            }
            Console.WriteLine("--------------------------------------------");
            Console.WriteLine(id+" | "+tipo_item+" | "+modelo_impressora+" | "+quantidade+" | ");
            Console.WriteLine("--------------------------------------------");
        }
        reader.Close();
        Console.ResetColor();
    }
    public static void Update(MySqlConnection connection)
    {
        int id, quantidade;
        Console.WriteLine("Digite o ID");
        id = int.Parse(Console.ReadLine());
        Console.WriteLine("Quantidade:");
        quantidade = int.Parse(Console.ReadLine());
        MySqlCommand cmd = new MySqlCommand("UPDATE itens SET quantidade = " + quantidade + " WHERE id = " + id,connection);
        try
        {
            cmd.ExecuteNonQuery();
            Console.WriteLine("Quantidade alterada com sucesso");
        }
        catch (Exception ex)
        {
            Console.WriteLine(ex.Message);
        }
    }
    public static void Relatory(MySqlConnection connection)
    {
        MySqlCommand cmd = new MySqlCommand("SELECT * FROM itens", connection);
        MySqlDataAdapter da = new MySqlDataAdapter(cmd);
        DataTable dt = new DataTable();
        da.Fill(dt);
        using (var workbook = new XLWorkbook())
        {
            var planilha = workbook.Worksheets.Add("Relatório");
            planilha.Cell(1, 1).InsertTable(dt);
            string caminhoArquivo = Environment.GetFolderPath(Environment.SpecialFolder.Desktop)
                + $"\\Relatorio_EstoqueImpressoras_{DateTime.Now:yyyy-MM-dd_HH-mm-ss}.xlsx";
            workbook.SaveAs(caminhoArquivo);
            Console.WriteLine($"Relatório gerado com sucesso!\n\n{caminhoArquivo}", "Sucesso");
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo()
            {
                FileName = caminhoArquivo,
                UseShellExecute = true
            });
        }
    }
}
public class Connection
{
    string connectionString = "Server=gondola.proxy.rlwy.net;Port=57277;Database=railway;User Id=root;Password=YSJpfpuOKHuqVkcBwhaUCEvyXRJRdJUo;SslMode=Required;";
    public MySqlConnection GetConnection()
    {
        try
        {
            MySqlConnection connection = new MySqlConnection(connectionString);
            connection.Open();
            return connection;
        }
        catch(Exception ex) 
        {
            Console.WriteLine(ex.Message);
            return null;
        }
    }
}