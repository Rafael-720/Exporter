# Database Exporter

O "Database Exporter" é um software desenvolvido em C# que facilita a exportação de dados de um banco de dados PostgreSQL para uma planilha Excel, utilizando conexão via Npgsql. Este software é projetado para exportar dados de três tabelas distintas, organizando-os em abas separadas na planilha, o que facilita a visualização e análise dos dados. Além disso, a aplicação minimiza-se para a bandeja do sistema ao iniciar, permitindo operações contínuas sem interromper o fluxo de trabalho do usuário.

## Funcionalidades

- **Exportação para Excel**: Exporta dados de três tabelas diferentes para abas separadas em um arquivo Excel (.xlsx).
- **Minimização para Bandeja do Sistema**: Permite que o software opere discretamente sem ocupar espaço na barra de tarefas.
- **Monitoramento Contínuo de Alterações**: Utiliza um timer para verificar periodicamente alterações nas tabelas, atualizando os dados exportados conforme necessário.
- **Integração com o OneDrive**: Após a exportação, o software realiza o upload automático do arquivo para o OneDrive, garantindo acesso remoto e seguro aos dados.

## Tecnologias Utilizadas

- **.NET Framework / .NET Core**: Plataforma de desenvolvimento utilizada para construir a aplicação.
- **C# e Windows Forms**: Para a criação da interface gráfica e lógica de programação.
- **Npgsql**: Biblioteca para conexão com o banco de dados PostgreSQL.
- **MiniExcel**: Biblioteca para a criação de arquivos Excel diretamente de datasets.

## Configuração e Execução

1. **Clone o Repositório**:
   ```bash
   git clone https://github.com/seuusuario/seuprojeto.git
   cd seuprojeto
Compilação:
Abra o projeto no Visual Studio e compile-o.

Execução:
Execute o arquivo .exe gerado após a compilação para iniciar o software.

Uso
Após iniciar, o software se minimiza automaticamente para a bandeja do sistema. Interaja com ele através do ícone na bandeja:

Clique Duplo: Restaura a aplicação.
Clique Direito: Abre o menu com opções para fechar a aplicação.
Licença
Este projeto está sob a licença MIT. Veja o arquivo LICENSE.md para detalhes.

Autor:
Rafael Oliveira - [LinkedIn](https://linkedin.com/in/rafael-oliveira720)

Desenvolvido com ❤️ por Rafael para automação e eficiência em processos.
