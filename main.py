# Importa do módulo db (arquivo db.py) as funções “conectar” e “criar_banco_e_tabela”.
# - conectar → abre uma conexão com o servidor MySQL.
# - criar_banco_e_tabela → garante que o banco e a tabela necessários existam, criando-os se preciso.
# Essas funções são essenciais para preparar o ambiente do banco de
#                dados antes que a interface gráfica seja exibida.
from db import conectar, criar_banco_e_tabela

# Importa do módulo app (arquivo app.py) a classe “AppCadastro”.
# Essa classe representa toda a interface gráfica (GUI) do sistema feita em Tkinter.
# Ela contém os botões, campos de texto, tabelas e toda a lógica de interação do usuário.
from app import AppCadastro

# A condição abaixo verifica se o arquivo atual está sendo executado
#                diretamente (e não importado por outro módulo).
# "__name__" é uma variável especial do Python:
#   - Quando o arquivo é executado diretamente, __name__ == "__main__".
#   - Quando é importado por outro arquivo, __name__ contém o nome do módulo.
# Assim, o código dentro deste bloco só será executado quando o
#                programa for iniciado diretamente pelo usuário.
if __name__ == "__main__":

    # Antes de abrir a interface gráfica (Tkinter), o sistema garante que o
    #                banco de dados e a tabela existam.
    # Isso evita erros posteriores, por exemplo, quando o usuário tenta
    #                cadastrar algo mas o banco ainda não existe.
    #
    # O bloco try/except é usado para capturar erros de conexão ou criação do banco
    #                (como servidor desligado, senha incorreta, etc.) e evitar que o
    #                programa quebre de forma abrupta.
    try:

        # Chama a função conectar, mas com o parâmetro usar_banco=False.
        # Isso significa que a conexão será feita **somente ao servidor MySQL**,
        #                sem escolher ainda um banco.
        # Essa etapa é necessária, pois o banco pode não existir ainda — e
        #                será criado logo em seguida.
        c = conectar(usar_banco=False)

        # Chama a função “criar_banco_e_tabela”, passando a conexão aberta.
        # Essa função cria o banco de dados (schema) e a tabela principal (pessoas),
        #                caso ainda não existam.
        criar_banco_e_tabela(c)

        # Fecha a conexão temporária usada apenas para criar o banco e a tabela.
        # Isso libera recursos do servidor MySQL antes de abrir a conexão principal usada pela aplicação.
        c.close()

    # Caso aconteça qualquer erro durante essa etapa de conexão ou criação do banco,
    #                o programa não interrompe a execução.
    # Em vez disso, o except captura a exceção silenciosamente.
    # O próprio PyMySQL imprime a mensagem de erro no console, o que ajuda na depuração.
    except Exception:

        # “pass” significa literalmente “não faça nada” — ele ignora o erro e segue a execução.
        pass

    # Cria uma instância da classe “AppCadastro”, que representa a janela principal do programa.
    # Ao ser criada, essa instância monta toda a interface gráfica (campos, botões, tabelas, etc.)
    #                e conecta automaticamente ao banco já preparado.
    app = AppCadastro()

    # Executa o loop principal da interface Tkinter.
    # O método mainloop() mantém a janela aberta e escuta as ações do
    #                 usuário (como cliques e digitação).
    # Ele é essencial para que o programa gráfico permaneça ativo até que o usuário o feche.
    app.mainloop()