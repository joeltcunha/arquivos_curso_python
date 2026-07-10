# Importa o módulo principal 'pymysql', que é a biblioteca usada para
#                conectar o Python ao banco de dados MySQL.
# Essa biblioteca permite executar comandos SQL (como SELECT, INSERT,
#                UPDATE, DELETE) diretamente a partir do Python.
import pymysql

# Importa o submódulo 'err' de dentro do pacote pymysql e o renomeia para 'pymysql_err'.
# Isso é feito para facilitar a leitura do código quando formos tratar erros específicos do MySQL.
# Exemplo: em vez de escrever "pymysql.err.OperationalError",
#                podemos usar "pymysql_err.OperationalError".
from pymysql import err as pymysql_err

# Importa do arquivo 'config.py' todas as variáveis de configuração
#                necessárias para a conexão com o banco de dados.
# São elas:
#   HOST → Endereço do servidor MySQL (ex.: 'localhost')
#   PORTA → Número da porta usada para comunicação (ex.: 3306)
#   USUARIO → Nome do usuário do banco (ex.: 'root')
#   SENHA → Senha do usuário para autenticação
#   NOME_BANCO → Nome do banco de dados (schema) que será usado ou criado
#   NOME_TABELA → Nome da tabela que será manipulada dentro desse banco
# Essas variáveis são essenciais para criar a conexão e executar comandos SQL corretamente.
from config import HOST, PORTA, USUARIO, SENHA, NOME_BANCO, NOME_TABELA

# Define a função 'conectar', responsável por abrir uma conexão com o banco de dados MySQL.
# O parâmetro 'usar_banco' (por padrão = True) indica se a conexão já
#                deve ser feita diretamente com o banco de dados
#                definido em NOME_BANCO. Caso ele ainda não exista, o
#                código cria o banco e depois reconecta.
def conectar(usar_banco=True):

    """
    Abre conexão com o MySQL usando PyMySQL.
    Se o banco ainda não existir, cria e reconecta.
    """

    # Inicia um bloco try/except para capturar e tratar possíveis erros de conexão,
    #                como falha na autenticação, servidor MySQL desligado ou banco inexistente.
    try:

        # Verifica se a função deve se conectar diretamente ao banco de dados já existente.
        # Caso 'usar_banco' seja True, tenta conectar usando o nome do banco definido em NOME_BANCO.
        if usar_banco:

            # Retorna uma conexão ativa com o MySQL, configurada com as
            #                informações definidas em config.py.
            # Essa conexão é usada para executar comandos SQL em outras partes do sistema.
            return pymysql.connect(

                # host → endereço do servidor MySQL ('localhost' para conexões locais)
                host=HOST,

                # port → porta de comunicação usada pelo MySQL (padrão: 3306)
                port=PORTA,

                # user → nome do usuário autorizado a acessar o banco de dados
                user=USUARIO,

                # password → senha do usuário definida no servidor MySQL
                password=SENHA,

                # database → nome do banco de dados que será acessado (ex.: 'cadastro_db')
                database=NOME_BANCO,

                # charset → define a codificação de caracteres usada para armazenar textos.
                # utf8mb4 é uma versão aprimorada do UTF-8, que suporta
                #                emojis e todos os caracteres especiais.
                charset="utf8mb4",

                # autocommit=True → faz com que cada comando SQL seja confirmado automaticamente
                #                sem precisar chamar conn.commit() manualmente após cada operação.
                autocommit=True,

                # cursorclass → define o tipo de cursor usado para manipular os
                #                resultados das consultas SQL.
                # Aqui, pymysql.cursors.Cursor é usado, o que retorna os
                #                dados como tuplas (linhas de valores).
                cursorclass=pymysql.cursors.Cursor)


        # Caso o parâmetro 'usar_banco' seja False, o sistema se conecta apenas ao servidor MySQL,
        #                sem especificar o nome do banco de dados (NOME_BANCO).
        # Isso é útil em situações em que o banco de dados ainda não existe e precisa ser criado.
        else:

            # Retorna uma conexão ativa com o servidor MySQL (sem selecionar um banco específico).
            # Essa conexão é usada, por exemplo, para executar o comando CREATE DATABASE.
            return pymysql.connect(

                # host → endereço do servidor MySQL, “localhost”,
                #                o que significa que o servidor está rodando na
                #                mesma máquina do programa.
                host=HOST,

                # port → número da porta usada pelo servidor MySQL (por padrão, 3306).
                port=PORTA,

                # user → nome do usuário autorizado no banco de dados (“root” em ambientes locais).
                user=USUARIO,

                # password → senha correspondente ao usuário informado.
                # Essa autenticação é obrigatória para permitir o acesso ao servidor MySQL.
                password=SENHA,

                # charset → define o conjunto de caracteres usado pela conexão.
                # “utf8mb4” é uma codificação moderna e completa, que
                #                suporta todos os caracteres Unicode.
                charset="utf8mb4",

                # autocommit=True → faz com que cada operação SQL seja automaticamente confirmada (commit),
                #                sem precisar chamar manualmente o método .commit() após
                #                cada inserção ou atualização.
                autocommit=True,

                # cursorclass → define o tipo de cursor usado para ler os resultados das consultas SQL.
                # O cursor “Cursor” padrão retorna as linhas dos resultados como tuplas simples.
                cursorclass=pymysql.cursors.Cursor

            )


    # Captura um erro específico do tipo "OperationalError", que ocorre quando há falha operacional
    #                durante a tentativa de conexão ao banco de dados MySQL.
    # Exemplos de falhas: banco inexistente, servidor desligado, senha incorreta, etc.
    except pymysql_err.OperationalError as e:

        # O código de erro 1049 no MySQL significa "Unknown database" (banco de dados desconhecido),
        #                ou seja, o nome do banco definido em NOME_BANCO ainda não existe.
        # Nesse caso, o sistema cria o banco automaticamente e tenta se reconectar.
        if getattr(e, "args", None) and e.args and e.args[0] == 1049:

            # Cria uma nova conexão com o servidor MySQL,
            #                mas desta vez **sem** especificar o nome do banco (pois
            #                ele ainda não existe).
            # Essa conexão temporária é usada apenas para criar o banco e a tabela.
            conn = pymysql.connect(

                # host → endereço do servidor MySQL (“localhost”)
                host=HOST,

                # port → número da porta do servidor MySQL (padrão 3306)
                port=PORTA,

                # user → nome de usuário com permissão para criar bancos e tabelas (ex.: “root”)
                user=USUARIO,

                # password → senha do usuário
                password=SENHA,

                # charset → define a codificação de caracteres usada pela conexão.
                # “utf8mb4” garante compatibilidade com caracteres especiais e emojis.
                charset="utf8mb4",

                # autocommit=True → ativa a confirmação automática das operações SQL,
                #                evitando a necessidade de chamar .commit() manualmente.
                autocommit=True,

                # cursorclass → tipo de cursor usado para executar os comandos SQL.
                cursorclass=pymysql.cursors.Cursor

            )

            # Chama a função 'criar_banco_e_tabela', responsável por:
            #   1. Criar o banco de dados (caso não exista)
            #   2. Criar a tabela “pessoas” dentro do banco.
            # Essa função está definida mais abaixo neste mesmo arquivo.
            criar_banco_e_tabela(conn)

            # Fecha a conexão temporária criada apenas para montar o banco e tabela.
            # Isso é importante para liberar recursos antes de abrir uma nova conexão definitiva.
            conn.close()

            # Após criar o banco de dados e a tabela, é necessário estabelecer uma nova conexão —
            #                desta vez incluindo o nome do banco (NOME_BANCO), pois agora ele já existe.
            # Esse retorno final entrega uma conexão pronta para uso em todo o sistema.
            return pymysql.connect(

                # host → endereço do servidor MySQL, “localhost”
                host=HOST,

                # port → número da porta usada pelo MySQL (3306)
                port=PORTA,

                # user → nome do usuário do MySQL (ex.: “root”)
                user=USUARIO,

                # password → senha associada ao usuário
                password=SENHA,

                # database → nome do banco de dados recém-criado (ex.: “cadastro_db”)
                # Aqui ele é incluído para que todas as operações SQL sejam
                #                executadas dentro do banco certo.
                database=NOME_BANCO,

                # charset → define o padrão de codificação de texto (utf8mb4 é recomendado por
                #                suportar todos os caracteres)
                charset="utf8mb4",

                # autocommit=True → garante que cada comando SQL seja confirmado automaticamente,
                #                evitando a necessidade de chamar manualmente o
                #                método commit() após cada operação.
                autocommit=True,

                # cursorclass → define o tipo de cursor usado nas consultas.
                # pymysql.cursors.Cursor retorna os resultados das consultas como tuplas simples.
                cursorclass=pymysql.cursors.Cursor

            )

        # Se o erro capturado não for do tipo “banco inexistente” (1049),
        #                o Python reenvia (raise) a exceção original, permitindo
        #                que ela seja tratada em outro nível do código.
        # Isso evita suprimir erros graves (como senha incorreta, servidor desligado etc.).
        raise


# Define a função 'criar_banco_e_tabela', que recebe como parâmetro uma conexão ativa com o MySQL (conn).
# Essa função tem a responsabilidade de garantir que o banco de dados e a
#                tabela necessários existam antes de o sistema começar a utilizá-los.
# Se ainda não existirem, eles serão criados automaticamente.
def criar_banco_e_tabela(conn):

    """Cria o database e a tabela (se não existirem)."""

    # O comando "with conn.cursor()" cria um cursor de forma temporária.
    # O cursor é o objeto que permite executar comandos SQL (como CREATE, SELECT, INSERT, etc.)
    #                dentro da conexão aberta (conn).
    # A palavra-chave "as cur" dá o nome “cur” ao cursor, facilitando seu uso dentro do bloco.
    # Ao final do bloco "with", o cursor é automaticamente fechado, mesmo se ocorrer algum erro.
    with conn.cursor() as cur:

        # Este trecho de código é o coração da criação do banco de dados.
        # Ele executa um comando SQL (Structured Query Language), que é a linguagem usada
        #               para conversar com o MySQL e manipular bancos de dados (criar,
        #               apagar, modificar, etc).
        #
        # 👉 Vamos entender cada parte detalhadamente:

        # ▶ f"CREATE DATABASE IF NOT EXISTS {NOME_BANCO}"
        # - "CREATE DATABASE" é um comando SQL que cria um novo banco de dados.
        # - "IF NOT EXISTS" é uma cláusula de segurança que impede o MySQL de gerar um erro
        #                  caso o banco já exista — ele apenas ignora o
        #                  comando se o banco já estiver criado.
        # - {NOME_BANCO} é uma variável Python que é substituída dinamicamente pelo nome do banco,
        #                  por exemplo, "cadastro_db". O "f" antes das aspas indica que é
        #                  uma *f-string*, ou seja, permite inserir variáveis dentro da string.

        # ▶ "DEFAULT CHARACTER SET utf8mb4"
        # - "CHARACTER SET" define qual conjunto de caracteres será usado para armazenar textos no banco.
        # - "utf8mb4" é uma codificação moderna e completa que suporta todos os caracteres Unicode,
        #                  incluindo acentos (como á, ê, ç) e até emojis (😀🔥❤️).
        #   Usar utf8mb4 evita problemas de caracteres “bugados” quando se insere textos com acentos.

        # ▶ "COLLATE utf8mb4_general_ci"
        # - "COLLATE" define as regras de *comparação* e *ordenação* de textos.
        # - "utf8mb4_general_ci" significa:
        #     - “utf8mb4” → mesmo conjunto de caracteres do CHARSET.
        #     - “general” → tipo de regra geral de ordenação.
        #     - “ci” → *case insensitive*, ou seja, ignora diferença entre maiúsculas e minúsculas.
        #       Exemplo: “Maria” e “maria” são considerados iguais.
        #
        # Em resumo, esse comando cria o banco de dados com suporte a todos os
        #                caracteres e comparações amigáveis, o que é essencial para
        #                aplicativos em português.

        # O método .execute() envia o comando SQL ao servidor MySQL através do cursor “cur”.
        # O MySQL interpreta a string e executa o comando, criando fisicamente o banco no servidor.
        cur.execute(
            f"CREATE DATABASE IF NOT EXISTS {NOME_BANCO} "
            "DEFAULT CHARACTER SET utf8mb4 COLLATE utf8mb4_general_ci;"
        )

        # Em seguida, executa o comando SQL "USE", que seleciona o banco de dados recém-criado
        #                (ou existente) para que os próximos comandos SQL sejam executados dentro dele.
        # Esse comando é obrigatório antes de criar tabelas, pois o MySQL precisa saber
        #                em qual banco de dados a tabela será criada.
        cur.execute(f"USE {NOME_BANCO};")

        # Este comando cria a **tabela** principal dentro do banco de dados, se ela ainda não existir.
        # Vamos analisar o comando SQL linha por linha, pois ele é fundamental para compreender
        #                como os dados são realmente armazenados dentro do MySQL.

        # ▶ f""" ... """
        # - O prefixo “f” indica que é uma *f-string*, permitindo inserir variáveis
        #                Python diretamente no texto SQL.
        # - As três aspas """ """ permitem escrever o comando SQL em múltiplas
        #                linhas, deixando o código mais legível.
        # - {NOME_TABELA} será substituído pelo nome definido na variável (ex.: “pessoas”).
        #
        # Assim, o Python enviará algo como:
        #   CREATE TABLE IF NOT EXISTS pessoas (...)
        cur.execute(
            f"""
            CREATE TABLE IF NOT EXISTS {NOME_TABELA} (
                id INT AUTO_INCREMENT PRIMARY KEY,
                nome VARCHAR(120) NOT NULL,
                email VARCHAR(150) NOT NULL,
                telefone VARCHAR(30) NOT NULL,
                criado_em TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;
            """
        )

        # Agora, vamos entender detalhadamente cada parte do comando SQL acima:

        # ▶ CREATE TABLE IF NOT EXISTS {NOME_TABELA}
        # - “CREATE TABLE” é o comando SQL usado para criar uma tabela dentro do banco de dados.
        # - “IF NOT EXISTS” evita erro caso a tabela já exista — o MySQL apenas ignora o comando.
        # - {NOME_TABELA} é o nome da tabela, por exemplo “pessoas”.
        #   Essa tabela armazenará todos os cadastros feitos na aplicação.

        # ▶ id INT AUTO_INCREMENT PRIMARY KEY
        # - “id” é o nome da primeira coluna. Ela servirá como **identificador único** de cada registro.
        # - “INT” significa que o tipo de dado é um número inteiro.
        # - “AUTO_INCREMENT” faz com que o MySQL incremente automaticamente o valor do ID
        #                  toda vez que uma nova pessoa for cadastrada (1, 2, 3, 4...).
        # - “PRIMARY KEY” define essa coluna como chave primária, ou seja,
        #                  ela identifica unicamente cada linha da tabela e não pode se repetir.

        # ▶ nome VARCHAR(120) NOT NULL
        # - “nome” é o campo onde será armazenado o nome completo da pessoa.
        # - “VARCHAR(120)” define um texto variável (string) de até **120 caracteres**.
        #   “VARCHAR” significa “VARiable CHARacter”, ou seja, o tamanho do texto pode variar.
        # - “NOT NULL” significa que o campo é **obrigatório** — não pode ficar vazio.

        # ▶ email VARCHAR(150) NOT NULL
        # - “email” armazena o endereço de e-mail da pessoa.
        # - O tipo “VARCHAR(150)” permite até 150 caracteres.
        # - Também é obrigatório (“NOT NULL”), garantindo que nenhum cadastro fique sem e-mail.

        # ▶ telefone VARCHAR(30) NOT NULL
        # - “telefone” guarda o número de telefone da pessoa.
        # - “VARCHAR(30)” é suficiente para armazenar formatos variados (com DDD, traços, espaços etc.).
        # - Também é obrigatório (“NOT NULL”).

        # ▶ criado_em TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        # - “criado_em” é um campo do tipo “TIMESTAMP”, ou seja, guarda **data e hora**.
        # - “DEFAULT CURRENT_TIMESTAMP” faz com que o MySQL insira automaticamente a data e hora
        #                  no momento em que o registro é criado.
        #   Exemplo: “2030-11-04 18:32:00”.
        #   Assim, o sistema sempre saberá quando cada cadastro foi feito.

        # ▶ ENGINE=InnoDB
        # - “ENGINE” indica o tipo de mecanismo de armazenamento usado pelo MySQL.
        # - “InnoDB” é o mais moderno e seguro, suportando transações e chaves estrangeiras
        #                  (úteis para relacionar tabelas no futuro, caso o sistema cresça).

        # ▶ DEFAULT CHARSET=utf8mb4
        # - Define que o padrão de codificação de texto da tabela será “utf8mb4”,
        #                  o mesmo usado na criação do banco, garantindo
        #                  compatibilidade com acentos e emojis.

        # Em resumo:
        # Esse comando cria uma tabela totalmente estruturada e otimizada para armazenar
        # cadastros de pessoas com segurança, padronização e compatibilidade internacional.