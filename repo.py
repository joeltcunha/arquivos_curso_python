# Importa dois tipos de dados da biblioteca 'typing': List e Tuple.
# Esses tipos são usados apenas para **anotação de tipo** (type hinting), ou seja,
#                para indicar ao leitor (ou ao editor de código) o tipo esperado
#                das variáveis e retornos das funções.
# - List → representa uma lista (por exemplo, [1, 2, 3])
# - Tuple → representa uma tupla (por exemplo, (1, "Maria", "maria@email.com"))
# Essas anotações não alteram o funcionamento do programa, mas tornam o
#                código mais legível e profissional.
from typing import List, Tuple

# Importa da configuração (arquivo config.py) o nome da tabela do banco de dados.
# A variável NOME_TABELA contém o nome da tabela usada no MySQL — neste caso, “pessoas”.
# Isso permite usar o mesmo nome em diferentes partes do código sem
#                precisar digitá-lo manualmente, facilitando manutenção e
#                evitando erros de digitação.
from config import NOME_TABELA


# Define a classe PessoaRepo (abreviação de "Pessoa Repository").
# Essa classe é responsável por **acessar diretamente o banco de dados**,
#                executando comandos SQL para cadastrar, listar, atualizar e excluir registros.
# Em sistemas profissionais, essa camada é conhecida como **DAO** (Data Access Object)
#                ou **repositório de dados**, pois ela isola a lógica de
#                acesso ao banco da lógica da interface.
class PessoaRepo:

    """Camada de acesso a dados (CRUD na tabela pessoas)."""
    # A docstring acima descreve resumidamente a função da classe:
    #                ela implementa as operações CRUD (Create, Read, Update, Delete)
    #                sobre a tabela “pessoas” do banco de dados MySQL.

    # Define o método construtor (__init__), que é executado automaticamente
    #               toda vez que uma nova instância da classe PessoaRepo é criada.
    # Ele inicializa o repositório e prepara a conexão com o banco de dados.
    def __init__(self, conn):

        # Armazena dentro da classe a conexão com o banco de dados (objeto 'conn'),
        #                recebida como parâmetro no momento em que a classe é instanciada.
        # Isso permite que todos os métodos da classe utilizem a mesma conexão,
        #                evitando a necessidade de abrir e fechar conexões repetidamente.
        self.conn = conn

        # Cria um cursor a partir da conexão.
        # O cursor é o objeto que permite executar comandos SQL (como SELECT, INSERT, UPDATE, DELETE).
        # Ele funciona como um “canal de comunicação” entre o Python e o MySQL.
        # Cada cursor é ligado à conexão que o criou (self.conn).
        self.cur = conn.cursor()

    # ----- CREATE
    # Define o método “inserir”, responsável por adicionar (criar) um
    #                 novo registro na tabela do banco de dados.
    # Esse método faz parte da operação "C" do CRUD — **Create** (criar).
    #
    # Parâmetros:
    #   nome (str) → nome da pessoa que será cadastrada.
    #   email (str) → endereço de e-mail da pessoa.
    #   telefone (str) → número de telefone.
    #
    # O "-> None" indica que a função **não retorna nenhum valor**,
    #                apenas executa uma ação (insere no banco).
    def inserir(self, nome: str, email: str, telefone: str) -> None:

        # Executa um comando SQL do tipo INSERT, que serve para inserir novos dados em uma tabela.
        # O método “execute()” envia esse comando para o MySQL através do cursor (self.cur).
        #
        # A string passada é uma *f-string* (por causa do prefixo f" "),
        #                permitindo inserir variáveis Python diretamente dentro da
        #                consulta — neste caso, o nome da tabela {NOME_TABELA}.
        #
        # A sintaxe geral do comando SQL é:
        #   INSERT INTO nome_da_tabela (coluna1, coluna2, coluna3)
        #   VALUES (valor1, valor2, valor3);
        #
        # No caso deste sistema:
        #   - {NOME_TABELA} → substituído por “pessoas”.
        #   - As colunas “nome”, “email” e “telefone” correspondem às colunas definidas no banco.
        #
        # ⚠ Importante: Os valores não são colocados diretamente dentro da string SQL.
        #   Em vez disso, usamos marcadores de posição **%s** — um para cada valor.
        #   Isso protege o sistema contra ataques de *SQL Injection*,
        #                  um tipo de ataque onde alguém tenta inserir comandos maliciosos no banco.
        #
        # Exemplo final que o MySQL receberá internamente:
        #   INSERT INTO pessoas (nome, email, telefone) VALUES ('Maria', 'maria@email.com', '9999-0000');
        self.cur.execute(
            f"INSERT INTO {NOME_TABELA} (nome, email, telefone) VALUES (%s, %s, %s)",

            # Aqui é passada uma tupla contendo os três valores que substituirão os %s acima:
            #                (nome, email, telefone)
            #
            # Por exemplo, se nome="Maria", email="maria@email.com", telefone="9999-0000",
            #                o PyMySQL automaticamente substituirá e escapará os valores corretamente
            #                antes de enviar o comando SQL ao servidor MySQL.
            (nome, email, telefone),

        )

    # ----- READ (listagens/pesquisa)
    # Define o método “listar_todos”, responsável por buscar **todos os registros**
    #                que estão armazenados na tabela do banco de dados.
    # Essa operação faz parte do "R" do CRUD — **Read** (ler ou consultar dados).
    #
    # O "-> List[Tuple]" indica que o método retorna uma **lista** contendo **tuplas**:
    #                cada tupla representa uma linha (registro) do banco de dados.
    # Exemplo de retorno:
    # [
    #   (1, "Maria", "maria@email.com", "9999-0000", "2020-11-04 18:32:00"),
    #   (2, "João", "joao@email.com", "9888-1234", "2030-11-04 18:45:21")
    # ]
    def listar_todos(self) -> List[Tuple]:

        # Executa o comando SQL do tipo SELECT, que serve para **consultar dados** no banco de dados.
        # O método “execute()” envia o comando ao MySQL e guarda o resultado internamente no cursor.
        #
        # Vamos entender o comando:
        #
        # ▶ SELECT id, nome, email, telefone, criado_em
        #   - “SELECT” é o comando que busca dados no banco.
        #   - Depois dele, listamos as colunas que queremos obter:
        #     “id”, “nome”, “email”, “telefone” e “criado_em”.
        #   - Isso significa que para cada registro encontrado, o banco devolverá esses cinco campos.
        #
        # ▶ FROM {NOME_TABELA}
        #   - Indica de qual tabela os dados serão extraídos.
        #   - A variável {NOME_TABELA} será substituída pelo nome da tabela (ex.: “pessoas”),
        #                    vindo do arquivo config.py.
        #
        # ▶ ORDER BY id DESC
        #   - “ORDER BY” ordena os resultados.
        #   - “id” é a coluna usada para a ordenação.
        #   - “DESC” significa *decrescente*, ou seja, o registro mais recente aparece primeiro.
        #
        # Exemplo de consulta final:
        #   SELECT id, nome, email, telefone, criado_em FROM pessoas ORDER BY id DESC;
        self.cur.execute(
            f"SELECT id, nome, email, telefone, criado_em "
            f"FROM {NOME_TABELA} ORDER BY id DESC"
        )

        # O método “fetchall()” recupera **todas as linhas** retornadas pela consulta SQL executada.
        # Ele retorna uma lista de tuplas, onde cada tupla representa um registro do banco.
        #
        # Exemplo de resultado:
        # [
        #   (1, "Maria", "maria@email.com", "9999-0000", "2030-11-04 18:32:00"),
        #   (2, "João", "joao@email.com", "9888-1234", "2030-11-04 18:45:21")
        # ]
        #
        # Esse retorno é enviado de volta ao código que chamou o método,
        #                para ser exibido na tabela (Treeview) da interface gráfica.
        return self.cur.fetchall()

    # Define o método “pesquisar”, responsável por **filtrar registros** no banco de dados
    #                com base em um texto digitado pelo usuário (chamado “termo”).
    #
    # Esse método faz parte do "R" do CRUD (Read) e retorna uma **lista de tuplas**
    #                contendo apenas os registros que correspondem parcialmente ao termo buscado.
    #
    # Exemplo: se o usuário digitar “mar”, o método pode retornar registros como:
    #   “Maria”, “Marcos”, “maria@email.com”, “(11) 9 7777-4321” etc.
    #
    # O tipo de retorno “-> List[Tuple]” indica que o resultado será uma lista de tuplas,
    #                em que cada tupla representa uma linha da tabela.
    def pesquisar(self, termo: str) -> List[Tuple]:

        # Cria uma variável “like” que adiciona o caractere curinga "%" antes e
        #                depois do termo digitado.
        #
        # Em SQL, o operador LIKE é usado para realizar comparações parciais:
        #   - O símbolo % significa “qualquer sequência de caracteres”.
        #   - Por exemplo:
        #       termo = "mar"
        #       like = "%mar%"
        #     Isso fará o banco procurar qualquer valor que contenha “mar”,
        #                    como “Maria”, “Marcos”, “maria@email.com” etc.
        like = f"%{termo}%"

        # Executa o comando SQL SELECT, que consulta registros na tabela de pessoas,
        #                filtrando por nome, email ou telefone que contenham o texto buscado.
        #
        # Vamos entender o comando completo:
        #
        # ▶ SELECT id, nome, email, telefone, criado_em
        #   - Define as colunas que queremos recuperar da tabela.
        #
        # ▶ FROM {NOME_TABELA}
        #   - Indica de qual tabela os dados serão buscados.
        #
        # ▶ WHERE nome LIKE %s OR email LIKE %s OR telefone LIKE %s
        #   - O operador WHERE filtra as linhas que correspondem à condição.
        #   - O operador LIKE permite fazer buscas **parciais** (com os %).
        #   - O “OR” indica que basta uma das três condições ser verdadeira
        #                    (nome, email ou telefone conter o termo) para o registro ser retornado.
        #   - Os marcadores %s são substituídos com segurança pelos valores da tupla
        #                    (like, like, like), evitando vulnerabilidades como SQL Injection.
        #
        # ▶ ORDER BY id DESC
        #   - Ordena os resultados do mais recente para o mais antigo.
        #
        # Exemplo de consulta que será executada internamente:
        #   SELECT id, nome, email, telefone, criado_em
        #   FROM pessoas
        #   WHERE nome LIKE '%mar%' OR email LIKE '%mar%' OR telefone LIKE '%mar%'
        #   ORDER BY id DESC;
        self.cur.execute(
            f"""
            SELECT id, nome, email, telefone, criado_em
            FROM {NOME_TABELA}
            WHERE nome LIKE %s OR email LIKE %s OR telefone LIKE %s
            ORDER BY id DESC
            """,

            # Esta tupla fornece os valores que substituirão os três marcadores %s do comando SQL acima.
            # Cada campo (nome, email e telefone) receberá o mesmo padrão de busca “%termo%”.
            # O PyMySQL se encarrega de escapar corretamente os valores,
            #               impedindo que caracteres especiais quebrem o comando SQL.
            (like, like, like),

        )

        # O método fetchall() recupera todas as linhas encontradas pela consulta.
        # O retorno é uma lista de tuplas, sendo cada tupla um registro completo do banco.
        #
        # Exemplo:
        # [
        #   (1, "Maria", "maria@email.com", "9999-0000", "2030-11-04 18:32:00"),
        #   (3, "Marcos", "marcos@email.com", "9888-1234", "2030-11-04 19:15:44")
        # ]
        return self.cur.fetchall()

    # ----- UPDATE / DELETE
    # Define o método “atualizar”, responsável por modificar (editar) os dados de
    #                 uma pessoa já existente na tabela.
    # Essa operação faz parte do "U" do CRUD — **Update** (atualizar registros).
    #
    # Parâmetros:
    #   id_ (int) → identificador único da pessoa (chave primária).
    #   nome (str) → novo nome da pessoa.
    #   email (str) → novo e-mail da pessoa.
    #   telefone (str) → novo número de telefone.
    #
    # O "-> None" indica que este método não retorna valor — apenas executa a ação no banco.
    def atualizar(self, id_: int, nome: str, email: str, telefone: str) -> None:

        # Executa um comando SQL do tipo UPDATE, que altera dados existentes em uma tabela.
        #
        # A estrutura geral do comando SQL é:
        #   UPDATE nome_tabela
        #   SET coluna1 = novo_valor1, coluna2 = novo_valor2, ...
        #   WHERE condição;
        #
        # Nesse caso:
        # - UPDATE {NOME_TABELA} → atualiza a tabela definida em config.py (ex.: “pessoas”).
        # - SET nome=%s, email=%s, telefone=%s → define os novos valores para as colunas.
        # - WHERE id=%s → indica qual registro deve ser atualizado, baseado no ID.
        #
        # ⚠ Importante:
        # O uso de %s (placeholders) é fundamental para evitar *SQL Injection*.
        # O PyMySQL substitui automaticamente cada marcador pelo valor correspondente
        #                de forma segura, escapando caracteres especiais.
        #
        # Exemplo do comando final que será enviado ao MySQL:
        #   UPDATE pessoas SET nome='Ana Souza', email='ana@email.com', telefone='98888-1111' WHERE id=3;
        self.cur.execute(
            f"UPDATE {NOME_TABELA} SET nome=%s, email=%s, telefone=%s WHERE id=%s",

            # Aqui são passados os valores que substituirão cada %s, na mesma ordem:
            #   1º → nome
            #   2º → email
            #   3º → telefone
            #   4º → id_ (o identificador do registro que será alterado)
            # Essa tupla é enviada junto com o comando SQL para garantir que os dados
            #                sejam formatados e protegidos adequadamente.
            (nome, email, telefone, id_),

        )

    # Define o método “excluir”, responsável por **remover** um registro
    #                 específico da tabela no banco de dados.
    # Essa operação faz parte do "D" do CRUD — **Delete** (excluir).
    #
    # Parâmetros:
    #   id_ (int) → representa o identificador único (chave primária) do registro que será deletado.
    #
    # O método não retorna nada (-> None), pois sua função é apenas executar o comando de exclusão.
    def excluir(self, id_: int) -> None:

        # Executa um comando SQL do tipo DELETE, que serve para apagar
        #                uma ou mais linhas de uma tabela.
        #
        # A estrutura geral do comando SQL é:
        #   DELETE FROM nome_tabela WHERE condição;
        #
        # - DELETE FROM {NOME_TABELA} → indica qual tabela será afetada (ex.: “pessoas”).
        # - WHERE id=%s → especifica **qual linha** deve ser removida, com base no ID único da pessoa.
        #   Essa condição é **essencial**, pois sem o “WHERE” o MySQL excluiria
        #                  TODAS as linhas da tabela.
        #
        # ⚠ Por isso, sempre use “WHERE” ao excluir dados, especialmente com chaves primárias.
        #
        # O marcador %s será substituído com segurança pelo valor do parâmetro id_.
        # Exemplo final:
        #   DELETE FROM pessoas WHERE id=5;
        #
        # Isso apagará apenas o registro cujo ID for 5.
        self.cur.execute(

            # O comando SQL DELETE é enviado ao MySQL através do cursor (self.cur).
            f"DELETE FROM {NOME_TABELA} WHERE id=%s",

            # Passa uma tupla contendo o valor do ID que substituirá o marcador %s.
            # Note que mesmo com apenas um valor, ainda usamos uma vírgula no final (id_,)
            #                para garantir que o Python entenda corretamente que se trata de uma tupla.
            (id_,)

        )