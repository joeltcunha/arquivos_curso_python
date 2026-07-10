# ===============================================================
# Importações principais — módulo de interface gráfica, manipulação
#           de arquivos Excel e comunicação com o banco de dados MySQL.
# Cada importação abaixo tem uma função específica dentro da aplicação.
# ===============================================================

# Importa o módulo principal do Tkinter com o apelido 'tk'.
# Este módulo contém a base para criar janelas, labels, botões, etc.
import tkinter as tk

# Importa de dentro do módulo tkinter componentes adicionais:
# - ttk: fornece widgets "temáticos" mais modernos (Label, Button, Entry, Treeview).
# - messagebox: exibe caixas de mensagem (informações, alertas, erros, confirmações).
# - filedialog: cria caixas de diálogo para escolher e salvar
#           arquivos (ex.: salvar Excel).
from tkinter import ttk, messagebox, filedialog

# Importa a classe Workbook da biblioteca openpyxl.
# A classe Workbook é usada para criar, manipular e salvar arquivos Excel (.xlsx).
from openpyxl import Workbook

# Importa a função get_column_letter da openpyxl.utils.
# Essa função converte números de coluna (1, 2, 3...) em letras (A, B, C...),
#           usada ao ajustar larguras de colunas em planilhas Excel.
from openpyxl.utils import get_column_letter

# ===============================================================
# Importações internas (do próprio projeto).
# Estes módulos foram criados pelo desenvolvedor e fazem parte da arquitetura.
# ===============================================================

# Importa a função 'conectar' do módulo db.py.
# Esta função é responsável por abrir a conexão com o banco de dados MySQL
#           (e criar o banco/tabela automaticamente se ainda não existirem).
from db import conectar

# Importa a classe 'PessoaRepo' do módulo repo.py.
# Essa classe é a camada de acesso a dados (CRUD — Create, Read, Update, Delete)
#           que contém todos os métodos de interação com a tabela de pessoas no banco.
from repo import PessoaRepo

# Importa funções de estilo do módulo style.py.
# - aplicar_estilo: aplica um tema visual agradável aos componentes ttk.
# - aplicar_zebra_treeview: define cores alternadas nas linhas da tabela (efeito "zebra").
# - preencher_treeview_com_zebra: preenche a tabela com os
#           registros aplicando o padrão de cor.
from style import aplicar_estilo, aplicar_zebra_treeview, preencher_treeview_com_zebra


# =========================
# Validação simples
# =========================

# Define uma função chamada validar.
# Uma função é um bloco de código que só é executado quando é
#           chamado em outro lugar do programa.
# Essa função serve para checar se os dados digitados pelo usuário (nome, e-mail e telefone)
#           estão corretos antes de serem salvos no banco de dados.
# A função recebe três informações (chamadas de parâmetros): nome, email e telefone.
# Depois de cada nome de parâmetro, temos ": str", que é chamado de *anotação de tipo*.
# Isso indica que esperamos que esses valores sejam do tipo string (texto).
def validar(nome: str, email: str, telefone: str):

    # Aqui começa a primeira verificação: o nome da pessoa.
    # A expressão "not nome" verifica se o campo nome está vazio.
    # Em Python, uma string vazia ("") é considerada "falsa", então "not nome"
    #           será verdadeiro se o nome estiver em branco.
    # A função strip() remove os espaços em branco do início e do fim da string.
    # Isso evita que alguém digite apenas espaços e o programa ache que digitou algo válido.
    # A função len() retorna o tamanho da string, ou seja, quantos caracteres ela possui.
    # Portanto, len(nome.strip()) < 3 verifica se, após remover os
    #           espaços, o nome tem menos de 3 letras.
    if not nome or len(nome.strip()) < 3:

        # Se a condição acima for verdadeira (ou seja, o nome está vazio ou é muito curto),
        #           o programa entra neste bloco de código.
        # O comando return faz a função parar de executar e devolver um resultado.
        # Aqui ela devolve dois valores:
        # 1º - False → indicando que a validação falhou (os dados são inválidos)
        # 2º - Uma mensagem de texto explicando o erro ao usuário
        return False, "Informe um nome com pelo menos 3 caracteres."

    # Agora é feita a verificação do e-mail.
    # A primeira parte "not email" testa se o campo de e-mail está vazio.
    # A segunda parte '"@" not in email' verifica se o símbolo "@" está presente no texto.
    # Em um endereço de e-mail válido, o caractere "@" é obrigatório (exemplo: exemplo@gmail.com).
    # Se faltar esse símbolo, o e-mail não é considerado válido.
    if not email or "@" not in email:

        # Caso o e-mail seja inválido (vazio ou sem @),
        #           a função também para imediatamente e retorna False e
        #           uma mensagem de aviso.
        return False, "Informe um e-mail válido (ex.: nome@dominio.com)."

    # Agora o programa verifica o telefone.
    # Assim como antes, "not telefone" verifica se o campo está vazio.
    # A função strip() novamente remove espaços em branco no começo e no fim,
    #               e len(telefone.strip()) < 6 garante que o número tenha
    #               pelo menos 6 caracteres.
    # Isso serve para evitar telefones curtos demais, como "123" ou "9999",
    #               que não seriam números de telefone reais.
    if not telefone or len(telefone.strip()) < 6:

        # Se o telefone for inválido (muito curto ou vazio),
        #           a função retorna False e uma mensagem explicando o erro.
        return False, "Informe um telefone válido."

    # Se nenhuma das condições anteriores foi acionada (ou seja, tudo está correto),
    #           o programa chega até esta linha final.
    # Aqui a função retorna dois valores:
    # 1º - True → indicando que todos os dados foram validados com sucesso.
    # 2º - Uma string vazia "" → significando que não há mensagem de erro.
    # Esse tipo de retorno (True/False + mensagem) é muito útil para outras partes do sistema,
    #           porque permite que o código principal saiba rapidamente
    #           se os dados são válidos e, se não forem, mostre a
    #           mensagem correta na tela.
    return True, ""


# =========================
# Janela principal
# =========================

# Define uma classe chamada AppCadastro.
# Em Python, uma classe é uma estrutura que serve como modelo
#           para criar objetos que agrupam dados (variáveis) e
#           comportamentos (funções).
# Aqui, a classe AppCadastro representa a janela principal do
#           de cadastro de pessoas.
# Ela herda (recebe características) da classe tk.Tk, que é a janela raiz do Tkinter.
class AppCadastro(tk.Tk):

    # Define o método construtor __init__, que é executado automaticamente
    #           sempre que criamos um novo objeto da classe AppCadastro.
    # Ele serve para inicializar a janela e todos os seus componentes.
    def __init__(self):

        # Usa super().__init__() para chamar o construtor da classe pai (tk.Tk),
        #           garantindo que a janela principal do Tkinter seja
        #           corretamente inicializada.
        super().__init__()

        # Define o título da janela que aparece na barra superior.
        # Este título ajuda o usuário a identificar o propósito da aplicação.
        self.title("Cadastro de Pessoas - Tkinter + MySQL")

        # Define o tamanho inicial da janela (largura x altura) em pixels.
        # Aqui, a janela começa com 980 pixels de largura e 580 de altura.
        self.geometry("980x580")     # tamanho inicial

        # Define o tamanho mínimo que a janela pode ter ao ser redimensionada.
        # Isso impede que o usuário a reduza tanto a ponto de
        #               prejudicar a visualização dos elementos.
        self.minsize(900, 520)

        # -------------------------------------------------------------
        # Aplicação do tema visual (estilo) à interface.
        # A função aplicar_estilo é importada do módulo style.py.
        # Ela define cores, fontes, espaçamentos e aparência geral dos botões,
        #           labels, caixas de texto e da tabela (Treeview), deixando o layout
        #           mais agradável e padronizado.
        # O argumento 'self' é passado porque a janela principal (AppCadastro)
        #           é o elemento raiz onde o estilo será aplicado.
        # -------------------------------------------------------------
        aplicar_estilo(self)

        # -------------------------------------------------------------
        # Criação da conexão com o banco de dados e do repositório
        # -------------------------------------------------------------

        # Chama a função conectar(), que foi importada do módulo db.py.
        # Essa função abre uma conexão com o banco de dados MySQL.
        # O parâmetro usar_banco=True indica que a conexão deve ser feita
        #           diretamente com o banco de dados já existente (e não
        #           apenas com o servidor).
        # Caso o banco ou a tabela ainda não existam, o código do módulo db.py
        #           é capaz de criá-los automaticamente antes de continuar.
        self.conn = conectar(usar_banco=True)

        # Cria um objeto da classe PessoaRepo, que foi importada do módulo repo.py.
        # Essa classe contém todos os métodos responsáveis pelas operações de
        #           banco de dados (CRUD): inserir, listar, pesquisar,
        #           atualizar e excluir.
        # Passamos a conexão (self.conn) como parâmetro para que o repositório
        #           possa executar os comandos SQL usando essa mesma conexão aberta.
        self.repo = PessoaRepo(self.conn)

        # -------------------------------------------------------------
        # Declaração das variáveis de controle dos campos da interface
        # -------------------------------------------------------------
        # No Tkinter, variáveis como StringVar() servem como "ponte" entre
        #           os widgets (campos de texto, labels etc.) e o código Python.
        # Elas permitem que, ao alterar o texto em um campo, o valor seja
        #           automaticamente atualizado na variável, e vice-versa.

        # Cria uma variável para armazenar o ID da pessoa.
        # O ID é um número gerado automaticamente pelo banco (chave primária),
        #           e não pode ser editado pelo usuário.
        self.var_id = tk.StringVar()

        # Cria uma variável para armazenar o nome da pessoa digitado no formulário.
        # Essa variável será usada para ler o nome e também para limpar o campo.
        self.var_nome = tk.StringVar()

        # Cria uma variável para armazenar o e-mail digitado pelo usuário.
        # O valor será lido quando o usuário clicar em “Cadastrar” ou “Atualizar”.
        self.var_email = tk.StringVar()

        # Cria uma variável para armazenar o telefone informado no cadastro.
        # Assim como as demais, ela é atualizada automaticamente junto
        #               com o campo de entrada.
        self.var_telefone = tk.StringVar()

        # Cria uma variável que guardará o texto digitado no campo de pesquisa.
        # Essa variável é usada para filtrar registros por nome, e-mail ou telefone.
        self.var_pesquisa = tk.StringVar()

        # Cria uma variável para exibir o status da aplicação (mensagens informativas).
        # O parâmetro value define o valor inicial como "Pronto.",
        #               que será exibido na barra de status da janela inferior.
        self.var_status = tk.StringVar(value="Pronto.")

        # -------------------------------------------------------------
        # Montagem da tela (Interface Gráfica)
        # -------------------------------------------------------------
        # Após definir todas as variáveis e configurar a conexão,
        #           é hora de montar os elementos visuais da janela.
        # O Tkinter permite que criemos funções separadas para organizar melhor
        #           o código da interface, mantendo cada parte da tela em
        #           seu próprio método.
        # As funções abaixo (com prefixo "_montar_") criam e posicionam os widgets.

        # Chama o método que cria o formulário de entrada de dados.
        # Este formulário contém os campos de texto (ID, Nome, E-mail, Telefone)
        #           e também o campo de pesquisa.
        # Dentro dessa função são criados os Labels (rótulos) e
        #           Entrys (caixas de texto) que o usuário utilizará
        #           para digitar as informações.
        self._montar_formulario()

        # Chama o método que cria a área de botões principais da aplicação.
        # Esses botões executam as ações: Cadastrar, Atualizar, Excluir,
        # Limpar Campos e Pesquisar.
        # Cada botão tem sua função associada a um método do programa.
        self._montar_botoes()

        # Chama o método responsável por criar e configurar a tabela (Treeview).
        # Essa tabela exibirá todos os registros cadastrados no banco de dados.
        # Ela é configurada com cabeçalhos, largura de colunas,
        #           rolagem vertical e horizontal, e evento de duplo clique
        #           para carregar o registro no formulário.
        self._montar_tabela()

        # Chama o método que cria a barra de status (statusbar).
        # Essa barra fica na parte inferior da janela e exibe mensagens
        #           como “Pronto.”, “Cadastro realizado com sucesso.”,
        #           “Erro ao atualizar.” etc.
        self._montar_statusbar()

        # -------------------------------------------------------------
        # Centralização da janela principal na tela
        # -------------------------------------------------------------
        # Após todos os elementos visuais serem criados e posicionados,
        #           é feita a centralização da janela para que ela apareça
        # bem no centro da tela do monitor do usuário.
        # Essa centralização leva em conta a largura e altura reais da janela
        #           (após o carregamento dos widgets) e ajusta automaticamente
        #           a posição X e Y no monitor.
        self._centralizar_janela()

        # -------------------------------------------------------------
        # Atalhos de teclado (atalhos úteis para o usuário)
        # -------------------------------------------------------------
        # O método .bind() serve para associar uma combinação de teclas
        #           a uma determinada ação (função) dentro da janela do Tkinter.
        # Assim o usuário pode realizar tarefas rapidamente sem clicar em botões.

        # Aqui estamos criando um atalho para limpar todos os campos do formulário.
        # "<Control-n>" representa a combinação de teclas Ctrl + N.
        # Quando o usuário pressiona essas teclas, é executada uma
        #           função anônima (lambda) que chama o método _limpar_campos().
        # Esse método apaga o conteúdo dos campos Nome, E-mail e Telefone,
        #           e também limpa qualquer seleção feita na tabela.
        self.bind("<Control-n>", lambda e: self._limpar_campos())  # limpar

        # Cria outro atalho de teclado, agora para atualizar a lista de registros.
        # "<F5>" representa a tecla F5 do teclado, muito usada para "atualizar" em sistemas.
        # Quando o usuário pressiona F5, a função listar_todos() é chamada,
        #           atualizando a tabela (Treeview) com todos os
        #           registros mais recentes do banco.
        self.bind("<F5>", lambda e: self.listar_todos())  # atualizar lista

        # -------------------------------------------------------------
        # Preenchimento inicial da tabela
        # -------------------------------------------------------------
        # Logo ao abrir o programa, chamamos o método listar_todos().
        # Esse método busca todos os registros do banco de dados e
        #           exibe os resultados na tabela (Treeview).
        # Dessa forma, o usuário vê imediatamente todos os cadastros existentes
        #           assim que a janela é carregada.
        self.listar_todos()

        # -------------------------------------------------------------
        # Encerramento correto da aplicação
        # -------------------------------------------------------------
        # O método protocol() define o comportamento quando o usuário tenta
        #            fechar a janela principal clicando no “X” no canto superior.
        # O parâmetro "WM_DELETE_WINDOW" é o evento de fechamento da janela.
        # O segundo argumento (self._fechar) é o método que será chamado nesse momento.
        # Dentro desse método (definido mais abaixo no código),
        #            a conexão com o banco de dados é fechada corretamente antes
        #            de encerrar o programa, evitando erros ou conexões pendentes.
        self.protocol("WM_DELETE_WINDOW", self._fechar)


    # -------------------------------------------------------------
    # Método utilitário: centralizar a janela na tela do usuário
    # -------------------------------------------------------------
    # O prefixo "_" no nome (_centralizar_janela) indica que este método
    #            é considerado "interno" à classe, ou seja, é usado apenas dentro
    #           da própria classe AppCadastro e não deve ser chamado de fora.
    #
    # O objetivo deste método é calcular automaticamente a posição (x, y)
    #            onde a janela deve ser exibida, para que ela apareça centralizada
    #            na tela do usuário, de forma visualmente agradável.
    # -------------------------------------------------------------
    def _centralizar_janela(self):

        """
        Centraliza a janela na tela do usuário. Chame depois que todos
        os widgets forem criados (para ter largura/altura reais).
        """

        # ---------------------------------------------------------
        # update_idletasks() é um comando do Tkinter que força a atualização
        #            de todas as tarefas pendentes da interface (como cálculos de tamanho,
        #            redimensionamento de widgets, etc.).
        # Isso garante que a largura e altura da janela obtidas logo abaixo
        #            sejam as medidas reais e não valores antigos.
        # ---------------------------------------------------------
        self.update_idletasks()

        # ---------------------------------------------------------
        # Obtém a largura atual da janela (em pixels) usando o método winfo_width().
        # Este valor é necessário para calcular a posição horizontal (x) central.
        # ---------------------------------------------------------
        largura = self.winfo_width()

        # ---------------------------------------------------------
        # Obtém a altura atual da janela (em pixels) com winfo_height().
        # Esse valor é usado para calcular a posição vertical (y) central.
        # ---------------------------------------------------------
        altura = self.winfo_height()

        # ---------------------------------------------------------
        # winfo_screenwidth() retorna a largura total da tela do monitor do usuário.
        # Com isso, podemos descobrir o "meio" da tela horizontalmente.
        # ---------------------------------------------------------
        tela_l = self.winfo_screenwidth()

        # ---------------------------------------------------------
        # winfo_screenheight() retorna a altura total da tela do monitor do usuário.
        # Isso será usado para calcular o "meio" da tela verticalmente.
        # ---------------------------------------------------------
        tela_a = self.winfo_screenheight()

        # ---------------------------------------------------------
        # Aqui é feito o cálculo da posição horizontal (x).
        # Subtrai-se a largura da janela (largura) da largura total da tela (tela_l),
        #            e o resultado é dividido por 2.
        # Isso coloca a janela exatamente no centro horizontal do monitor.
        # O operador "//" faz uma divisão inteira (sem casas decimais).
        # ---------------------------------------------------------
        x = (tela_l - largura) // 2

        # ---------------------------------------------------------
        # Aqui é feito o cálculo da posição vertical (y).
        # A lógica é parecida com a anterior, mas em vez de dividir por 2,
        #            dividimos a diferença por 3.
        # Isso faz a janela aparecer levemente acima do centro da tela,
        #            o que costuma ser mais agradável visualmente (não muito "baixo").
        # ---------------------------------------------------------
        y = (tela_a - altura) // 3  # levemente acima do centro, mais agradável

        # ---------------------------------------------------------
        # Finalmente, o método geometry() define o tamanho e a posição da janela.
        # O formato usado é:
        #     "largura x altura + posiçãoX + posiçãoY"
        # O "f-string" (f"{...}") permite inserir variáveis dentro do texto.
        # Assim, a janela é exibida com as dimensões originais (largura e altura)
        #            e centralizada nas coordenadas (x, y) calculadas acima.
        # ---------------------------------------------------------
        self.geometry(f"{largura}x{altura}+{x}+{y}")

    # ---------- layout ----------
    # -------------------------------------------------------------
    # Método: _montar_formulario
    # -------------------------------------------------------------
    # Esta função cria a parte superior da interface onde o usuário
    #            digita as informações que serão cadastradas, como:
    # ID, Nome, E-mail, Telefone e também o campo de pesquisa.
    #
    # O uso do prefixo "_" indica que este método é interno da classe
    #            (não deve ser chamado diretamente de fora da classe AppCadastro).
    # -------------------------------------------------------------
    def _montar_formulario(self):

        # ---------------------------------------------------------
        # Cria um contêiner visual chamado LabelFrame (um tipo especial
        #            de Frame que tem um título visível). Esse contêiner servirá
        #            como uma “caixa” agrupando todos os campos de entrada de dados.
        #
        # Parâmetros:
        # - self → indica que o contêiner pertence à janela principal.
        # - text="Dados da Pessoa" → define o título que aparece na borda do quadro.
        #
        # Exemplo visual na tela:
        # ┌───────────────────────────────┐
        # │     Dados da Pessoa           │
        # │  [ID:] [Nome:] [E-mail:] ...  │
        # └───────────────────────────────┘
        # ---------------------------------------------------------
        frm = ttk.LabelFrame(self, text="Dados da Pessoa")

        # ---------------------------------------------------------
        # O método pack() é usado para posicionar o LabelFrame na janela.
        #
        # Parâmetros usados:
        # - side=tk.TOP → posiciona o frame na parte superior da janela.
        # - fill=tk.X → faz o frame ocupar toda a largura horizontal (X),
        #              mas sem expandir verticalmente.
        # - padx=10 e pady=10 → criam margens (espaço em pixels)
        #              nas laterais e na parte superior/inferior, para que o frame
        #              não encoste nas bordas da janela, deixando o layout mais bonito.
        # ---------------------------------------------------------
        frm.pack(side=tk.TOP, fill=tk.X, padx=10, pady=10)

        # ---------------------------------------------------------
        # Linha 1 do formulário de cadastro
        # ---------------------------------------------------------
        # Aqui são criados e posicionados os primeiros campos:
        #            o campo de ID e o campo de Nome da pessoa.
        # Cada campo é composto por um Label (rótulo de texto)
        #            e um Entry (caixa de entrada para digitação).
        # O método grid() é usado para posicionar cada elemento
        #            em uma grade (tabela imaginária), onde cada widget ocupa
        #            uma linha (row) e uma coluna (column) específicas.
        # ---------------------------------------------------------

        # Cria um rótulo (Label) que exibirá o texto "ID:" dentro do frame chamado 'frm'.
        # Esse rótulo serve como uma legenda para que o usuário saiba
        #            que o campo ao lado mostrará o número identificador do registro.
        # É apenas um texto fixo, o usuário não interage com ele.
        #
        # Abaixo estão os parâmetros utilizados:
        #
        # frm → indica que o rótulo será colocado dentro do frame principal do formulário.        #
        #            text="ID:" → define o texto que será exibido na tela.
        # Aqui o texto literal “ID:” aparece antes do campo de entrada.        #
        # O método .grid() é usado para posicionar o rótulo em um sistema de grade (tabela invisível).        #
        # row=0 → significa que o rótulo será colocado na primeira linha (linhas começam do 0).        #
        # column=0 → indica que o rótulo será colocado na primeira coluna da grade.        #
        # padx=5 → adiciona uma margem horizontal de 5 pixels,
        #            criando espaço entre o rótulo e os elementos à esquerda/direita.        #
        # pady=5 → adiciona uma margem vertical de 5 pixels,
        #            criando espaço acima e abaixo do rótulo.        #
        # sticky="e" → define o alinhamento dentro da célula da grade.
        # A letra "e" vem de “east” (leste), o que significa alinhado à direita.
        # Isso faz o texto “ID:” ficar encostado no lado direito da célula,
        #            mais próximo do campo de entrada.
        ttk.Label(frm, text="ID:").grid(row=0, column=0, padx=5, pady=5, sticky="e")

        # Cria uma caixa de entrada (Entry), que é um campo de texto
        #            usado para exibir ou permitir a digitação de informações.
        # Nesse caso específico, o campo exibirá o valor do ID,
        #            mas o usuário não poderá alterá-lo, pois o ID é definido
        #            automaticamente pelo banco de dados.
        #
        # Parâmetros utilizados:
        #
        # frm → indica que a caixa de entrada também pertence ao frame 'frm',
        #           mantendo a hierarquia visual e lógica.
        #
        # textvariable=self.var_id → conecta a caixa de entrada a uma variável do tipo StringVar()
        #           chamada self.var_id. Essa variável funciona como um
        #           "espelho" do conteúdo do campo.
        # Se o programa alterar o valor da variável, o campo será atualizado automaticamente.
        # E, se o campo mudar (por exemplo, em outro caso onde seja editável),
        #            a variável também refletiria essa mudança.
        #
        # width=8 → define a largura do campo de entrada. O número 8 significa
        #            que ele terá espaço suficiente para cerca de 8 caracteres visíveis.
        # Como o ID é um número curto, uma largura pequena é suficiente.
        #
        # state="readonly" → torna o campo apenas de leitura,
        #            ou seja, o usuário pode ver o número do ID mas não pode digitar nada nele.
        # Isso evita que o usuário altere acidentalmente o número do registro.
        #
        # O método .grid() novamente é usado para posicionar o campo:
        #
        # row=0 → coloca o campo na mesma linha do rótulo “ID:”, de forma que fiquem lado a lado.        #
        # column=1 → coloca o campo na segunda coluna da grade (ao lado direito do rótulo).        #
        # padx=5 → adiciona espaçamento horizontal de 5 pixels entre o campo e outros elementos.        #
        # pady=5 → adiciona espaçamento vertical de 5 pixels, garantindo equilíbrio no layout.        #
        # sticky="w" → define o alinhamento dentro da célula.
        # "w" vem de “west” (oeste), o que significa alinhado à esquerda.
        # Isso faz com que o campo fique encostado no lado esquerdo da célula,
        #            logo após o rótulo “ID:”.
        ttk.Entry(frm, textvariable=self.var_id, width=8, state="readonly").grid(
            row=0, column=1, padx=5, pady=5, sticky="w")

        # Cria um rótulo (Label) que mostrará o texto “Nome:” na tela.
        # Este rótulo é uma legenda fixa usada para indicar ao usuário
        #            que o campo ao lado deve ser utilizado para digitar o nome da pessoa.
        # Ele ajuda na organização visual e na compreensão do formulário.
        # frm: define que este rótulo será adicionado dentro do frame (caixa) 'frm',
        #           que agrupa todos os campos do formulário.
        # text="Nome:": define o texto que será exibido na tela, neste caso “Nome:”.
        # grid(): organiza os componentes da tela como se fosse uma tabela invisível.
        # row=0: define que o rótulo “Nome:” ficará na primeira linha (linha 0).
        # column=2: define que ele ficará na terceira coluna (as colunas começam em 0),
        #            ou seja, duas colunas à direita do rótulo “ID:”.
        # padx=5: adiciona uma margem horizontal de 5 pixels, criando espaço nas laterais.
        # pady=5: adiciona uma margem vertical de 5 pixels, criando espaço acima e abaixo.
        # sticky="e": alinha o rótulo à direita (east), fazendo o texto “Nome:”
        #              ficar encostado no lado direito da célula, próximo à caixa de texto.
        ttk.Label(frm, text="Nome:").grid(row=0, column=2, padx=5, pady=5, sticky="e")

        # Cria uma caixa de entrada (Entry) onde o usuário poderá digitar o nome.
        # Esse campo é interativo e ligado a uma variável de controle,
        #            permitindo que o código acesse, limpe ou altere o texto digitado.
        # frm: indica que esta caixa pertence ao mesmo frame “frm” do rótulo “Nome:”.
        # textvariable=self.var_nome: conecta o campo de texto à variável self.var_nome,
        #                              que armazena dinamicamente o conteúdo digitado.
        #                              Isso significa que tudo o que o usuário digitar
        #                              poderá ser lido com self.var_nome.get().
        # width=40: define a largura do campo, permitindo digitar nomes longos.
        # grid(): posiciona o campo dentro da grade visual (ao lado do rótulo “Nome:”).
        # row=0: coloca o campo na mesma linha do rótulo.
        # column=3: posiciona o campo na quarta coluna, logo à direita do texto “Nome:”.
        # padx=5: cria um espaçamento horizontal de 5 pixels.
        # pady=5: cria um espaçamento vertical de 5 pixels.
        # sticky="w": alinha o campo à esquerda (west), fazendo-o começar logo após o rótulo.
        ttk.Entry(frm, textvariable=self.var_nome, width=40).grid(row=0, column=3, padx=5, pady=5, sticky="w")

        # Linha 2
        # Cria um novo rótulo (Label) com o texto “E-mail:” para identificar o campo de e-mail.
        # Este rótulo é apenas informativo, servindo como legenda para o campo ao lado.
        # frm: indica que o rótulo pertence ao frame principal do formulário.
        # text="E-mail:": define o texto que será exibido ao usuário.
        # grid(): posiciona o rótulo na grade (sistema de linhas e colunas invisível).
        # row=1: posiciona o rótulo na segunda linha (a contagem começa em 0).
        # column=0: posiciona o rótulo na primeira coluna da linha.
        # padx=5: cria uma margem horizontal de 5 pixels para separar o rótulo de outros elementos.
        # pady=5: cria uma margem vertical de 5 pixels para deixar o layout mais espaçado.
        # sticky="e": alinha o texto do rótulo à direita (east), aproximando-o do campo de entrada.
        ttk.Label(frm, text="E-mail:").grid(row=1, column=0, padx=5, pady=5, sticky="e")

        # Cria o campo de entrada (Entry) onde o usuário digitará o e-mail da pessoa.
        # Este campo é interativo e está ligado a uma variável de controle.
        # frm: indica que o campo pertence ao mesmo frame “frm”.
        # textvariable=self.var_email: vincula o conteúdo digitado à variável self.var_email.
        #                 Assim, o valor digitado pode ser recuperado com self.var_email.get().
        # width=40: define a largura da caixa de texto, permitindo digitar endereços de e-mail completos.
        # grid(): posiciona o campo de texto na grade do formulário.
        # row=1: coloca o campo na segunda linha, junto ao rótulo “E-mail:”.
        # column=1: posiciona o campo na segunda coluna (ao lado direito do rótulo).
        # padx=5: adiciona espaço horizontal de 5 pixels entre o campo e outros elementos.
        # pady=5: adiciona espaço vertical de 5 pixels, deixando o espaçamento uniforme.
        # sticky="w": alinha o campo à esquerda (west), mantendo o alinhamento natural de leitura.
        ttk.Entry(frm, textvariable=self.var_email, width=40).grid(row=1, column=1, padx=5, pady=5, sticky="w")

        # Cria um rótulo (Label) com o texto “Telefone:”, que serve como legenda
        # para indicar ao usuário onde ele deve digitar o número de telefone.
        # frm: define que o rótulo será inserido dentro do mesmo frame do formulário (frm),
        #           agrupando-o visualmente com os outros campos.
        # text="Telefone:": define o texto exibido na tela, neste caso “Telefone:”.
        # grid(): método que posiciona o rótulo dentro da grade do layout.
        # row=1: define que o rótulo ficará na segunda linha (lembrando que a contagem começa em 0).
        # column=2: coloca o rótulo na terceira coluna, logo após o grupo de campos de e-mail.
        # padx=5: adiciona 5 pixels de espaçamento horizontal entre o rótulo e outros elementos.
        # pady=5: adiciona 5 pixels de espaçamento vertical para manter a harmonia visual entre as linhas.
        # sticky="e": alinha o texto do rótulo à direita (east), fazendo o texto “Telefone:”
        #             ficar encostado no lado direito da célula, mais próximo da caixa de entrada.
        ttk.Label(frm, text="Telefone:").grid(row=1, column=2, padx=5, pady=5, sticky="e")

        # Cria o campo de entrada (Entry) onde o usuário digitará o número de telefone.
        # Esse campo permite a inserção de texto e está ligado a uma variável controlada.
        # frm: indica que o campo pertence ao frame “frm” do formulário.
        # textvariable=self.var_telefone: vincula o campo à variável self.var_telefone,
        #             que armazena automaticamente o valor digitado pelo usuário.
        #             Isso permite que o programa leia o valor com self.var_telefone.get()
        #             ou altere seu conteúdo com self.var_telefone.set("novo número").
        # width=22: define a largura do campo de texto — 22 caracteres são suficientes
        #             para suportar números de telefone com DDD, traços ou espaços.
        # grid(): posiciona o campo dentro da grade do frame.
        # row=1: coloca o campo na segunda linha, junto ao rótulo “Telefone:”.
        # column=3: posiciona o campo na quarta coluna, logo à direita do rótulo.
        # padx=5: adiciona um pequeno espaço horizontal de 5 pixels entre o campo e outros widgets.
        # pady=5: adiciona um pequeno espaço vertical para manter o alinhamento da linha.
        # sticky="w": alinha o campo à esquerda (west), fazendo com que comece logo após o rótulo.
        ttk.Entry(frm, textvariable=self.var_telefone, width=22).grid(row=1, column=3, padx=5, pady=5, sticky="w")

        # Linha 3
        # Cria um rótulo (Label) que exibe o texto “Pesquisar (nome, e-mail ou telefone):”.
        # Esse rótulo serve como legenda para o campo de pesquisa, indicando ao usuário
        #           que ele pode procurar por qualquer um desses três tipos de informação.
        # frm: indica que o rótulo será adicionado dentro do frame principal (frm) do formulário.
        # text="Pesquisar (nome, e-mail ou telefone):": define o texto exibido na interface,
        #           informando claramente o propósito do campo.
        # grid(): posiciona o rótulo dentro da grade que organiza os componentes na tela.
        # row=2: define que este rótulo ficará na terceira linha (a contagem começa em 0),
        #           ou seja, logo abaixo das linhas de cadastro de e-mail e telefone.
        # column=0: coloca o rótulo na primeira coluna da linha, alinhando-o à esquerda da tela.
        # padx=5: adiciona uma margem horizontal de 5 pixels, separando o rótulo de outros elementos.
        # pady=5: adiciona uma margem vertical de 5 pixels, criando espaço entre as linhas.
        # sticky="e": alinha o texto do rótulo à direita (east), encostando-o no lado direito da célula,
        #           próximo ao campo de entrada de pesquisa.
        ttk.Label(frm, text="Pesquisar (nome, e-mail ou telefone):").grid(row=2, column=0, padx=5, pady=5, sticky="e")

        # Cria o campo de entrada (Entry) onde o usuário digitará o termo de pesquisa.
        # Este campo permite buscar registros pelo nome, e-mail ou telefone cadastrados.
        # frm: define que o campo pertence ao mesmo frame “frm” do rótulo acima.
        # textvariable=self.var_pesquisa: conecta o campo à variável self.var_pesquisa,
        #                                 que armazena o texto digitado pelo usuário.
        #                                 Esse texto será usado no comando de pesquisa
        #                                 para filtrar os registros do banco de dados.
        # width=60: define a largura do campo, permitindo digitar termos longos.
        # grid(): posiciona o campo dentro da grade, na mesma linha do rótulo de pesquisa.
        # row=2: indica que o campo estará na terceira linha (abaixo dos outros campos).
        # column=1: coloca o campo na segunda coluna, logo à direita do rótulo.
        # columnspan=3: faz o campo ocupar três colunas da grade, tornando-o mais largo
        #               e centralizado, ideal para um campo de pesquisa extenso.
        # padx=5: adiciona 5 pixels de espaçamento lateral (horizontal).
        # pady=5: adiciona 5 pixels de espaçamento vertical (entre linhas).
        # sticky="w": alinha o campo à esquerda (west), começando logo após o rótulo.
        ttk.Entry(frm,
                  textvariable=self.var_pesquisa,
                  width=60).grid(row=2,
                                 column=1,
                                 columnspan=3,
                                 padx=5,
                                 pady=5,
                                 sticky="w")

        # Ajuste de colunas
        # Define a configuração da coluna número 3 dentro do frame “frm”.
        # O objetivo é permitir que essa coluna se expanda automaticamente
        #            quando a janela for redimensionada, garantindo que os campos
        #            de entrada (Entry) cresçam proporcionalmente e mantenham
        #            uma aparência agradável e responsiva.
        #
        # frm: representa o frame principal que contém o formulário.
        # grid_columnconfigure(): método que ajusta o comportamento
        #            das colunas da grade (grid layout).
        # O primeiro argumento (3) identifica qual coluna será configurada —
        #            neste caso, a quarta coluna (pois a contagem começa em 0).
        # weight=1: define o “peso” de expansão da coluna.
        # Isso significa que, quando o usuário redimensionar a janela,
        #           esta coluna se expandirá proporcionalmente mais do que as outras.
        #           Se outras colunas tiverem weight=0 (padrão), apenas esta crescerá.
        frm.grid_columnconfigure(3, weight=1)



    # Define um método interno chamado _montar_botoes, responsável
    #            por criar e organizar todos os botões da interface (como
    #            Cadastrar, Atualizar, Excluir etc.).
    # Esse método é executado durante a inicialização da janela principal
    #            e adiciona os botões dentro de um frame separado para
    #            melhor organização visual.
    def _montar_botoes(self):

        # Cria um novo frame (área/container) dentro da janela principal “self”.
        # Esse frame servirá como uma "faixa" horizontal onde os botões ficarão posicionados.
        # ttk.Frame(self): cria o frame como um elemento filho da janela principal.
        frm = ttk.Frame(self)

        # Posiciona o frame na parte superior da janela usando o método pack().
        # side=tk.TOP → define que o frame ficará no topo da janela.
        # fill=tk.X → faz o frame se expandir horizontalmente (no eixo X),
        #            ocupando toda a largura disponível da janela.
        # padx=10 → adiciona uma margem horizontal de 10 pixels nas laterais,
        #            para que o frame não encoste nas bordas da janela.
        # pady=5 → adiciona uma margem vertical de 5 pixels, separando o frame
        #           dos outros elementos acima ou abaixo dele.
        frm.pack(side=tk.TOP, fill=tk.X, padx=10, pady=5)

        # Cria um botão com o texto “Cadastrar” dentro do frame “frm”.
        # Esse botão será responsável por inserir um novo registro no banco de dados.
        # Parâmetros explicados:
        # frm → define o frame pai onde o botão será exibido.
        # text="Cadastrar" → define o texto que aparecerá no botão.
        # command=self.cadastrar → associa a ação do botão ao método self.cadastrar,
        #                           ou seja, quando o usuário clicar no botão,
        #                           o método cadastrar() será executado.
        # O método pack() define como o botão será posicionado dentro do frame:
        # side=tk.LEFT → coloca o botão alinhado à esquerda dentro do frame.
        # padx=5 → adiciona um espaçamento horizontal de 5 pixels entre este botão
        #           e outros botões que possam estar ao lado.
        ttk.Button(frm, text="Cadastrar", command=self.cadastrar).pack(side=tk.LEFT, padx=5)

        # Cria um botão com o texto “Atualizar”, usado para modificar um registro existente.
        # frm → indica que o botão pertence ao frame “frm”, onde estão todos os botões do painel.
        # text="Atualizar" → define o texto que será exibido no botão.
        # command=self.atualizar → associa o clique do botão ao método atualizar(),
        #                          que será executado quando o usuário clicar nele.
        # pack(): método usado para posicionar o botão dentro do frame.
        # side=tk.LEFT → alinha o botão à esquerda, ao lado dos outros botões.
        # padx=5 → cria um espaçamento horizontal de 5 pixels
        #           entre este botão e o anterior, garantindo uma
        #           separação visual entre eles.
        ttk.Button(frm, text="Atualizar", command=self.atualizar).pack(side=tk.LEFT, padx=5)

        # Cria o botão “Excluir”, responsável por remover um registro selecionado do banco de dados.
        # frm → indica o container (frame) onde o botão será exibido.
        # text="Excluir" → define o texto que aparece no botão.
        # command=self.excluir → define a ação a ser executada ao clicar no botão,
        #           chamando o método excluir() da classe.
        # pack(): posiciona o botão à esquerda, junto aos demais.
        # side=tk.LEFT → mantém o alinhamento em linha, da esquerda para a direita.
        # padx=5 → adiciona espaçamento horizontal entre este botão e o próximo.
        ttk.Button(frm, text="Excluir", command=self.excluir).pack(side=tk.LEFT, padx=5)

        # Cria o botão “Limpar Campos (Ctrl+N)”, usado para limpar todos os
        #           campos de texto do formulário.
        # frm → define o frame pai do botão.
        # text="Limpar Campos (Ctrl+N)" → define o texto exibido no botão,
        #           informando também o atalho de teclado.
        # command=self._limpar_campos → associa o botão ao método _limpar_campos(),
        #           que apaga o conteúdo digitado nos campos e reseta as variáveis.
        # pack(): organiza o botão dentro do frame.
        # side=tk.LEFT → mantém o alinhamento horizontal à esquerda.
        # padx=5 → cria uma pequena margem horizontal entre este botão e os outros.
        ttk.Button(frm, text="Limpar Campos (Ctrl+N)", command=self._limpar_campos).pack(side=tk.LEFT, padx=5)

        # Cria o botão “Pesquisar”, que permite ao usuário buscar
        #               registros no banco de dados.
        # frm → indica que o botão está dentro do frame de botões.
        # text="Pesquisar" → define o texto do botão visível na interface.
        # command=self.pesquisar → define que, ao clicar neste botão,
        #               o método pesquisar() será chamado,
        #               executando uma busca com base no texto digitado no campo de pesquisa.
        # pack(): organiza o botão dentro do frame.
        # side=tk.LEFT → posiciona o botão à esquerda, na mesma linha dos demais.
        # padx=5 → cria espaçamento horizontal entre este botão e os vizinhos.
        ttk.Button(frm, text="Pesquisar", command=self.pesquisar).pack(side=tk.LEFT, padx=5)

        # Cria o botão “Exportar para Excel”, responsável por gerar um arquivo Excel (.xlsx)
        #               contendo todos os registros atualmente exibidos na tabela (Treeview).
        # Esse botão facilita o compartilhamento e análise dos dados fora do programa.
        #
        # frm → indica que o botão pertence ao mesmo frame “frm” que contém os demais botões.
        #
        # text="Exportar para Excel" → define o texto exibido no botão, deixando claro
        #               para o usuário qual é a ação que será executada ao clicar.
        #
        # command=self.exportar_excel → associa o clique do botão ao método exportar_excel(),
        #               que é responsável por abrir uma janela para salvar o arquivo
        #               e criar a planilha com os dados exportados do sistema.
        #
        # pack(): define a forma como o botão será posicionado dentro do frame.
        #
        # side=tk.RIGHT → diferente dos outros botões (que estão alinhados à esquerda),
        #                este é alinhado à direita, separando visualmente as ações principais
        #                (cadastrar, atualizar, excluir) da função de exportação.
        #
        # padx=5 → adiciona um espaçamento horizontal de 5 pixels,
        #               garantindo que o botão não encoste na borda da janela.
        ttk.Button(frm,
                   text="Exportar para Excel",
                    command=self.exportar_excel).pack(side=tk.RIGHT, padx=5)

    # Define o método interno _montar_tabela, responsável por criar e configurar
    #            a tabela (Treeview) que exibirá os registros cadastrados no banco de dados.
    # Essa tabela mostrará colunas com informações como ID, nome,
    #            e-mail, telefone e data de criação.
    def _montar_tabela(self):

        # Cria um LabelFrame (um tipo especial de frame com borda e título)
        #            que servirá como contêiner visual da tabela.
        # ttk.LabelFrame(self, text="Registros") → cria um quadro com o título “Registros”.
        # Esse frame agrupa visualmente a tabela e facilita o entendimento do usuário
        #            de que ali estão listados todos os dados cadastrados.
        frm = ttk.LabelFrame(self, text="Registros")

        # Posiciona o LabelFrame dentro da janela principal usando o método pack().
        # side=tk.TOP → posiciona o frame na parte superior da área de conteúdo da janela.
        # fill=tk.BOTH → faz o frame ocupar todo o espaço disponível, tanto na largura (X)
        #                quanto na altura (Y), permitindo que ele se expanda conforme a
        #                janela é redimensionada.
        # expand=True → permite que o frame cresça e diminua proporcionalmente
        #               quando a janela muda de tamanho.
        # padx=10 → adiciona uma margem horizontal de 10 pixels nas laterais.
        # pady=10 → adiciona uma margem vertical de 10 pixels acima e abaixo,
        #               mantendo o espaçamento visual equilibrado.
        frm.pack(side=tk.TOP, fill=tk.BOTH, expand=True, padx=10, pady=10)

        # Cria uma tupla chamada “cols” contendo os nomes das colunas que a tabela terá.
        # Cada string representa o identificador interno de uma coluna no Treeview.
        # “id” → identificador numérico único do registro.
        # “nome” → nome da pessoa cadastrada.
        # “email” → endereço de e-mail do registro.
        # “telefone” → número de telefone cadastrado.
        # “criado_em” → data e hora em que o registro foi criado no banco de dados.
        cols = ("id", "nome", "email", "telefone", "criado_em")

        # Cria o componente principal da tabela, que é o ttk.Treeview.
        # O Treeview é usado em interfaces Tkinter para exibir listas,
        #                tabelas ou hierarquias de dados.
        # frm → define o frame pai onde a tabela será colocada (no caso, o frame “Registros”).
        # columns=cols → define as colunas que a tabela vai ter, baseando-se na tupla “cols”.
        # show="headings" → exibe apenas os cabeçalhos (títulos das colunas) e os dados,
        #                ocultando a coluna padrão de árvore (que aparece à esquerda).
        # selectmode="browse" → permite que o usuário selecione apenas uma linha por vez,
        #                impedindo múltiplas seleções simultâneas (ideal para editar ou excluir).
        self.tree = ttk.Treeview(frm, columns=cols, show="headings", selectmode="browse")

        # Define o cabeçalho (título da coluna) para a coluna “id”.
        # self.tree.heading() é o método que permite configurar o título, o texto visível
        #                e até ações que ocorrem ao clicar no cabeçalho.
        # "id" → nome interno da coluna que será configurada.
        # text="ID" → texto exibido no cabeçalho da coluna, visível ao usuário.
        # command=lambda: self._ordenar_por(0) → define que, ao clicar sobre o título “ID”,
        #               a tabela será ordenada pela coluna de índice 0 (a primeira).
        #               O uso de “lambda” cria uma função anônima que chama
        #               o método _ordenar_por passando o número da coluna.
        self.tree.heading("id", text="ID", command=lambda: self._ordenar_por(0))

        # Define o cabeçalho da segunda coluna, que exibirá os nomes das pessoas.
        # "nome" → identifica a coluna a ser configurada.
        # text="Nome" → texto exibido no topo da coluna.
        # command=lambda: self._ordenar_por(1) → associa a ação de clique no título da coluna
        #               ao método _ordenar_por(), passando o índice 1.
        # Isso permite que o usuário clique no cabeçalho “Nome”
        #               para ordenar a tabela alfabeticamente.
        self.tree.heading("nome", text="Nome", command=lambda: self._ordenar_por(1))

        # Configura o cabeçalho da terceira coluna, referente aos endereços de e-mail.
        # "email" → nome interno da coluna.
        # text="E-mail" → texto que aparece no título da coluna.
        # command=lambda: self._ordenar_por(2) → permite ordenar os registros
        #               por e-mail quando o usuário clica no título.
        self.tree.heading("email", text="E-mail", command=lambda: self._ordenar_por(2))

        # Define o cabeçalho da quarta coluna, que exibirá os telefones.
        # "telefone" → nome interno da coluna.
        # text="Telefone" → texto visível no topo da tabela.
        # command=lambda: self._ordenar_por(3) → define a função de ordenação
        #               quando o usuário clicar sobre o título “Telefone”.
        self.tree.heading("telefone", text="Telefone", command=lambda: self._ordenar_por(3))

        # Configura o cabeçalho da quinta coluna, que mostra a data e
        #               hora de criação do registro.
        # "criado_em" → nome interno dessa coluna no Treeview.
        # text="Criado em" → título exibido na interface para o usuário.
        # command=lambda: self._ordenar_por(4) → ao clicar no cabeçalho,
        #               chama o método de ordenação para organizar os registros
        #               conforme a data de criação.
        self.tree.heading("criado_em", text="Criado em", command=lambda: self._ordenar_por(4))

        # Configura as propriedades visuais e de alinhamento da coluna “id”.
        # self.tree.column() → método usado para definir a largura e o alinhamento de uma coluna.
        # "id" → identifica qual coluna está sendo configurada.
        # width=60 → define a largura da coluna em pixels. Como o ID é um
        #                número pequeno, 60px é suficiente.
        # anchor="center" → centraliza o conteúdo dentro da célula da
        #                coluna (tanto números quanto textos).
        self.tree.column("id", width=60, anchor="center")

        # Define as configurações da coluna “nome”, que mostrará o
        #                nome das pessoas cadastradas.
        # "nome" → nome interno da coluna.
        # width=260 → define uma largura de 260 pixels, suficiente para
        #                nomes de tamanho médio a longo.
        # anchor="w" → alinha o texto à esquerda (w = west), o que é
        #                mais natural para leitura de nomes.
        self.tree.column("nome", width=260, anchor="w")

        # Configura a coluna “email”, onde serão exibidos os endereços de e-mail das pessoas.
        # "email" → identifica a coluna.
        # width=260 → largura semelhante à do nome, pois e-mails costumam ser longos.
        # anchor="w" → mantém o texto alinhado à esquerda para melhor legibilidade.
        self.tree.column("email", width=260, anchor="w")

        # Define o formato da coluna “telefone”, responsável por mostrar os números de contato.
        # "telefone" → nome interno da coluna.
        # width=140 → largura suficiente para números de telefone com DDD e
        #                formatação (ex: (11) 98888-7777).
        # anchor="center" → centraliza os números dentro da célula,
        #                deixando o layout mais equilibrado.
        self.tree.column("telefone", width=140, anchor="center")

        # Configura a última coluna “criado_em”, que mostra a data e hora de criação do registro.
        # "criado_em" → identifica a coluna a ser configurada.
        # width=160 → define uma largura de 160 pixels, o suficiente para
        #                exibir a data e hora completas.
        # anchor="center" → centraliza o texto para manter a simetria com as
        #                colunas de dados numéricos.
        self.tree.column("criado_em", width=160, anchor="center")

        # Cria uma barra de rolagem vertical (ScrollBar) associada à tabela (Treeview).
        # ttk.Scrollbar(frm) → cria a barra de rolagem dentro do mesmo frame “frm” da tabela.
        # orient="vertical" → define que esta barra será posicionada na vertical (de cima para baixo).
        # command=self.tree.yview → conecta a barra de rolagem à função de rolagem vertical da Treeview.
        # Assim, quando o usuário arrastar a barra, a tabela rola verticalmente.
        vsb = ttk.Scrollbar(frm, orient="vertical", command=self.tree.yview)

        # Cria uma barra de rolagem horizontal, permitindo rolar os dados lateralmente.
        # ttk.Scrollbar(frm) → cria a barra dentro do frame “frm”.
        # orient="horizontal" → define a orientação da barra (da esquerda para a direita).
        # command=self.tree.xview → conecta a barra à rolagem horizontal da tabela.
        # Quando o conteúdo ultrapassar a largura visível,
        #               o usuário poderá mover a barra para ver o restante.
        hsb = ttk.Scrollbar(frm, orient="horizontal", command=self.tree.xview)

        # Configura a Treeview (tabela) para usar as barras de rolagem criadas.
        # yscroll=vsb.set → associa a barra vertical (vsb) à rolagem vertical da tabela.
        # xscroll=hsb.set → associa a barra horizontal (hsb) à rolagem horizontal.
        # Dessa forma, ao rolar a tabela com o mouse ou pelas barras,
        #                ambas permanecem sincronizadas.
        self.tree.configure(yscroll=vsb.set, xscroll=hsb.set)

        # Posiciona a barra de rolagem vertical no lado direito do frame.
        # side=tk.RIGHT → coloca a barra encostada na borda direita do frame.
        # fill=tk.Y → faz com que a barra ocupe toda a altura do frame.
        vsb.pack(side=tk.RIGHT, fill=tk.Y)

        # Posiciona a barra de rolagem horizontal na parte inferior do frame.
        # side=tk.BOTTOM → alinha a barra na parte inferior da área do frame.
        # fill=tk.X → faz com que a barra ocupe toda a largura do frame.
        hsb.pack(side=tk.BOTTOM, fill=tk.X)

        # Posiciona a tabela (Treeview) dentro do frame.
        # side=tk.LEFT → faz a tabela ocupar o lado esquerdo do frame,
        #                deixando espaço para a barra de rolagem vertical à direita.
        # fill=tk.BOTH → faz com que a tabela ocupe todo o espaço disponível
        #                tanto na largura quanto na altura do frame.
        # expand=True → permite que a tabela se expanda automaticamente
        #                quando a janela for redimensionada.
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # Aplica o estilo de “zebra” na tabela (Treeview),
        # alternando as cores de fundo das linhas (por exemplo, branco e cinza-claro)
        #                para facilitar a leitura dos registros e melhorar a estética da interface.
        # aplicar_zebra_treeview(self.tree): chama a função definida no módulo style.py,
        #               passando a tabela “self.tree” como parâmetro.
        # Essa função configura as tags “evenrow” e “oddrow” dentro da Treeview,
        #                permitindo que cada linha receba uma cor alternada automaticamente.
        aplicar_zebra_treeview(self.tree)

        # Associa um evento de duplo clique (“<Double-1>”) na tabela à função _on_duplo_clique().
        # self.tree.bind(evento, ação) → serve para detectar interações do usuário.
        # "<Double-1>" → significa “duplo clique do botão esquerdo do mouse”.
        # self._on_duplo_clique → é o método que será executado quando o evento ocorrer.
        # Nesse caso, o método captura a linha clicada e preenche automaticamente
        #                os campos do formulário (ID, Nome, E-mail e Telefone)
        #                com os dados do registro selecionado, facilitando a edição ou exclusão.
        self.tree.bind("<Double-1>", self._on_duplo_clique)

    # Define o método interno _montar_statusbar, responsável por criar a “barra de status”
    #                na parte inferior da janela principal. Essa barra
    #                serve para exibir mensagens sobre o que está acontecendo no
    #                sistema (ex: “Cadastro realizado com sucesso.”).
    def _montar_statusbar(self):

        # Cria um rótulo (Label) que exibirá informações dinâmicas de status.
        # ttk.Label(self): cria o rótulo dentro da janela principal (self).
        # textvariable=self.var_status → conecta o texto exibido à variável self.var_status,
        #               permitindo que o conteúdo seja atualizado dinamicamente
        #               por meio de self.var_status.set("mensagem").
        # anchor="w" → alinha o texto à esquerda (w = west), de forma que a mensagem
        #               comece do canto esquerdo da barra, como é comum em status bars.
        ttk.Label(self, textvariable=self.var_status, anchor="w").pack(

            # Posiciona o rótulo na parte inferior da janela.
            # side=tk.BOTTOM → coloca o componente na base (rodapé) da interface.
            # fill=tk.X → faz o rótulo ocupar toda a largura da janela,
            #             criando uma linha horizontal completa que se comporta
            #             como uma barra de status.
            side=tk.BOTTOM, fill=tk.X)

    # ---------- ações ----------
    # Define o método cadastrar, responsável por inserir um novo registro no banco de dados.
    # Esse método é acionado quando o usuário clica no botão “Cadastrar”.
    def cadastrar(self):

        # Captura o valor digitado no campo “Nome”.
        # self.var_nome → é uma variável do tipo StringVar associada ao campo de entrada.
        # get() → obtém o texto atual armazenado nessa variável.
        # strip() → remove espaços extras no início e no final do texto,
        #               evitando erros de digitação.
        nome = self.var_nome.get().strip()

        # Captura o valor digitado no campo “E-mail”.
        # get() → lê o conteúdo da variável vinculada ao campo.
        # strip() → garante que o e-mail não tenha espaços indevidos (antes ou depois do texto).
        email = self.var_email.get().strip()

        # Captura o valor digitado no campo “Telefone”.
        # get() → obtém o texto digitado pelo usuário.
        # strip() → remove espaços adicionais, deixando apenas o conteúdo
        #               principal (números, traços etc.).
        telefone = self.var_telefone.get().strip()

        # Chama a função validar() passando os três valores coletados (nome, e-mail e telefone).
        # Essa função faz uma checagem simples para garantir que:
        #   - O nome tenha pelo menos 3 caracteres.
        #   - O e-mail contenha o caractere “@”.
        #   - O telefone tenha ao menos 6 dígitos.
        # A função retorna dois valores:
        #   ok → booleano (True se tudo estiver certo, False se houver erro).
        #   msg → mensagem de erro (caso exista problema em algum campo).
        ok, msg = validar(nome, email, telefone)

        # Verifica se a validação falhou (ok == False).
        # "not ok" significa que algum dos campos obrigatórios (nome, e-mail ou telefone)
        #           não passou na verificação feita pela função validar().
        if not ok:

            # Exibe uma caixa de aviso (messagebox.showwarning) para informar ao usuário
            #           que há um erro nos dados digitados.
            # "Validação" → é o título da janela de aviso.
            # msg → é o texto retornado pela função validar() contendo a explicação do erro.
            messagebox.showwarning("Validação", msg)

            # Interrompe a execução do método, impedindo que continue para a parte de cadastro.
            # Isso evita inserir dados inválidos no banco.
            return

        # Inicia um bloco try/except para capturar possíveis erros
        #           durante a tentativa de inserção no banco de dados.
        try:

            # Chama o método inserir() do repositório (self.repo),
            #           que realiza o comando SQL INSERT no banco de dados,
            #           adicionando um novo registro com os valores de nome, e-mail e telefone.
            self.repo.inserir(nome, email, telefone)

            # Atualiza a barra de status exibindo uma mensagem de sucesso para o usuário.
            # Isso confirma que o cadastro foi feito corretamente.
            self.var_status.set("Cadastro realizado com sucesso.")

            # Atualiza a tabela exibida na tela chamando o método listar_todos(),
            #           para que o novo registro apareça imediatamente na lista.
            self.listar_todos()

            # Limpa todos os campos do formulário,
            #           deixando-os prontos para um novo cadastro.
            self._limpar_campos()

        # Caso ocorra algum erro durante o processo (como falha na conexão com o banco),
        #               o bloco “except” será executado.
        except Exception as e:

            # Exibe uma mensagem de erro detalhada usando messagebox.showerror().
            # O título “Erro ao Cadastrar” indica o tipo de falha,
            #           e a variável “e” mostra a mensagem de exceção gerada pelo Python,
            #           ajudando a identificar o problema (exemplo: falha de conexão, erro SQL etc.).
            messagebox.showerror("Erro ao Cadastrar", f"Falha ao inserir:\n\n{e}")

    # Define o método atualizar, responsável por modificar um registro existente
    #           no banco de dados com base nas informações preenchidas nos
    #           campos do formulário.
    # Esse método é acionado quando o usuário clica no botão “Atualizar”.
    def atualizar(self):

        # Verifica se existe um ID selecionado antes de tentar atualizar.
        # self.var_id.get() → obtém o valor atual do campo ID (ligado à variável tk.StringVar()).
        # Se o valor estiver vazio, significa que o usuário não selecionou nenhum registro na tabela.
        if not self.var_id.get():

            # Exibe uma caixa de informação para orientar o usuário sobre o procedimento correto.
            # messagebox.showinfo() → mostra uma mensagem simples de aviso.
            # “Atualizar” → é o título da janela de aviso.
            # O texto informa que o usuário deve selecionar um registro (por duplo clique na tabela)
            #                antes de tentar atualizar.
            messagebox.showinfo("Atualizar",
                                "Selecione um registro (duplo clique na tabela).")

            # O comando return interrompe a execução do método, impedindo que ele continue
            #                e evite tentar atualizar sem um ID válido.
            return

        # Obtém o valor atual do campo “Nome”.
        # self.var_nome.get() → lê o conteúdo da variável associada ao campo.
        # strip() → remove espaços extras no início e no fim da string.
        nome = self.var_nome.get().strip()

        # Obtém o valor atual do campo “E-mail”.
        # Assim como no cadastro, strip() remove espaços acidentais
        #                que poderiam causar falha na validação ou no banco de dados.
        email = self.var_email.get().strip()

        # Obtém o valor atual do campo “Telefone”.
        # Também utiliza strip() para garantir que o texto esteja limpo e formatado corretamente.
        telefone = self.var_telefone.get().strip()

        # Chama novamente a função validar(), passando os dados coletados do formulário:
        #                nome, email e telefone.
        # A função faz as mesmas verificações de consistência usadas no cadastro:
        # - Nome com pelo menos 3 caracteres.
        # - E-mail com o símbolo “@”.
        # - Telefone com pelo menos 6 dígitos.
        # Ela retorna dois valores:
        #   ok → True se todos os campos estiverem válidos, False se houver erro.
        #   msg → uma mensagem de erro descritiva, usada para avisar o usuário.
        ok, msg = validar(nome, email, telefone)

        # Verifica se a validação falhou (ou seja, se ok é False).
        # Se algum dos campos estiver inválido, o programa exibe um alerta e interrompe o processo.
        if not ok:

            # Mostra uma janela de aviso (warning) informando o motivo da falha.
            # "Validação" → título da janela.
            # msg → mensagem de erro retornada pela função validar().
            messagebox.showwarning("Validação", msg)

            # O comando return interrompe a execução do método,
            #                impedindo que o sistema tente atualizar um
            #                registro com dados incorretos.
            return

        # Inicia um bloco try/except para capturar possíveis erros
        #                que possam ocorrer durante a tentativa de atualização no banco de dados.
        try:

            # Chama o método atualizar() do repositório (self.repo),
            #               responsável por executar o comando SQL UPDATE na tabela.
            # int(self.var_id.get()) → converte o valor do campo ID (StringVar) para inteiro,
            #               pois o banco de dados espera um número, não texto.
            # nome, email, telefone → são os novos valores coletados dos campos do formulário,
            #               que substituirão os dados anteriores do registro selecionado.
            self.repo.atualizar(int(self.var_id.get()), nome, email, telefone)

            # Atualiza o texto exibido na barra de status,
            #                informando que a atualização foi concluída com sucesso.
            self.var_status.set("Registro atualizado com sucesso.")

            # Atualiza a lista de registros da tabela chamando o método listar_todos(),
            #                garantindo que o novo conteúdo do banco apareça imediatamente na tela.
            self.listar_todos()

            # Limpa os campos do formulário após a atualização,
            #                deixando-os prontos para uma nova operação.
            self._limpar_campos()

        # Caso ocorra um erro durante o processo (por exemplo, falha de conexão,
        #                erro no SQL ou valor inválido), o bloco “except” é executado.
        except Exception as e:

            # Exibe uma mensagem de erro informando ao usuário o motivo da falha.
            # messagebox.showerror() → cria uma janela com o ícone de erro.
            # “Erro ao Atualizar” → título da janela.
            # f"Falha ao atualizar:\n\n{e}" → exibe uma mensagem detalhada, incluindo o erro técnico.
            messagebox.showerror("Erro ao Atualizar", f"Falha ao atualizar:\n\n{e}")

    # Define o método excluir, responsável por remover um registro do banco de dados.
    # Esse método é acionado quando o usuário clica no botão “Excluir”.
    def excluir(self):

        # Verifica se o campo ID está vazio.
        # self.var_id.get() → retorna o valor do campo ID, vinculado à variável tk.StringVar().
        # Caso o valor esteja vazio (nenhum registro selecionado),
        #                significa que o usuário ainda não escolheu uma
        #                linha da tabela para exclusão.
        if not self.var_id.get():

            # Exibe uma mensagem de informação avisando que é necessário
            # selecionar um registro antes de tentar excluir.
            # "Excluir" → é o título da janela exibida.
            # A mensagem explica que o usuário deve clicar duas vezes na tabela
            #                para carregar o registro que deseja excluir.
            messagebox.showinfo("Excluir",
                                "Selecione um registro (duplo clique na tabela).")

            # Interrompe a execução do método, evitando que a exclusão prossiga sem um ID válido.
            return

        # Exibe uma caixa de confirmação antes de excluir o registro.
        # messagebox.askyesno() → exibe uma janela com duas opções: “Sim” e “Não”.
        # Retorna True se o usuário clicar em “Sim”, e False se clicar em “Não”.
        # “Confirmação” → título da janela de aviso.
        # A mensagem pergunta ao usuário se ele realmente deseja excluir o registro.
        if not messagebox.askyesno("Confirmação",
                                   "Deseja realmente excluir este registro?"):

            # Caso o usuário escolha “Não”, o método retorna imediatamente,
            #                cancelando a exclusão e mantendo o registro no banco de dados.
            return

        # Inicia um bloco try/except para executar a exclusão e capturar erros,
        #                garantindo que o programa não trave caso ocorra algum problema.
        try:

            # Chama o método excluir() do repositório (self.repo),
            #                que executa um comando SQL DELETE no banco de dados.
            # int(self.var_id.get()) → converte o valor do campo ID (que é texto)
            #               em número inteiro, pois o banco precisa do ID numérico
            #               para identificar o registro a ser excluído.
            self.repo.excluir(int(self.var_id.get()))

            # Atualiza a barra de status exibindo uma mensagem de confirmação ao usuário,
            #                informando que o registro foi excluído com sucesso.
            self.var_status.set("Registro excluído com sucesso.")

            # Atualiza a tabela chamando o método listar_todos(),
            #                removendo da tela o registro que acabou de ser excluído.
            self.listar_todos()

            # Limpa os campos do formulário (Nome, E-mail, Telefone, etc.)
            #                para evitar que o usuário veja dados de um registro que já foi apagado.
            self._limpar_campos()

        # Caso aconteça algum erro (como perda de conexão com o banco,
        #                problema de permissão ou erro SQL), o bloco except será executado.
        except Exception as e:

            # Exibe uma mensagem de erro informando o motivo da falha.
            # messagebox.showerror() → cria uma janela de erro com ícone vermelho.
            # “Erro ao Excluir” → título da janela.
            # f"Falha ao excluir:\n\n{e}" → mostra o erro técnico detalhado (exceção “e”),
            #                ajudando a identificar o problema durante o desenvolvimento.
            messagebox.showerror("Erro ao Excluir", f"Falha ao excluir:\n\n{e}")


    # Define o método pesquisar, responsável por realizar buscas na base de dados
    #                com base no texto digitado pelo usuário no campo de pesquisa.
    # Ele permite filtrar registros por nome, e-mail ou telefone.
    def pesquisar(self):

        """
        Se houver texto no campo de pesquisa, filtra por nome/e-mail/telefone (contém).
        Se o campo estiver vazio, lista todos os registros.
        """

        # Captura o texto digitado pelo usuário no campo de pesquisa.
        # self.var_pesquisa.get() → obtém o conteúdo da variável vinculada ao campo.
        # strip() → remove espaços em branco no início e no final do texto.
        # Isso evita problemas caso o usuário digite espaços acidentalmente.
        termo = self.var_pesquisa.get().strip()

        # Usa um bloco try/except para evitar que erros travem o programa
        #                caso haja falhas de conexão com o banco ou erros inesperados.
        try:

            # Verifica se o campo de pesquisa está vazio.
            if not termo:

                # Se o campo estiver vazio, busca todos os registros no banco de dados.
                # self.repo.listar_todos() → retorna uma lista com todos os registros cadastrados.
                rows = self.repo.listar_todos()

                # Preenche a tabela (Treeview) com os dados retornados,
                #                aplicando o efeito “zebra” para melhor leitura.
                preencher_treeview_com_zebra(self.tree, rows)

                # Atualiza a barra de status, mostrando quantos registros foram encontrados.
                # len(rows) → conta o total de linhas retornadas pela consulta.
                self.var_status.set(f"Total: {len(rows)} registro(s).")

                # Interrompe a execução do método aqui, já que todos os registros já foram listados.
                return

            # Caso o campo de pesquisa contenha algum texto, realiza uma busca filtrada.
            # self.repo.pesquisar(termo) → faz uma consulta SQL usando o termo como filtro,
            #                verificando se o texto aparece no nome, e-mail ou telefone.
            rows = self.repo.pesquisar(termo)

            # Preenche a tabela (Treeview) com os resultados da pesquisa.
            preencher_treeview_com_zebra(self.tree, rows)

            # Atualiza a barra de status, informando que a pesquisa foi concluída
            #                e mostrando quantos registros correspondem ao termo digitado.
            self.var_status.set(f"Pesquisa concluída: {len(rows)} registro(s).")

        # Caso ocorra algum erro durante o processo (como falha na conexão com o banco),
        #                o bloco “except” captura o erro e exibe uma mensagem amigável ao usuário.
        except Exception as e:

            # Exibe uma janela de erro detalhando o motivo da falha.
            # messagebox.showerror() → cria uma janela de erro com ícone vermelho.
            # “Erro na Pesquisa” → título da janela.
            # f"Falha ao pesquisar:\n\n{e}" → mensagem mostrando o erro técnico completo.
            messagebox.showerror("Erro na Pesquisa", f"Falha ao pesquisar:\n\n{e}")


    # Define o método listar_todos, responsável por buscar e exibir todos os registros
    #                salvos no banco de dados, sem aplicar nenhum tipo de filtro.
    # Esse método é chamado, por exemplo, ao iniciar o programa ou após cadastrar,
    #                atualizar ou excluir um registro.
    def listar_todos(self):

        # Inicia um bloco try/except para tratar erros de forma segura.
        # Isso evita que o programa trave caso haja algum problema na conexão
        #                com o banco ou no carregamento dos dados.
        try:

            # Chama o método listar_todos() do repositório (self.repo),
            #                que executa uma consulta SQL (SELECT) no banco de dados,
            # retornando uma lista (rows) com todos os registros encontrados.
            rows = self.repo.listar_todos()

            # Chama a função preencher_treeview_com_zebra(), responsável por:
            # - Limpar a tabela (Treeview) antes de inserir novos dados.
            # - Inserir cada linha da lista “rows” na tabela.
            # - Aplicar o efeito visual “zebra” (linhas alternadas em cores diferentes),
            #                  o que facilita a leitura dos dados na interface.
            preencher_treeview_com_zebra(self.tree, rows)

            # Atualiza a barra de status na parte inferior da janela,
            #                exibindo o total de registros carregados.
            # len(rows) → retorna o número total de registros encontrados no banco.
            self.var_status.set(f"Total: {len(rows)} registro(s).")

        # Caso ocorra algum erro (por exemplo, falha na conexão com o banco ou erro SQL),
        #                o bloco except será executado.
        except Exception as e:

            # Exibe uma janela de erro informando que a listagem falhou.
            # messagebox.showerror() → cria uma janela de alerta visual com ícone vermelho.
            # “Erro na Listagem” → título da janela de erro.
            # f"Falha ao listar:\n\n{e}" → exibe a mensagem de erro técnica (objeto da exceção “e”)
            #                para auxiliar no diagnóstico durante o desenvolvimento.
            messagebox.showerror("Erro na Listagem", f"Falha ao listar:\n\n{e}")

    # ---------- helpers UI ----------
    # Define o método _on_duplo_clique, responsável por tratar o evento de duplo clique na tabela (Treeview).
    # Quando o usuário dá dois cliques em uma linha, o sistema carrega os dados dessa linha
    #                nos campos do formulário, permitindo edição ou exclusão.
    # O nome começa com "_" (underline) para indicar que é um método interno (privado) da classe.
    def _on_duplo_clique(self, _event):

        # Obtém o identificador (ID interno do Treeview) do item atualmente selecionado pelo usuário.
        # self.tree.focus() → retorna o "item ID" do elemento em foco (clicado).
        # Se o usuário ainda não clicou em nenhuma linha, retorna uma string vazia.
        item = self.tree.focus()

        # Verifica se há algum item selecionado na tabela.
        # Caso o usuário tenha clicado fora das linhas, “item” estará vazio.
        if not item:

            # Se nenhum item estiver selecionado, o método é encerrado imediatamente
            #                com “return”, evitando erros ao tentar acessar dados inexistentes.
            return

        # Obtém todos os valores (as colunas) associados à linha selecionada.
        # self.tree.item(item, "values") → retorna uma tupla com os dados da linha clicada,
        #                por exemplo: ("1", "Maria", "maria@email.com",
        #                "9999-9999", "2030-11-04 17:00").
        vals = self.tree.item(item, "values")

        # Verifica se a variável “vals” contém valores.
        # Isso garante que a linha selecionada realmente tenha dados válidos.
        if not vals:

            # Caso não haja valores (por algum erro de leitura ou linha vazia),
            #                o método é interrompido para evitar erros no restante do código.
            return

        # Desempacota os valores retornados da linha selecionada (vals) em variáveis separadas.
        # Cada posição da tupla "vals" corresponde a uma coluna da tabela (Treeview):
        #   vals[0] → ID do registro.
        #   vals[1] → Nome da pessoa.
        #   vals[2] → E-mail da pessoa.
        #   vals[3] → Telefone.
        #   vals[4] → Data de criação (gerada automaticamente pelo banco).
        # O uso do underline (_) em "_id" e "_criado" indica que essas
        #                variáveis são usadas apenas internamente.
        _id, nome, email, telefone, _criado = vals

        # Define o valor do campo de texto vinculado à variável self.var_id.
        # O método .set() atualiza automaticamente o conteúdo exibido no campo "ID" da interface.
        # Isso preenche o formulário com o ID do registro selecionado na tabela.
        self.var_id.set(_id)

        # Define o valor do campo "Nome" na interface com o nome selecionado.
        # O valor digitado é vinculado à variável self.var_nome.
        self.var_nome.set(nome)

        # Define o valor do campo "E-mail" da interface, preenchendo automaticamente
        #                com o e-mail do registro clicado.
        self.var_email.set(email)

        # Define o valor do campo "Telefone", exibindo o número da pessoa selecionada.
        # Assim, o usuário pode editar ou excluir o registro diretamente.
        self.var_telefone.set(telefone)

    # Define o método _limpar_campos, responsável por limpar todos os campos do formulário
    #                e remover qualquer seleção ativa na tabela (Treeview).
    # Esse método é usado após cadastrar, atualizar, excluir ou ao
    #                pressionar o atalho Ctrl+N.
    def _limpar_campos(self):

        # Limpa o campo "ID".
        # O método .set("") define o valor da variável vinculada (self.var_id) como vazio.
        # Isso faz com que o campo de texto correspondente na interface fique em branco.
        self.var_id.set("")

        # Limpa o campo "Nome", removendo qualquer texto digitado anteriormente.
        self.var_nome.set("")

        # Limpa o campo "E-mail", apagando o conteúdo do campo na interface.
        self.var_email.set("")

        # Limpa o campo "Telefone", deixando-o vazio para uma nova digitação.
        self.var_telefone.set("")

        # Remove a seleção atual da tabela (Treeview), caso o usuário
        #                tenha clicado em alguma linha.
        # self.tree.selection() retorna uma lista com os IDs internos dos itens selecionados.
        # self.tree.selection_remove(...) desmarca esses itens visualmente,
        #                garantindo que nenhuma linha continue destacada após a limpeza.
        self.tree.selection_remove(self.tree.selection())

    # Define o método _ordenar_por, responsável por ordenar os dados da tabela (Treeview)
    #                quando o usuário clica no cabeçalho de uma coluna.
    # Esse método alterna entre ordem crescente e decrescente a cada clique.
    # O parâmetro "coluna_idx" indica qual coluna será usada para a
    #                ordenação (exemplo: 0 = ID, 1 = Nome).
    def _ordenar_por(self, coluna_idx: int):
        """Ordena a tabela por coluna (toggle asc/desc) localmente."""

        # Cria uma lista chamada "dados" contendo os valores de todas as linhas da tabela.
        # self.tree.get_children() retorna todos os itens (linhas) do Treeview.
        # Para cada item, self.tree.item(i, "values") obtém uma tupla com os valores de suas colunas.
        # Assim, "dados" se torna uma lista de tuplas, exemplo:
        # [("1", "Ana", "ana@email.com", "9999-0000"), ("2", "Bruno",
        #                "bruno@email.com", "8888-1111"), ...]
        dados = [self.tree.item(i, "values") for i in self.tree.get_children()]

        # Se a lista estiver vazia (não há registros na tabela),
        #                o método é encerrado imediatamente para evitar erros.
        if not dados:
            return

        # Recupera o valor atual do atributo "_ord_asc" (indica se a
        #                última ordenação foi crescente ou não).
        # getattr(self, "_ord_asc", True):
        #   → tenta obter o atributo "_ord_asc" do objeto atual (self).
        #   → se o atributo não existir ainda (primeira execução), usa True como valor padrão.
        # Isso significa que a primeira ordenação será sempre em ordem crescente.
        asc = getattr(self, "_ord_asc", True)

        # Inverte o valor de "_ord_asc" para alternar entre as duas direções:
        # - Se a ordenação atual for crescente (True), muda para decrescente (False).
        # - Se for decrescente (False), muda para crescente (True).
        # Assim, cada clique no cabeçalho da coluna troca a direção da ordenação.
        self._ord_asc = not asc

        # Define uma função interna chamada "chave", usada para indicar ao Python
        #                qual valor deve ser considerado como base da ordenação.
        # Essa função será passada como argumento "key" para a função sorted().
        def chave(row):

            # Obtém o valor da coluna específica usada para ordenação.
            # "row" é uma tupla representando uma linha da tabela.
            # "coluna_idx" é o índice da coluna clicada (0 para ID, 1 para Nome, etc.).
            v = row[coluna_idx]

            # Se a coluna clicada for a primeira (índice 0 → ID),
            #                o sistema tenta converter o valor em número inteiro.
            # Isso garante que a ordenação de IDs ocorra de forma numérica (1, 2, 10)
            #                em vez de textual (1, 10, 2), o que aconteceria se
            #                fossem comparados como strings.
            if coluna_idx == 0:

                try:

                    # Retorna o valor convertido para inteiro (usado como chave de ordenação).
                    return int(v)

                # Caso ocorra erro na conversão (exemplo: valor vazio ou não numérico),
                #                retorna 0, garantindo que a função não quebre.
                except:
                    return 0

            # Para as demais colunas (como Nome, E-mail e Telefone),
            #                converte o valor para string e deixa todas as
            #                letras minúsculas com .lower().
            # Isso garante que a ordenação ignore diferenças entre maiúsculas e minúsculas.
            return str(v).lower()

        # Cria uma nova lista chamada "ordenado" com os dados classificados.
        # A função sorted() ordena a lista "dados" de acordo com o valor retornado pela função "chave".
        # Parâmetros:
        #   key=chave → define que a função "chave" será usada para determinar o critério de ordenação.
        #   reverse=asc is False → se asc for False, a ordenação será descendente; se for True, ascendente.
        ordenado = sorted(dados, key=chave, reverse=asc is False)

        # Atualiza a tabela (Treeview) com a lista de registros ordenada.
        # A função preencher_treeview_com_zebra() limpa a tabela e reinsere os dados,
        #                aplicando novamente o efeito “zebra” para alternar as cores das linhas.
        preencher_treeview_com_zebra(self.tree, ordenado)

    # ---------- exportação ----------
    # Define o método exportar_excel, responsável por exportar todos os registros
    #                exibidos na tabela (Treeview) para um arquivo Excel (.xlsx).
    # Essa funcionalidade é útil para gerar relatórios, salvar backups ou
    #                compartilhar os dados.
    def exportar_excel(self):

        # Cria uma lista chamada "linhas" contendo os dados de todas as linhas da tabela (Treeview).
        # self.tree.get_children() retorna os identificadores (IDs internos) de todas as linhas da tabela.
        # Para cada identificador “i”, o comando self.tree.item(i, "values") obtém os valores de cada coluna,
        #                retornando-os como uma tupla, por exemplo: ("1", "Maria", "maria@email.com",
        #                "9999-0000", "2030-11-04").
        # No final, a lista "linhas" conterá uma estrutura como:
        # [
        #   ("1", "Maria", "maria@email.com", "9999-0000", "2030-11-04"),
        #   ("2", "João", "joao@email.com", "9888-1111", "2030-11-02"),
        #   ...
        # ]
        linhas = [self.tree.item(i, "values") for i in self.tree.get_children()]

        # Verifica se a lista "linhas" está vazia (ou seja, se não há nenhum dado para exportar).
        # Isso evita que o sistema tente criar um arquivo Excel em branco.
        if not linhas:

            # Exibe uma caixa de informação avisando o usuário de que não há nada a exportar.
            # messagebox.showinfo() → cria uma janela de aviso informativo (ícone azul).
            # "Exportar para Excel" → título da janela.
            # "Nada para exportar." → mensagem exibida ao usuário.
            messagebox.showinfo("Exportar para Excel", "Nada para exportar.")

            # O comando “return” interrompe a execução do método imediatamente,
            #                evitando que o restante do código (que cria o
            #                arquivo Excel) seja executado.
            return

        # Abre uma janela de diálogo para o usuário escolher onde salvar o arquivo Excel.
        # O método filedialog.asksaveasfilename() pertence ao módulo tkinter.filedialog
        #                e permite que o usuário navegue pelas pastas do computador e
        #                defina o nome do arquivo de saída.
        caminho = filedialog.asksaveasfilename(

            # Define a extensão padrão do arquivo gerado.
            # Caso o usuário não digite uma extensão manualmente, o arquivo será salvo com “.xlsx”,
            #                que é o formato padrão usado pelo Excel (OpenXML Workbook).
            defaultextension=".xlsx",

            # Define os tipos de arquivos aceitos na caixa de diálogo.
            # Aqui, apenas arquivos com extensão “.xlsx” serão exibidos no filtro do explorador.
            # O primeiro valor ("Arquivo Excel") é apenas uma descrição
            #                legível mostrada ao usuário.
            filetypes=[("Arquivo Excel", "*.xlsx")],

            # Define o título exibido na janela do diálogo de salvamento.
            # Esse texto aparece na barra de título da janela, orientando o
            #                usuário sobre o que está fazendo.
            title="Salvar como",

        )

        # Após o usuário escolher (ou não) o local e o nome do arquivo, a
        #                função retorna o caminho completo
        # (por exemplo: “C:/Usuários/João/Documentos/relatorio.xlsx”).
        # Se o usuário clicar em “Cancelar”, o método retorna uma string vazia ("").
        # O código abaixo verifica se o usuário cancelou a operação.
        if not caminho:

            # Se nenhum caminho for escolhido (ou seja, o usuário cancelou a ação),
            #                o método retorna imediatamente, interrompendo a execução da exportação.
            # Isso impede que o programa tente criar um arquivo sem um nome ou diretório válido.
            return

        # Inicia um bloco try/except para tratar erros durante a criação e gravação do arquivo Excel.
        # Isso impede que o programa trave caso ocorra algum problema (como permissão
        #                de pasta, arquivo em uso, etc.).
        try:

            # Cria um novo arquivo Excel em memória utilizando a biblioteca openpyxl.
            # Workbook() cria um objeto representando um arquivo Excel vazio (.xlsx),
            #                que será posteriormente salvo no caminho definido pelo usuário.
            wb = Workbook()

            # Obtém a planilha ativa (a primeira criada automaticamente dentro do novo arquivo).
            # Por padrão, o Excel cria uma aba chamada “Sheet”.
            # Aqui, a variável "ws" é usada para manipular essa planilha diretamente.
            ws = wb.active

            # Altera o nome da aba padrão para “Pessoas”,
            # tornando o arquivo mais legível e descritivo para o usuário.
            ws.title = "Pessoas"

            # Cria uma lista chamada “cab”, contendo os nomes das colunas do cabeçalho.
            # Essas colunas correspondem às informações exibidas na interface (Treeview),
            #                garantindo que o arquivo Excel tenha os mesmos campos:
            #   ID → Identificador único da pessoa.
            #   Nome → Nome completo da pessoa cadastrada.
            #   E-mail → Endereço de e-mail.
            #   Telefone → Número de telefone.
            #   Criado em → Data/hora do cadastro no banco de dados.
            cab = ["ID", "Nome", "E-mail", "Telefone", "Criado em"]

            # Adiciona o cabeçalho (lista "cab") como a primeira linha da planilha.
            # O método append() insere os valores em uma nova linha automaticamente.
            # Resultado esperado na planilha:
            # ┌────┬────────┬──────────────────────┬──────────────┬────────────────────┐
            # │ ID │ Nome   │ E-mail               │ Telefone     │ Criado em          │
            # └────┴────────┴──────────────────────┴──────────────┴────────────────────┘
            ws.append(cab)

            # Percorre a lista “linhas”, que contém todos os registros coletados da Treeview.
            # Cada item de “linhas” é uma tupla representando um registro completo
            # (exemplo: ("1", "Maria", "maria@email.com", "9999-0000", "2030-11-04 17:00")).
            for row in linhas:

                # Converte cada tupla em uma lista (para garantir compatibilidade com o método append)
                #                e adiciona os valores como uma nova linha na planilha.
                # Assim, cada registro do sistema é inserido em uma linha do Excel.
                ws.append(list(row))

                # Cria uma lista chamada “larguras” com os valores numéricos de largura para cada coluna.
                # Esses números representam quantos caracteres cabem aproximadamente em cada coluna.
                # A ideia é ajustar automaticamente o tamanho das colunas para que o conteúdo fique visível.
                # Cada número da lista corresponde a uma coluna do Excel, na mesma ordem do cabeçalho:
                #   6  → largura da coluna "ID"
                #   32 → largura da coluna "Nome"
                #   36 → largura da coluna "E-mail"
                #   18 → largura da coluna "Telefone"
                #   22 → largura da coluna "Criado em"
                larguras = [6, 32, 36, 18, 22]

                # Usa um laço "for" com enumerate() para percorrer a lista de larguras.
                # O parâmetro "start=1" faz com que a contagem comece em 1 (e não em 0),
                #                pois as colunas no Excel são numeradas a partir de 1 (A=1, B=2, C=3...).
                for idx, largura in enumerate(larguras, start=1):

                    # A função get_column_letter(idx) converte o número da coluna
                    #                para a letra correspondente:
                    # 1 → "A", 2 → "B", 3 → "C", etc.
                    # ws.column_dimensions[...] acessa as propriedades de largura
                    #                da coluna correspondente.
                    # O atributo “.width” define a largura da coluna no Excel.
                    ws.column_dimensions[get_column_letter(idx)].width = largura

            # Salva o arquivo Excel no caminho que o usuário escolheu anteriormente.
            # wb.save(caminho) escreve todo o conteúdo (planilha, cabeçalho e dados) no arquivo físico.
            # Exemplo: “C:/Usuários/João/Documentos/Cadastro_Pessoas.xlsx”.
            wb.save(caminho)

            # Exibe uma mensagem de sucesso para o usuário, informando o local onde o arquivo foi salvo.
            # O "\n" insere uma quebra de linha para deixar a mensagem mais legível.
            messagebox.showinfo("Exportar para Excel", f"Arquivo salvo com sucesso em:\n{caminho}")

            # Atualiza a barra de status na interface principal, exibindo o
            #                caminho do arquivo exportado.
            # Isso serve como confirmação visual na parte inferior da janela.
            self.var_status.set(f"Exportado: {caminho}")

        # Caso ocorra qualquer erro durante o processo (por exemplo, o usuário tentar salvar
        #                em um local protegido ou o arquivo já estiver aberto
        #                no Excel), o bloco “except” é executado automaticamente.
        except Exception as e:

            # Exibe uma mensagem de erro para o usuário, informando que a exportação falhou.
            # O f-string inclui o conteúdo da exceção “e”, que mostra a causa técnica do erro.
            messagebox.showerror("Erro ao Exportar", f"Falha ao exportar:\n\n{e}")

    # ---------- encerrar ----------
    # Define o método _fechar, responsável por encerrar o programa de
    #                 forma segura e controlada.
    # Ele é executado automaticamente quando o usuário tenta fechar a
    #                janela principal (clicando no "X").
    def _fechar(self):

        # Inicia um bloco try/except para garantir que o encerramento ocorra sem erros,
        #                mesmo que algo dê errado ao tentar fechar a conexão com o banco de dados.
        try:

            # Verifica se existe uma conexão ativa com o banco de dados (self.conn).
            # Se a variável “self.conn” não for None, significa que há uma conexão aberta.
            if self.conn:

                # Fecha a conexão com o banco de dados MySQL para liberar os recursos do sistema.
                # Isso é importante porque conexões abertas consomem memória e podem travar o banco
                #                se não forem encerradas corretamente.
                self.conn.close()

        # Caso ocorra qualquer erro ao tentar fechar a conexão (por exemplo,
        #                se ela já estiver fechada ou nunca tiver sido aberta), o
        #                bloco “except” captura a exceção e executa o comando “pass”, que
        #                simplesmente ignora o erro sem exibir mensagens.
        except:
            pass

        # Encerra a janela principal do programa, destruindo todos os widgets (botões,
        #                campos, tabelas, etc.) e liberando completamente os
        #                recursos utilizados pela interface gráfica.
        # O método “destroy()” é o equivalente, no Tkinter, a “fechar o aplicativo”.
        self.destroy()