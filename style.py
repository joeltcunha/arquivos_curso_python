# Importa o módulo principal do Tkinter, que fornece as ferramentas
#                básicas para criar interfaces gráficas.
# O alias “tk” é usado para acessar os componentes padrão (como
#                janela principal, botões e rótulos).
import tkinter as tk

# Importa o submódulo “ttk” (Themed Tkinter), que contém widgets com aparência moderna,
#                como botões, labels e tabelas (Treeview). O ttk permite
#                aplicar temas e estilos personalizados.
from tkinter import ttk


# Define uma função chamada “aplicar_estilo” que recebe um parâmetro “root”.
# O parâmetro “root” é o objeto da janela principal (criado como tk.Tk()).
# Essa função personaliza o estilo visual da interface — fontes,
#                espaçamento, tema e aparência geral.
def aplicar_estilo(root: tk.Tk):

    """
    Aplica um tema limpo e consistente ao ttk.
    - Tema 'clam'
    - Botões e campos com padding
    - Treeview com zebra (linhas alternadas)
    """

    # Cria um objeto de estilo (Style) vinculado à janela principal “root”.
    # Esse objeto controla todas as propriedades visuais dos componentes ttk.
    # Ele permite alterar cores, tamanhos, fontes e o tema visual aplicado ao programa inteiro.
    style = ttk.Style(root)

    # Define o tema visual base do programa.
    # O método “theme_use” aplica o tema especificado — neste caso, “clam”.
    # O tema “clam” é amplamente utilizado por ser simples, neutro e compatível com Windows e Linux.
    # Ele serve como base para as modificações personalizadas feitas depois.
    try:
        style.theme_use("clam")

    # O bloco “except” captura o erro tk.TclError caso o tema “clam” não exista no sistema do usuário.
    # Isso pode acontecer em versões antigas do Tkinter ou em alguns ambientes Linux.
    # Nesse caso, o código simplesmente ignora o erro e mantém o tema padrão,
    #                garantindo que o programa continue funcionando.
    except tk.TclError:
        pass


    # Fonte e tamanhos (ajuste se quiser maior)
    # Define uma tupla chamada “fonte_padrao”, que será usada como fonte principal da interface.
    # A tupla contém dois valores:
    #   - "Segoe UI" → nome da fonte (é a fonte moderna usada no Windows).
    #   - 10 → tamanho da fonte em pontos.
    # Essa fonte é escolhida por ser limpa, legível e padrão de muitos aplicativos modernos.
    fonte_padrao = ("Segoe UI", 10)

    # Usa o método “option_add” para aplicar a fonte padrão a todos os widgets da janela.
    # "*Font" significa que a configuração se aplica globalmente a todos os componentes
    #                que exibem texto (Labels, Buttons, Entrys, etc.).
    # Dessa forma, todos os textos da interface terão a mesma aparência e tamanho.
    root.option_add("*Font", fonte_padrao)

    # ---------- Estilização dos Widgets Individuais ----------

    # Aplica um espaçamento interno (padding) de 2 pixels em cada direção (horizontal e vertical)
    #                aos rótulos (Label). Isso evita que o texto fique colado nas
    #                bordas e melhora a legibilidade.
    style.configure("TLabel", padding=(2, 2))

    # Aplica padding de 4 pixels aos campos de texto (Entry).
    # Isso aumenta a distância entre o texto digitado e as bordas do campo,
    #                tornando-o mais confortável de usar.
    style.configure("TEntry", padding=(4, 4))

    # Define padding de 6 pixels na horizontal e 4 na vertical para os botões (Button).
    # Isso cria botões com aparência equilibrada e confortável ao clique.
    style.configure("TButton", padding=(6, 4))

    # Adiciona espaçamento interno (8,8) ao contêiner “LabelFrame” — um tipo de
    #                moldura usada para agrupar elementos.
    # O padding interno garante que os componentes dentro do frame não fiquem colados nas bordas.
    style.configure("TLabelframe", padding=(8, 8))

    # Personaliza o texto do título do LabelFrame, aplicando uma fonte semibold (semi-negrito).
    # Isso ajuda a destacar o título do grupo, facilitando a leitura da seção.
    style.configure("TLabelframe.Label", font=("Segoe UI Semibold", 10))

    # ---------- Personalização do Cabeçalho da Tabela (Treeview) ----------

    # Configura o estilo do cabeçalho da tabela Treeview (a parte superior com os nomes das colunas).
    # A fonte “Segoe UI Semibold” em tamanho 10 torna o cabeçalho mais visível e profissional.
    # Essa diferenciação visual ajuda o usuário a distinguir facilmente o cabeçalho dos dados.
    style.configure("Treeview.Heading", font=("Segoe UI Semibold", 10))


# Define uma função chamada “aplicar_zebra_treeview” que recebe um parâmetro chamado “tree”.
# Esse parâmetro representa o widget Treeview (a tabela de dados do Tkinter).
# O objetivo dessa função é aplicar o efeito visual de “zebra”, ou
#                seja, alternar as cores das linhas (uma linha clara e outra um
#                pouco mais escura), o que facilita a leitura dos dados.
def aplicar_zebra_treeview(tree: ttk.Treeview):

    """
    Aplica 'zebra' na Treeview usando tags 'oddrow'/'evenrow'.
    Chame após repopular a tabela.
    """

    # Configura as propriedades visuais (cor de fundo) da tag “oddrow”.
    # Essa tag será atribuída às linhas ímpares (1ª, 3ª, 5ª, etc.).
    # A cor “#F7F7F7” é um cinza bem claro, escolhido para criar contraste suave com o branco.
    # O método “tag_configure” permite associar estilos visuais a tags
    #                 específicas dentro da Treeview.
    tree.tag_configure("oddrow", background="#F7F7F7")

    # Configura a tag “evenrow”, que será aplicada às linhas pares (2ª, 4ª, 6ª, etc.).
    # Define o fundo como branco puro “#FFFFFF”.
    # Quando as duas tags são aplicadas alternadamente, o resultado visual é uma tabela listrada
    # — semelhante às planilhas do Excel — facilitando a leitura horizontal dos dados.
    tree.tag_configure("evenrow", background="#FFFFFF")


# Define uma função chamada “preencher_treeview_com_zebra” que recebe dois parâmetros:
# 1. tree → o widget Treeview (a tabela visual onde os dados serão exibidos);
# 2. rows → uma lista contendo as linhas de dados (vindas do banco de dados).
# Essa função tem como objetivo limpar a tabela atual e, em seguida, preencher novamente
#                os dados aplicando o efeito “zebra” (linhas alternadas em duas cores diferentes).
def preencher_treeview_com_zebra(tree: ttk.Treeview, rows):

    """
    Limpa e preenche a Treeview aplicando zebra automática.
    """

    # ------------------- ETAPA 1: LIMPAR OS DADOS EXISTENTES -------------------

    # O método get_children() retorna uma lista com todos os identificadores (IDs)
    #                das linhas atualmente exibidas na Treeview.
    # O laço “for” percorre cada uma dessas linhas.
    for i in tree.get_children():

        # O método delete() remove a linha identificada por “i” da Treeview.
        # Isso garante que a tabela seja completamente limpa antes de inserir novos dados.
        tree.delete(i)


    # ------------------- ETAPA 2: PREENCHER NOVAS LINHAS COM O EFEITO “ZEBRA” -------------------

    # O laço “for” com enumerate() percorre todas as linhas (rows) recebidas como parâmetro.
    # enumerate() retorna dois valores em cada iteração:
    #   - idx → índice da linha (0, 1, 2, 3...);
    #   - r → os valores reais da linha (ex.: id, nome, e-mail, telefone...).
    for idx, r in enumerate(rows):

        # Cria uma tag para a linha atual com base na sua posição (par ou ímpar).
        # Se o índice (idx) for par (ou seja, divisível por 2), a tag será “evenrow”;
        #                caso contrário (ímpar), será “oddrow”.
        # O operador módulo (%) é usado para verificar isso:
        #   - idx % 2 == 0 → linha par (even)
        #   - idx % 2 != 0 → linha ímpar (odd)
        tag = "evenrow" if idx % 2 == 0 else "oddrow"

        # Insere uma nova linha na Treeview.
        # Parâmetros:
        #   "" → significa que o item será inserido na raiz (não é um item filho).
        #   "end" → insere o item no final da lista existente.
        #   values=r → define os valores (colunas) que serão exibidos nessa linha.
        #   tags=(tag,) → aplica a tag de cor correspondente (oddrow ou evenrow).
        # Dessa forma, cada linha terá a cor correta alternada automaticamente.
        tree.insert("", "end", values=r, tags=(tag,))