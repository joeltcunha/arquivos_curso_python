Arquivo: conexao.php
1. Colocar as mensagens em português

inserir.php
1. Colocar mensagens em português
2. Porque você precisa retornar o id gerado? linha 14
3. Você colocou uma mensagem na tela "New record created successfully". Na fase de teste pode ser util, mas após os testes você pretende utiizá-la? Linha 13
4. Gramática. Iniciar frases com maiúscula.
5. Padronizar as mensagens.

Cadastrar_livro.html
1. Porque $idEditora recebe NULL? linha 17
2. Depois me explique o uso desses códigos

cadastro_livro.html
1. Retirar o <style> e colocar num arquivo separado. Até porque você colocou um link para um css
2. Diminuir o border-radius. Está parecendo bullets
3. Não está responsivo
4. Indentação

cadastro_eletronicos.php
1. Idem ao anterior
2. Trocar "Itens que acompanham" por "Itens inclusos"

cadastro_usuario.html
1. Diminuir o tamanho da fonte dos textos de responsabilidade

devolucao.html
1. idem ao cadastro_livro.html
2. O campo Nº se refere a que?
3. Os campos de data estão pequenos e não cabem a data completa
4. Alterar para Data de empréstimo e Data de devolução
5. Data de devolução deve ser um campo do tipo "data atual". Não pode ser selecionável
6. Ao clicar em devolução de livro no menu lateral ele leva para uma página chamada devolucao_livro.html e não devolucao.html

emprestimo_livro.html
1. Idem aos itens 1, 2 e 3 do cadastro_livro.html
2. Alterar o label dos campos de data para "Data empréstimo" e Data máxima para devolução"
3. O campo data deve ser do tipo data atual. Não pode ser selecionável

emprestimo_eletronicos.html
1. Idem aos itens 1, 2 e 3 do cadastro_livro.html
2. Porque tem dois alunos?
3. Horário de entrega e de devolução devem ser do tipo data e hora atual
4. Alterar horário de entrega para "Horário de retirada"