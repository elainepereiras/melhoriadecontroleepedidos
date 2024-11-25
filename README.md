# Melhoria de Controle e Pedidos
O aplicativo foi projetado para automatizar e agilizar o processo de inserção de informações que serão posteriormente usadas para gerar relatórios no Power BI. Abaixo estão descritas as funcionalidades e etapas para o uso do aplicativo.

## 1 - Modos de inserção de Dados
Serão disponibilizados um arquivo compactado com dois modos para a inserção de dados. Os modos disponíveis são:
·Excel Macro-Enabled Workbook  - xlsm
·Visual Basic Script - Vbs. 

## 2 - Instruções para Extrair Arquivos e Criar Atalho do Aplicativo ACA.vbs
2.1 - Extrair Arquivos para o Disco C:
- Baixe o arquivo compactado disponibilizado.
- Localize o arquivo compactado, clicar com o botão direito sobre o arquivo.
- Selecionar a opção Extrair para Aca\.
- Copie e cole o arquivo descompactado no disco local C :\ .

2.2 - Criar Atalho para o Aplicativo ACA.vbs
- Navegue até o disco local C :\ e localize o arquivo aca.vbs.
- Clique com o botão direito no arquivo aca.vbs, selecione a opção Enviar para > área de trabalho (criar atalho).
  
## 3 - Iniciar o aplicativo
Com o atalho do aca.vbs criado, pode-se abrir o aplicativo sem precisar iniciar o Excel manualmente. Basta dar um duplo clique no atalho aca.vbs e o aplicativo será iniciado automaticamente. Isso abrirá o sistema e exibirá a tela de login

## 4 - Login no Sistema
  - Ao abrir o aplicativo, será exibida uma tela de login solicitando o nome de usuário e a senha. 
  - Exemplo de credenciais padrão:
·Usuário: admin
·Senha: admin
- Após inserir as credenciais corretas, clique no botão Login para acessar o sistema.
  
## 5 - Menu Principal
Depois de realizar o login, será direcionado para o menu principal do aplicativo. Este menu apresenta quatro opções principais:
·Cadastro
·Faturamento
·Utilitários
·Sair
- Para acessar qualquer uma das funcionalidades, basta clicar sobre a opção desejada para entrar na seção correspondente.

5.1 - Detalhamento das Funcionalidades do Menu: Cadastro

A seção de Cadastro permite a inserção e edição de dados essenciais para o funcionamento do sistema, como informações de clientes, produtos, vendedores, condições de pagamento e status de pedidos.

## 5.1.1 - Cadastro de Clientes

5.1.1.1 - Cadastrar um Novo Cliente
- Acesse Cadastro > Cliente.
- Clique no botão Novo, e o sistema automaticamente gerará um código exclusivo para o cliente que está sendo registrado.
- Preencha todas as informações solicitadas sobre o cliente, como nome, endereço, telefone, entre outros.
- Após inserir todas as informações, clique no botão Gravar para salvar o cadastro.
- Se durante o preenchimento você desejar cancelar a operação, clique no botão Limpar para apagar os dados inseridos até o momento.
- Para sair da tela de cadastro de cliente, clique em Fechar.

5.1.1.2 Pesquisar Cliente Cadastrado
- Acesse Cadastro > Cliente.
- Use o botão de pesquisa do campo razão social para localizar o cliente, digitando o código do cliente desejado. 
- O sistema exibirá as informações do cliente correspondente.

5.1.2 - Cadastro de Produtos

5.1.2.1 - Cadastrar Novo Produto
- Acesse Cadastro > Produto.
- Clique no botão Novo, e o sistema gerará um código exclusivo para o produto.
- Após inserir todas as informações, clique em Salvar para registrar o produto.
- Se desejar cancelar a operação, clique em Limpar para apagar os dados inseridos.
- Para sair da tela de cadastro, clique em Fechar.

5.1.2.2 - Alterar Produto
- Acesse Cadastro > Produto.
- Selecione o produto que deseja alterar na lista.
- Modifique as informações necessárias e clique em Alterar para salvar as mudanças.

5.1.2.3 Imprimir Lista de Produtos
 - Acesse Cadastro > Produto > Imprimir. O sistema gerará um arquivo PDF contendo as informações dos produtos cadastrados, pronto para impressão.

5.1.3 - Cadastro de Vendedores

5.1.3.1 - Cadastrar Novo Vendedor
- Acesse Cadastro > Vendedor.
- Clique no botão Novo para gerar um código exclusivo para o novo vendedor.
- Insira o nome do vendedor.
- Clique em Salvar para registrar o vendedor.
- Para sair da tela, clique em Fechar.

5.1.3.2 - Alterar Informações de Vendedores
- Acesse Cadastro > Vendedor.
- Selecione o vendedor que deseja alterar na lista.
- Modifique a informação necessárias e clique em Alterar para salvar as mudanças.

5.1.3.3 - Imprimir Lista dos Vendedores Cadastrados
-Acessar Cadastro > Produto > Imprimir. O sistema gerará um PDF com todas as informações preenchidas, pronto para impressão.

5.1.4 - Cadastro de Condições de Pagamento

5.1.4.1 - Cadastrar uma Nova Condição de Pagamento
- Acesse Cadastro > Cond. Pagamento.
- Clique em Novo para gerar um código exclusivo para a nova condição de pagamento.
- Preencha as informações necessárias e clique em Salvar para registrar a condição.
- Para sair da tela, clique em Fechar.

5.1.4.2 - Alterar Informações Sobre Condições de Pagamento
- Acesse Cadastro > Cond. Pagamento 
- Selecione o nome da condição de pagamento que deseja modificar, modifique as informações necessárias e clique no botão Alterar.
-Caso esteja alterando as informações do vendedor e desista de dar prosseguimento, clicar no botão Limpar, que as informações de alteração serão apagadas.

5.1.4.3 - Imprimir Lista de Condições de Pagamento
- Acesse Cadastro > Cond. Pagamento > Imprimir. O sistema gerará um arquivo PDF contendo das condições de pagamento cadastrados, pronto para impressão.

5.1.5- Cadastrar Status

5.1.5.1 – Cadastrar um Novo Status
- Acesse Cadastro > Status.
- Clique no botão Novo, o sistema gerará automaticamente um código para o novo status.
- Preencha o campo com o novo status e selecione Salvar.
- Se, durante o preenchimento, decidir não continuar, clique em Limpar para apagar as informações inseridas.
- Para sair da tela, clique em Fechar.

5.1.5.2 - Alterar Informações de um Status
- Acesse Cadastro > Status.
- Selecione o status que deseja modificar, altere as informações necessárias e clique em Alterar.
- Se, durante a alteração, optar por não continuar, clique em Limpar para desfazer as modificações.

5.1.5.3 - Imprimir a Lista de Status
- Acesse Cadastro > Status > Imprimir. O sistema gerará um arquivo PDF com a lista de status, pronto para impressão.
  
## 5.2 - Detalhamento das Funcionalidades do Menu: Faturamento

A aba de Faturamento permite o gerenciamento de pedidos e entregas, possibilitando a inserção de novas ordens e o acompanhamento de pedidos em andamento.

5.2.1 - Cadastrar Pedido
-Acesse Faturamento > Pedidos.
- Informe o número do pedido do cliente e o prazo de entrega.
- Selecione a condição de pagamento e o campo emitido por.
- No campo Razão Social, clique no botão de pesquisa e selecione a empresa correspondente na lista.
- Se o endereço de cobrança for o mesmo, marque a opção O mesmo, caso contrário, insira o endereço de cobrança manualmente.
-Para adicionar produtos, clique em Novo.
- Insira o código do produto e pressione Enter para que os campos relacionados sejam preenchidos automaticamente.
- Preencha a quantidade desejada no campo Qtd.(M) e pressione Enter.
- Clique em Salvar para salvar as informações inseridas.
- Para cancelar as informações, clique em Limpar.
- Se desejar incluir mais produtos, clique novamente em Novo e repita o processo de cadastro.
- Para modificar um item já inserido, dê um duplo clique no item desejado, faça as alterações necessárias, pressione Enter e clique em Alterar.
- Sobre as informações dos dados de entrega, preencha o nome da transportadora e o nome do motorista.
-Insira o CEP de entrega e clique em Consultar CEP para preencher automaticamente os campos restantes. Insira manualmente apenas o número do local.
-Após preencher todos os campos, clique em Gravar para salvar o pedido.
- Se precisar cancelar a operação, clique em Cancelar para apagar os dados inseridos.
- Para imprimir o pedido cadastrado, clique no botão Imprimir e o sistema gerará um PDF com todas as informações preenchidas, pronto para impressão.

5.2.2- Informar Entregas Parciais de Pedido
- Acesse Faturamento > Entregas Parciais.
- Dê duplo clique no pedido desejado.
-Preencha a quantidade desejada para liberação no campo Qtd.(L) Falt. e pressione Enter.
-Caso necessite modificar o status e/ou a data de entrega, esses campos estarão disponíveis para alteração. 
- Clique em Salvar para gravar as informações.
- Para alterar informações de algum pedido parcial, selecione com duplo clique no pedido desejado, modifique as informações, clicar em Enter e depois pressionar no botão Alterar.
- Caso tenha fornecido alguma informação e não quiser salvar ou modificar, pressione Limpar.
- Para imprimir o formulário, clique no botão Imprimir e o sistema gerará um PDF com todas as informações preenchidas, pronto para impressão.


## 5.3 - Detalhamento das Funcionalidades do Menu: Utilitários

5.3.1 – Solicitar ajuda remotamente
- Acesse Utilitários > Acesso Remoto. Com isso será gerado um número de identificação. Um dos membros da equipe utilizará esse número para acessar o computador remotamente e, se necessário, solucionará o erro identificado.

5.3.2 -Reiniciar o Aplicativo
- Acesse em Utilitários > Reiniciar, o sistema voltará automaticamente para a tela de login.

## 5.4 - Detalhamento da funcionalidade do Menu: Sair 
- Para encerrar o uso do aplicativo, clique em Sair no menu principal. Isso fechará o sistema de forma segura e encerrará a sessão do usuário.
