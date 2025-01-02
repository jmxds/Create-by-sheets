function onOpen() {
  var ui = SpreadsheetApp.getUi();
// Inserir o nome que deseja utilizar no Menu
  ui.createMenu('Nome Exemplo de Menu')

// Inserir o nome do item, o qual aparecerá no Menu e o nome da função escolhido. Como no Exemplo abaixo
    .addItem('Exemplo Suspensão', 'nome_da_funcao')
    .addItem('Exemplo Redefinir', 'nome_da_funcao2')
    .addItem('Exemplo Criação', 'nome_da_funcao3')
    .addItem('Exemplo Busca', 'nome_da_funcao4')
    .addToUi();
}
