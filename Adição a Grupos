// Adiciona usuários a Grupos pelo email

function adicionarUsuariosAoGrupo() {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

    // Nome da planilha de controle
    const abaCadastro = spreadsheet.getSheetByName('Nome da Aba/Planilha');

    if (!abaCadastro) {
        Logger.log('A aba "Nome da Aba/Planilha" não foi encontrada.');
        return;
    }




    // Define o grupo ao qual será adicionado
    const GRUPO_TODOS_ID = 'dominio.do.grupo@domínio.com.br';

    const dados = abaCadastro.getDataRange().getValues();

    // Pula o cabeçalho
    for (let i = 1; i < dados.length; i++) {
        const email = dados[i][0]; // Email na primeira coluna
        const status = dados[i][1]; // Status na segunda coluna

        // Verifica se o email não está vazio e ainda não foi processado
        if (email && status !== "Adicionado") {
            const resultado = adicionarAoGrupo(email, GRUPO_TODOS_ID);

            if (resultado === null) {
                // Sucesso na adição ao grupo
                abaCadastro.getRange(i + 1, 2).setValue("Adicionado");
                abaCadastro.getRange(i + 1, 3).setValue("Usuário adicionado com sucesso");
            } else {
                // Falha na adição ao grupo
                abaCadastro.getRange(i + 1, 2).setValue("Falha");
                abaCadastro.getRange(i + 1, 3).setValue(resultado);
            }

            SpreadsheetApp.flush();
        }
    }
}

function adicionarAoGrupo(email, grupoId) {
    try {
        // Validação básica de email
        if (!isValidEmail(email)) {
            return "Email inválido";
        }

        // Adiciona o usuário ao grupo
        AdminDirectory.Members.insert(
            {
                email: email,
                role: 'MEMBER'
            },
            grupoId
        );
        return null; // Sucesso

    } catch (error) {
        Logger.log('Erro ao adicionar usuário ao grupo: ' + error);
        return error.message; // Retorna a mensagem de erro
    }
}

function isValidEmail(email) {
    const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
    return emailRegex.test(email);
}
/* 
Observações:
- A adição é feita através do email fornecido;

- Para facilitar a execução é recomendado o uso da função: 
    function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Nome do Menu')
    .addItem('Nome do Item', 'adicionarUsuariosAoGrupo')
    }
ao final do código;
- Caso tenha dado outro nome a função principal, ao adicionar ao menu na funcção "onOpen" deve inserir o nome da função no lugar de "adicionarUsuariosAoGrupo"; 
- Caso já possua um Menu e deseje adicionar mais um item basta colar:
    .addItem('Nome do Item', 'adicionarUsuariosAoGrupo')
abaixo do item anterior.
*/
