// Redefine a senha dos emails para a senha padrão

function processarRedefinicaoSenha() {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const abaSenha = spreadsheet.getSheetByName('Nome da aba/planilha');

    if (!abaSenha) {
        Logger.log('A aba "Nome da aba/planilha" não foi encontrada.');
        return;
    }

    const emailCriador = Session.getActiveUser().getEmail();
    const dados = abaSenha.getDataRange().getValues();

    for (let i = 1; i < dados.length; i++) {
        // Armazena o email da primeira coluna
        const email = dados[i][0];

        // Verifica se há uma senha personalizada na segunda coluna
        const senhaPersonalizada = dados[i][1];

        // Define a nova senha (personalizada ou fixa)
        const novaSenha = senhaPersonalizada || "Inserir_Senha_Padrão";

        // Verifica se o email não está vazio e ainda não foi processado
        if (email && dados[i][2] !== "Redefinida") {
            // Redefine a senha
            const result = redefinirSenha(email, novaSenha);

            // Atualiza a planilha
            if (!result.success) {
                abaSenha.getRange(i + 1, 3).setValue("Falha");
                abaSenha.getRange(i + 1, 6).setValue(result.message);
            } else {
                // Registra a data e a hora atual
                const dataAtual = new Date();
                const dataFormatada = Utilities.formatDate(
                    dataAtual,
                    Session.getScriptTimeZone(),
                    "dd/MM/yyyy HH:mm:ss");
                    // Define as colunas onde serão armazenadas as informações
                abaSenha.getRange(i + 1, 3).setValue("Redefinida");
                abaSenha.getRange(i + 1, 6).setValue("Senha redefinida com sucesso");
                abaSenha.getRange(i + 1, 4).setValue(dataFormatada);
                abaSenha.getRange(i + 1, 5).setValue(emailCriador);

                // Se não foi digitada uma senha personalizada, preenche com a senha fixa
                if (!senhaPersonalizada) {
                    abaSenha.getRange(i + 1, 2).setValue("Inserir_Senha_Padrão");
                }
            }

            SpreadsheetApp.flush();
        }
    }
}

function redefinirSenha(email, novaSenha) {
    // Validações iniciais
    if (!email) {
        return {
            success: false,
            message: "Email não fornecido"
        };
    }

    // Validação de email
    if (!isValidEmail(email)) {
        return {
            success: false,
            message: "Email inválido: " + email
        };
    }

    try {
        // Tenta recuperar o usuário primeiro
        let user;
        try {
            user = AdminDirectory.Users.get(email);
        } catch (getUserError) {
            return {
                success: false,
                message: `Usuário não encontrado: ${email}. Erro: ${getUserError.message}`
            };
        }

        // Tenta redefinir a senha
        try {
            user.password = novaSenha;
            user.changePasswordAtNextLogin = true;

            AdminDirectory.Users.update(user, email);

            return {
                success: true,
                message: "Senha redefinida com sucesso"
            };
        } catch (updateError) {
            return {
                success: false,
                message: `Erro ao redefinir senha: ${updateError.message}`
            };
        }

    } catch (error) {
        Logger.log(`Erro inesperado ao redefinir senha para ${email}: ${error.toString()}`);
        return {
            success: false,
            message: `Erro inesperado: ${error.message}`
        };
    }
}

// Função de validação de email
function isValidEmail(email) {
    const emailRegex = /^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$/;
    return emailRegex.test(email);
}
/* 
Observações: 
- A senha é redefinida através do email fornecido;

- Para facilitar a execução é recomendado o uso da função: 
    function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Nome do Menu')
    .addItem('Nome do Item', 'processarRedefinicaoSenha')
    }
ao final do código;
- Caso tenha dado outro nome a função principal, ao adicionar ao menu na funcção "onOpen" deve inserir o nome da função no lugar de "processarRedefinicaoSenha"; 
- Caso já possua um Menu e deseje adicionar mais um item basta colar:
    .addItem('Nome do Item', 'processarRedefinicaoSenha')
abaixo do item anterior.
*/ 
