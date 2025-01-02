// Reativação de Usuários no domínio

function reativarUsuarios() {
    try {
        const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
        const abaReativacao = spreadsheet.getSheetByName('Nome Aba/Planilha');

        if (!abaReativacao) {
            Logger.log('A aba "Nome Aba/Planilha" não foi encontrada.');
            return;
        }
        // Registra o email do executor da ação
        const emailCriador = Session.getActiveUser().getEmail();
        const dados = abaReativacao.getDataRange().getValues();

        for (let i = 1; i < dados.length; i++) {
            const email = dados[i][0];  // Email na primeira coluna
            const statusAtual = dados[i][2];  // Status na terceira coluna

            if (email && statusAtual !== "Reativado") {
                Logger.log(`Processando email: ${email}`);

                const result = reativarUsuarioEAdicionarGrupo(email);

                if (result.success) {
                    // Registra a data e a Hora
                    const dataAtual = new Date();
                    const dataFormatada = Utilities.formatDate(
                        dataAtual,
                        Session.getScriptTimeZone(),
                        "dd/MM/yyyy HH:mm:ss");
                    // Define as linhas e colunas que serão utilizadas na execução
                    abaReativacao.getRange(i + 1, 2).setValue("Reativado");
                    abaReativacao.getRange(i + 1, 5).setValue(result.message);
                    abaReativacao.getRange(i + 1, 3).setValue(dataFormatada);
                    abaReativacao.getRange(i + 1, 4).setValue(emailCriador);
                } else {
                    abaReativacao.getRange(i + 1, 2).setValue("Falha");
                    abaReativacao.getRange(i + 1, 5).setValue(result.message);
                }

                SpreadsheetApp.flush();

                // Pausa para evitar limite de requisições
                Utilities.sleep(500);
            }
        }

        // Mensagem final
        SpreadsheetApp.getUi().alert('Processo de reativação concluído!');
    } catch (error) {
        Logger.log('Erro crítico no processo de reativação: ' + error.toString());
        SpreadsheetApp.getUi().alert('Erro no processo de reativação: ' + error.toString());
    }
}

function reativarUsuarioEAdicionarGrupo(email) {
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
        // Tenta recuperar o usuário
        let user;
        try {
            user = AdminDirectory.Users.get(email);
        } catch (getUserError) {
            return {
                success: false,
                message: `Usuário não encontrado: ${email}. Erro: ${getUserError.message}`
            };
        }

        // Verifica se o usuário está suspenso
        if (!user.suspended) {
            return {
                success: false,
                message: "Usuário não está suspenso"
            };
        }

        // Reativa o usuário
        try {
            user.suspended = false;
            AdminDirectory.Users.update(user, email);
            Logger.log(`Usuário ${email} reativado com sucesso`);
        } catch (reactivationError) {
            return {
                success: false,
                message: `Erro ao reativar usuário: ${reactivationError.message}`
            };
        }

        // Adiciona ao grupo
        try {
            const gruporeativa = AdminDirectory.Groups.get('Inserir.o.dominio@dominio.com');

            AdminDirectory.Members.insert({
                email: email,
                role: 'MEMBER'
            }, gruporeativa.id);

            Logger.log(`Usuário ${email} adicionado ao grupo`);
        } catch (grupoError) {
            // Mesmo que falhe a adição ao grupo, considera a reativação um sucesso
            return {
                success: true,
                message: `Usuário reativado, mas erro ao adicionar ao grupo: ${grupoError.message}`
            };
        }

        return {
            success: true,
            message: "Usuário reativado e adicionado ao grupo"
        };

    } catch (error) {
        Logger.log(`Erro inesperado ao reativar ${email}: ${error.toString()}`);
        return {
            success: false,
            message: `Erro inesperado: ${error.message}`
        };
    }
}

function isValidEmail(email) {
    const emailRegex = /^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$/;
    return emailRegex.test(email);
}

/* 
Observações:
- A reativação é feita através do email fornecido;

- Para facilitar a execução é recomendado o uso da função: 
    function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Nome do Menu')
    .addItem('Nome do Item', 'reativarUsuarios')
    }
ao final do código;
- Caso tenha dado outro nome a função principal, ao adicionar ao menu na funcção "onOpen" deve inserir o nome da função no lugar de "reativarUsuarios"; 
- Caso já possua um Menu e deseje adicionar mais um item basta colar:
    .addItem('Nome do Item', 'reativarUsuarios')
abaixo do item anterior.
*/
