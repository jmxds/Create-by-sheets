// Código para suspensão de Usuários, usará o email como base da busca

function processarSuspensao() {
    try {
        const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
        // Deve inserir o nome da planilha a qual criou para usar esse processo de suspensão
        const abaSuspensao = spreadsheet.getSheetByName('Nome da aba/planilha');

        if (!abaSuspensao) {
            Logger.log('A aba "Nome da aba/planilha" não foi encontrada.');
            return;
        }

        // Salva o email do usuário que está executando o script
        const emailCriador = Session.getActiveUser().getEmail();
        const dados = abaSuspensao.getDataRange().getValues();

        for (let i = 1; i < dados.length; i++) {
            const email = dados[i][0];
            const statusAtual = dados[i][2];

            if (email && statusAtual !== "Suspenso") {
                const result = suspenderUsuarioERemoverAcessos(email);

                if (result.success) {
                    // Registra a data e a hora de quando a ação foi executada
                    const dataAtual = new Date();
                    const dataFormatada = Utilities.formatDate(
                        dataAtual,
                        Session.getScriptTimeZone(),
                        "dd/MM/yyyy HH:mm:ss"
                    );
                    // Define onde serão adicionadas as informações geradas pelo código
                    abaSuspensao.getRange(i + 1, 2).setValue("Suspenso");
                    abaSuspensao.getRange(i + 1, 5).setValue(result.message);
                    abaSuspensao.getRange(i + 1, 3).setValue(dataFormatada);
                    abaSuspensao.getRange(i + 1, 4).setValue(emailCriador);
                } else {
                    abaSuspensao.getRange(i + 1, 2).setValue("Falha");
                    abaSuspensao.getRange(i + 1, 5).setValue(result.message);
                }

                SpreadsheetApp.flush();
            }
        }
    } catch (error) {
        Logger.log('Erro crítico no processamento de suspensão: ' + error.toString());
    }
}

function suspenderUsuarioERemoverAcessos(email) {
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

        // Verifica se o usuário já está suspenso
        if (user.suspended) {
            return {
                success: false,
                message: "Usuário já está suspenso"
            };
        }

        // Variáveis para acompanhar remoções
        let gruposRemovidos = 0;
        let dispositivosRemovidos = 0;
        let aplicativosRemovidos = 0;

        // 1. Remove o usuário de todos os grupos
        try {
            const grupos = AdminDirectory.Groups.list({
                userKey: email
            }).groups;

            if (grupos) {
                grupos.forEach(grupo => {
                    try {
                        AdminDirectory.Members.remove(grupo.id, email);
                        gruposRemovidos++;
                    } catch (removeError) {
                        Logger.log(`Erro ao remover usuário do grupo ${grupo.email}: ${removeError}`);
                    }
                });
            }
        } catch (listError) {
            Logger.log(`Erro ao listar grupos do usuário: ${listError}`);
        }

        // 2. Remover dispositivos conectados
        try {
            // Dispositivos Chrome OS
            const dispositivos = AdminDirectory.Chromeosdevices.list({
                customerId: 'my_customer',
                query: `user:${email}`
            });

            if (dispositivos.chromeosdevices) {
                dispositivos.chromeosdevices.forEach(dispositivo => {
                    try {
                        AdminDirectory.Chromeosdevices.action({
                            customerId: 'my_customer',
                            resourceId: dispositivo.deviceId,
                            action: 'disable'
                        });
                        dispositivosRemovidos++;
                    } catch (disableError) {
                        Logger.log(`Erro ao desativar dispositivo ${dispositivo.deviceId}: ${disableError}`);
                    }
                });
            }

            // Dispositivos móveis
            const dispositivosMobile = AdminDirectory.Mobiledevices.list({
                customerId: 'my_customer',
                query: `email:${email}`
            });

            if (dispositivosMobile.mobiledevices) {
                dispositivosMobile.mobiledevices.forEach(dispositivo => {
                    try {
                        AdminDirectory.Mobiledevices.delete('my_customer', dispositivo.resourceId);
                        dispositivosRemovidos++;
                    } catch (deleteError) {
                        Logger.log(`Erro ao remover dispositivo móvel ${dispositivo.resourceId}: ${deleteError}`);
                    }
                });
            }
        } catch (deviceError) {
            Logger.log(`Erro ao listar dispositivos do usuário: ${deviceError}`);
        }

        // 3. Remover aplicativos conectados
        try {
            const tokens = AdminDirectory.Tokens.list(email);

            if (tokens.items && tokens.items.length > 0) {
                tokens.items.forEach(token => {
                    try {
                        AdminDirectory.Tokens.remove(email, token.clientId);
                        aplicativosRemovidos++;
                    } catch (removeError) {
                        Logger.log(`Erro ao revogar o token do aplicativo: ${removeError}`);
                    }
                });
            }
        } catch (tokenError) {
            Logger.log(`Erro ao listar ou remover tokens: ${tokenError}`);
        }

        // 4. Suspender o usuário
        try {
            user.suspended = true;
            AdminDirectory.Users.update(user, email);

            return {
                success: true,
                message: `Usuário suspenso. Removido de ${gruposRemovidos} grupos, ${dispositivosRemovidos} dispositivos desativados e ${aplicativosRemovidos} aplicativos desconectados.`
            };
        } catch (updateError) {
            return {
                success: false,
                message: `Erro ao suspender usuário: ${updateError.message}`
            };
        }

    } catch (error) {
        // Captura qualquer erro não tratado
        Logger.log(`Erro inesperado ao suspender ${email}: ${error.toString()}`);
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
- A busca é feita através do email, e não pelo nome do usuário
- Caso queria facilitar a execução do script adicionando um menu, insira a função abaixo do código:
    function onOpen() {
    var ui = SpreadsheetApp.getUi();
    ui.createMenu('Nome do que deseja dar ao menu')
        .addItem('Nome do item que deseja colocar', 'Nome dado a função principal')
        // caso não tenha alterado o nome da função basta adicionar:  'processarSuspensao'
}

- Caso já possua um menu e só queira adicionar mais um item insira : .addItem('Nome do item', 'Nome da função principal') abaixo do item anterior
    

*/
