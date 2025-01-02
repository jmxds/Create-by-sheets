// Faz uma busca no domínio para através dos nome e sobrenome localizar o email correspondente

function buscarEmailPorNome() {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const abaBusca = spreadsheet.getSheetByName('Nome da Aba/Planilha');

    if (!abaBusca) {
        Logger.log('A aba "Nome da Aba/Planilha" não foi encontrada.');
        return;
    }

    // Limpa resultados anteriores (a partir da segunda linha)
    const ultimaLinha = abaBusca.getLastRow();
    if (ultimaLinha > 1) {
        abaBusca.getRange(2, 2, ultimaLinha - 1, 6).clearContent();
    }

    const dados = abaBusca.getDataRange().getValues();
    // Faz a busca em multiplos domínios
    const dominios = ['Inserir.domínio.com.br', 'Inserir.domínio2.com.br', 'Inserir.domínio3.com.br', 'Inserir.domínio4.com.br'];

    for (let i = 1; i < dados.length; i++) {
        const nomeCompleto = dados[i][0];

        if (nomeCompleto) {
            Logger.log(`Buscando usuário: ${nomeCompleto}`);
            let usuariosEncontrados = [];

            // Tenta encontrar em cada domínio
            for (let dominio of dominios) {
                const usuarioNoDominio = buscarUsuarioNoDominio(nomeCompleto, dominio);

                if (usuarioNoDominio) {
                    usuariosEncontrados.push({
                        dominio: dominio,
                        usuario: usuarioNoDominio
                    });
                }
            }

            // Filtra e prioriza usuários
            const usuariosFiltrados = filtrarUsuarios(usuariosEncontrados, nomeCompleto);

            if (usuariosFiltrados.length > 0) {
                // Preenche os resultados na planilha
                const emails = usuariosFiltrados.map(item => item.usuario.primaryEmail).join(', ');

                abaBusca.getRange(i + 1, 2).setValue(emails);

                // Status (verifica se algum usuário está suspenso)
                const status = usuariosFiltrados.some(item => item.usuario.suspended) ? "Suspenso" : "Ativo";
                abaBusca.getRange(i + 1, 3).setValue(status);

                // Informações de último login
                const ultimosLogins = usuariosFiltrados.map(item => {
                    let ultimoLogin = item.usuario.lastLoginTime;
                    if (!ultimoLogin || new Date(ultimoLogin).getTime() === 0) {
                        return `${item.usuario.primaryEmail}: Nunca fez login`;
                    } else {
                        return `${item.usuario.primaryEmail}: Último login ${Utilities.formatDate(
                            new Date(ultimoLogin),
                            Session.getScriptTimeZone(),
                            "dd/MM/yyyy HH:mm:ss"
                        )}`;
                    }
                }).join(' | ');

                abaBusca.getRange(i + 1, 4).setValue(ultimosLogins);

                // Mensagem
                abaBusca.getRange(i + 1, 5).setValue("Encontrado");

                // Domínios
                const dominiosEncontrados = usuariosFiltrados.map(item => item.dominio).join(', ');
                abaBusca.getRange(i + 1, 6).setValue(dominiosEncontrados);

                // Log para diagnóstico
                Logger.log(`Usuários encontrados para ${nomeCompleto}:`);
                usuariosFiltrados.forEach(item => {
                    Logger.log(`- Email: ${item.usuario.primaryEmail}, Domínio: ${item.dominio}`);
                });
            } else {
                // Se não encontrou, preenche as colunas com "não encontrado"
                Logger.log(`Usuário não encontrado: ${nomeCompleto}`);
                abaBusca.getRange(i + 1, 2).setValue("Não encontrado");
                abaBusca.getRange(i + 1, 3).setValue("N/A");
                abaBusca.getRange(i + 1, 4).setValue("N/A");
                abaBusca.getRange(i + 1, 5).setValue("Usuário não existe nos domínios");
                abaBusca.getRange(i + 1, 6).setValue("N/A");
            }
        }
    }
}

/* 
    Dentro da função "filtrarUsuarios" existe a variavel "usuariosPrioridade" e um if 
    que usa ela como referência, essa variável serve para definir a prioridade de busca
    em um dos domínos. Caso não vá utilizar esse domínio, ou tenha somente um domínio
    pode deletar esse if e essa "const".
*/
function filtrarUsuarios(usuariosEncontrados, nomeCompleto) {
    // Primeiro, verifica se há usuários no domínio definido como prioritário com nome completo exatamente igual
    const usuariosPrioridade = usuariosEncontrados.filter(item =>
        item.dominio === 'inserir.dominio.com.br' &&
        item.usuario.name.fullName.toLowerCase().trim() === nomeCompleto.toLowerCase().trim()
    );

    // Se encontrou um email com nome completo exatamente igual, retorna apenas esses
    if (usuariosPrioridade.length > 0) {
        return usuariosPrioridade;
    }

    // Prioriza usuários com nome completo correspondente em outros domínios
    const usuariosNomeCompleto = usuariosEncontrados.filter(item =>
        item.usuario.name.fullName.toLowerCase().trim() === nomeCompleto.toLowerCase().trim()
    );

    if (usuariosNomeCompleto.length > 0) {
        return usuariosNomeCompleto;
    }

    // Se não encontrou nome completo, tenta primeiro nome e último nome
    const partesNome = nomeCompleto.split(' ');
    const primeiroNome = partesNome[0];
    const ultimoNome = partesNome[partesNome.length - 1];

    const usuariosNomePartes = usuariosEncontrados.filter(item =>
        item.usuario.name.givenName.toLowerCase().trim() === primeiroNome.toLowerCase().trim() &&
        item.usuario.name.familyName.toLowerCase().trim() === ultimoNome.toLowerCase().trim()
    );

    if (usuariosNomePartes.length > 0) {
        return usuariosNomePartes;
    }

    // Se não encontrou, retorna todos os usuários encontrados
    return usuariosEncontrados;
}

function buscarUsuarioNoDominio(nomeCompleto, dominio) {
    const estrategias = [
        // Busca por nome completo
        () => {
            try {
                Logger.log(`Tentando buscar por fullName="${nomeCompleto}" no domínio ${dominio}`);
                const resultado = AdminDirectory.Users.list({
                    domain: dominio,
                    query: `name:"${nomeCompleto}"`,
                    maxResults: 1
                });
                return resultado.users && resultado.users.length > 0 ? resultado.users[0] : null;
            } catch (error) {
                Logger.log(`Erro ao buscar por fullName no domínio ${dominio}: ${error}`);
                return null;
            }
        },

        // Busca por primeiro nome e último nome
        () => {
            const partesNome = nomeCompleto.split(' ');
            if (partesNome.length > 1) {
                const primeiroNome = partesNome[0];
                const ultimoNome = partesNome[partesNome.length - 1];

                try {
                    Logger.log(`Tentando buscar por givenName="${primeiroNome}" familyName="${ultimoNome}" no domínio ${dominio}`);
                    const resultado = AdminDirectory.Users.list({
                        domain: dominio,
                        query: `givenName="${primeiroNome}" familyName="${ultimoNome}"`,
                        maxResults: 1
                    });
                    return resultado.users && resultado.users.length > 0 ? resultado.users[0] : null;
                } catch (error) {
                    Logger.log(`Erro ao buscar por givenName e familyName no domínio ${dominio}: ${error}`);
                    return null;
                }
            }
            return null;
        },

        // Busca por nome sem acentos
        () => {
            const nomeSemAcentos = removerAcentos(nomeCompleto);

            try {
                Logger.log(`Tentando buscar por nome sem acentos: "${nomeSemAcentos}" no domínio ${dominio}`);
                const resultado = AdminDirectory.Users.list({
                    domain: dominio,
                    query: `name:"${nomeSemAcentos}"`,
                    maxResults: 1
                });
                return resultado.users && resultado.users.length > 0 ? resultado.users[0] : null;
            } catch (error) {
                Logger.log(`Erro ao buscar por nome sem acentos no domínio ${dominio}: ${error}`);
                return null;
            }
        }
    ];

    // Tenta cada estratégia
    for (let estrategia of estrategias) {
        const resultado = estrategia();
        if (resultado) {
            return resultado;
        }
    }

    return null;
}

function removerAcentos(texto) {
    return texto
        .normalize('NFD')
        .replace(/[\u0300-\u036f]/g, '')
        .replace(/[áÁàÀâÂãÃäÄ]/g, 'a')
        .replace(/[éÉèÈêÊëË]/g, 'e')
        .replace(/[íÍìÌîÎïÏ]/g, 'i')
        .replace(/[óÓòÒôÔõÕöÖ]/g, 'o')
        .replace(/[úÚùÙûÛüÜ]/g, 'u')
        .replace(/[çÇ]/g, 'c');
}

