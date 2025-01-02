// Código para criação de emails, ele vai gerar o email através do primeiro nome e do sobrenome

function processarCriacaoUsuarios() {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const abaCriacao = spreadsheet.getSheetByName('Nome da aba sheets');

    if (!abaCriacao) {
        Logger.log('A aba "Nome da aba sheets" não foi encontrada.');
        return;
    }

    // Grupo que todos os usuários serão incluidos (Opcional)
    const GRUPO_TODOS_ID = 'Domínio do Grupo em que os usuários possa ser adicionados';

    // Salva o email de quem está executando o script
    const emailCriador = Session.getActiveUser().getEmail();

    const dados = abaCriacao.getDataRange().getValues();

    for (let i = 1; i < dados.length; i++) {
        const nomeCompleto = dados[i][0];

        if (nomeCompleto && dados[i][4] !== "Criado") {
            const nomePartes = nomeCompleto.trim().split(' ');
            let email = '';
            let mensagemEmail = '';
            let emailCriado = false;
            let ultimoNomePadrao = nomePartes[nomePartes.length - 1];
            let sobrenomeUsado = '';

            // Tenta criar email com diferentes combinações de sobrenomes
            for (let j = nomePartes.length - 1; j > 0; j--) {
                const primeiroNome = nomePartes[0];
                const ultimoNome = nomePartes[j];

                email = `${primeiroNome}.${ultimoNome}@inserir.seu.dominio.com`.toLowerCase();

                // Verifica se o email já existe no domínio
                if (!verificarEmailExistente(email)) {
                    emailCriado = true;
                    sobrenomeUsado = ultimoNome;

                    // Adiciona mensagem caso o email tenha sido criado com um sobrenome diferente do ultimo
                    if (ultimoNome !== ultimoNomePadrao) {
                        mensagemEmail = `Email criado com sobrenome diferente (${sobrenomeUsado})`;
                    }
                    break;
                }
            }

            // Se não conseguir criar email com nenhuma combinação
            if (!emailCriado) {
                abaCriacao.getRange(i + 1, 5).setValue("Falha");
                abaCriacao.getRange(i + 1, 8).setValue("Já existem emails com todos os sobrenomes informados");
                continue;
            }

            // Verifica se o email gerado é válido
            if (!isValidEmail(email)) {
                abaCriacao.getRange(i + 1, 5).setValue("Falha");
                abaCriacao.getRange(i + 1, 8).setValue("Email inválido");
                continue;
            }

            const senhaPadrao = "Inserir a senha que será a padrão em todos os emails criados";

            const message = criarUsuarioEAdicionarGrupo(nomeCompleto, email, senhaPadrao, GRUPO_TODOS_ID);

            if (message) {
                abaCriacao.getRange(i + 1, 5).setValue("Falha");
                abaCriacao.getRange(i + 1, 8).setValue(message);
            } else {
                // Vai informar a data e hora da criação do email
                const dataAtual = new Date();
                const dataFormatada = Utilities.formatDate(
                    dataAtual,
                    Session.getScriptTimeZone(),
                    "dd/MM/yyyy HH:mm:ss"
                );

                // O as informações que serão adicionadas em cada coluna
                abaCriacao.getRange(i + 1, 2).setValue(nomePartes[0]);
                abaCriacao.getRange(i + 1, 3).setValue(nomePartes[nomePartes.length - 1]);
                abaCriacao.getRange(i + 1, 4).setValue(email);
                abaCriacao.getRange(i + 1, 5).setValue("Criado");
                abaCriacao.getRange(i + 1, 6).setValue(dataFormatada);
                abaCriacao.getRange(i + 1, 7).setValue(emailCriador);

                // Coluna da mensagem
                if (mensagemEmail) {
                    abaCriacao.getRange(i + 1, 8).setValue(mensagemEmail);
                }
            }

            SpreadsheetApp.flush();
        }
    }
}

function criarUsuarioEAdicionarGrupo(nomeCompleto, email, senhaPadrao, grupoId) {
    try {
        // Criar usuário
        const user = {
            primaryEmail: email,
            name: {
                fullName: nomeCompleto,
                givenName: nomeCompleto.split(' ')[0],
                familyName: nomeCompleto.split(' ').slice(1).join(' ')
            },
            password: senhaPadrao,
            // Neste exemplo será pedido, ao usuário, quando acessar o email que seja escolhida uma nova senha
            changePasswordAtNextLogin: true
        };

        // Insere o usuário
        AdminDirectory.Users.insert(user);

        // Adiciona o usuário ao grupo
        try {
            AdminDirectory.Members.insert(
                {
                    email: email,
                    role: 'MEMBER'
                },
                grupoId
            );
        } catch (grupoError) {
            Logger.log('Erro ao adicionar usuário ao grupo: ' + grupoError);
            // Não interrompe o processo, apenas registra o erro
        }

        return null;
    } catch (e) {
        Logger.log(e);
        return e.message;
    }
}

function verificarEmailExistente(email) {
    try {
        try {
            AdminDirectory.Users.get(email);
            return true; // Email já existe
        } catch (error) {
            // Se der erro de não encontrado, significa que o email não existe
            return false;
        }
    } catch (error) {
        // Erro inesperado
        Logger.log(`Erro ao verificar email ${email}: ${error}`);
        return true; // Considera como existente em caso de erro
    }
}
// Imperde o email de ser criado caso contenha esses caracteres especiais
function isValidEmail(email) {
    const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
    return emailRegex.test(email);
}

/* 
Importante lembrar:
- Deve ser adicionada as APIs AdSense Managenment API e Admin SDK API

- Para facilitar a execução do Script, deve ser adicionada a função a baixo ao final do código:
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Nome que será exibido nas ferramentas do sheets ')
// Caso não tenha alterado o nome da função, o nome é: processarCriacaoUsuarios
    .addItem('Nome do item que vai aparecer no menu criado a cima', 'Nome da função')
}
*/
