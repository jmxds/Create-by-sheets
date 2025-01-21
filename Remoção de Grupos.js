function processarRemocaoGrupos() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const abaRemocao = spreadsheet.getSheetByName('Nome_planilha');
  
  if (!abaRemocao) {
    Logger.log('A aba "Nome_panilha" não foi encontrada.');
    return;
  }

  try {
    // Email do usuário que está executando o script
    const emailCriador = Session.getActiveUser().getEmail();
    const dados = abaRemocao.getDataRange().getValues();
    
    for (let i = 1; i < dados.length; i++) {
      const email = dados[i][0];
      const statusAtual = dados[i][1];
      
      if (email && statusAtual !== "Removido") {
        const result = removerUsuarioDeGrupos(email);
        
        if (result.success) {
          const dataAtual = new Date();
          const dataFormatada = Utilities.formatDate(
            dataAtual, 
            Session.getScriptTimeZone(), 
            "dd/MM/yyyy HH:mm:ss"
          );
          // Linha e coluna onde entrarão os dados
          abaRemocao.getRange(i + 1, 2).setValue("Removido");
          abaRemocao.getRange(i + 1, 3).setValue(result.message);
          abaRemocao.getRange(i + 1, 4).setValue(dataFormatada);
          abaRemocao.getRange(i + 1, 5).setValue(emailCriador);
        } else {
          abaRemocao.getRange(i + 1, 2).setValue("Falha");
          abaRemocao.getRange(i + 1, 3).setValue(result.message);
        }
        
        SpreadsheetApp.flush();
      }
    }
  } catch (error) {
    Logger.log('Erro crítico no processamento de remoção de grupos: ' + error.toString());
  }
}

function removerUsuarioDeGrupos(email) {
  try {
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

    // Verifica se o usuário existe
    let user;
    try {
      user = AdminDirectory.Users.get(email);
    } catch (getUserError) {
      return { 
        success: false, 
        message: `Usuário não encontrado: ${email}. Erro: ${getUserError.message}` 
      };
    }

    // Remove o usuário de todos os grupos
    let gruposRemovidos = 0;
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
      return { 
        success: false, 
        message: `Erro ao listar grupos: ${listError.message}` 
      };
    }

    return { 
      success: true, 
      message: `Usuário removido de ${gruposRemovidos} grupos` 
    };

  } catch (error) {
    Logger.log(`Erro inesperado ao remover usuário dos grupos: ${error.toString()}`);
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
