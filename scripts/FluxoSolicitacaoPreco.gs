/**
 * Função principal que busca por solicitações pendentes e as processa.
 * AJUSTADO: Inclui verificação da coluna "Comprovante de Verba".
 */
function processarSolicitacoesPendentes() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) {
    console.log("Não foi possível obter o bloqueio. Outra execução já está em andamento.");
    return;
  }

  try {
    const PLANILHA_APPSHEET_ID = "1GDjpgOJsxzKwmu5spOQK0szRoPWUXYhBj2dBg6toQi8";
    const NOME_ABA = "Solicitações do App";
    
    // Colunas de controle e dados fundamentais
    const COLUNA_PROCESSADO = "Processado";
    const COLUNA_STATUS = "Status do Processamento";
    const COLUNA_EMAIL_FORM = "Endereço de e-mail";
    const COLUNA_ID_FORM = "ID";
    const COLUNA_EMAIL_COPIA_FORM = "E-mail Cópia"; 
    const COLUNA_COMPROVANTE_FORM = "Comprovante de Verba"; // <<== ADICIONADO: Nome da coluna no AppSheet

    const aba = SpreadsheetApp.openById(PLANILHA_APPSHEET_ID).getSheetByName(NOME_ABA);
    const dados = aba.getDataRange().getValues();
    const cabecalhos = dados[0];
    
    const idxProcessado = cabecalhos.indexOf(COLUNA_PROCESSADO);
    const idxStatus = cabecalhos.indexOf(COLUNA_STATUS);
    const idxEmail = cabecalhos.indexOf(COLUNA_EMAIL_FORM);
    const idxId = cabecalhos.indexOf(COLUNA_ID_FORM);
    const idxEmailCopia = cabecalhos.indexOf(COLUNA_EMAIL_COPIA_FORM);
    const idxComprovante = cabecalhos.indexOf(COLUNA_COMPROVANTE_FORM); // <<== ADICIONADO: Pega o índice

    // Verificação se as colunas existem (Adicionado idxComprovante na validação opcional, se for obrigatório)
    if (idxProcessado === -1 || idxStatus === -1 || idxEmail === -1 || idxId === -1 || idxEmailCopia === -1) {
      throw new Error(`Uma ou mais colunas necessárias não foram encontradas na aba "${NOME_ABA}".`);
    }
    
    // Aviso caso a coluna nova não exista (opcional, mas bom para debug)
    if (idxComprovante === -1) {
       console.warn(`A coluna "${COLUNA_COMPROVANTE_FORM}" não foi encontrada na aba de origem. O campo ficará vazio.`);
    }

    for (let i = 1; i < dados.length; i++) {
      const linha = dados[i];
      const processado = linha[idxProcessado];
      const id = linha[idxId]; 

      if (!processado && id) {
        const statusRange = aba.getRange(i + 1, idxStatus + 1);
        const processadoRange = aba.getRange(i + 1, idxProcessado + 1);

        statusRange.setValue("Em processamento...");
        SpreadsheetApp.flush();

        const namedValues = {};
        cabecalhos.forEach((coluna, j) => {
            let nomeColuna = coluna === COLUNA_EMAIL_FORM ? "E-mail do Solicitante" : coluna;
            namedValues[nomeColuna] = [linha[j]];
        });

        const eventoSimulado = { namedValues };
        const statusMessage = processarSolicitacao(eventoSimulado);

        statusRange.setValue(statusMessage);
        processadoRange.setValue(new Date()).setNumberFormat("dd/MM/yyyy HH:mm:ss");

        if (statusMessage.toUpperCase().includes("ERRO") || statusMessage.toUpperCase().includes("AVISO")) {
          const emailSolicitante = linha[idxEmail];
          const idSolicitacao = linha[idxId];

          if (emailSolicitante && emailSolicitante.includes("@")) {
            try {
              const assunto = `Alerta no Processamento da Solicitação ID: ${idSolicitacao}`;
              const corpoEmail = `
                <p>Olá,</p>
                <p>Houve um problema durante o processamento da sua solicitação (ID: <b>${idSolicitacao}</b>).</p>
                <p>Por favor, verifique a mensagem de status abaixo, corrija o que for necessário (geralmente no arquivo anexo) e envie uma nova solicitação.</p>
                <br>
                <hr>
                <p><b>Detalhes do Status:</b></p>
                <p style="font-family: monospace; background-color: #f4f4f4; padding: 10px; border-radius: 5px; white-space: pre-wrap;">${statusMessage}</p>
                <hr>
                <br>
                <p>Atenciosamente,<br>Pricing</p>
              `;

              MailApp.sendEmail({
                to: emailSolicitante,
                subject: assunto,
                htmlBody: corpoEmail
              });
              console.log(`E-mail de alerta enviado para ${emailSolicitante} para a solicitação ID ${idSolicitacao}.`);

            } catch (emailError) {
              console.error(`Falha ao enviar e-mail de erro para ${emailSolicitante}. Erro: ${emailError.toString()}`);
              statusRange.setValue(statusMessage + " | AVISO: Falha ao enviar a notificação por e-mail.");
            }
          } else {
            console.warn(`Não foi possível enviar e-mail de alerta para a solicitação ID ${idSolicitacao} por falta de um endereço de e-mail válido.`);
          }
        }
      }
    }
  } catch (e) {
    console.error("Ocorreu um erro em processarSolicitacoesPendentes: " + e.toString());
  } finally {
    lock.releaseLock();
  }
}


/**
 * Processa uma única solicitação, validando e filtrando os dados do arquivo anexo.
 * AJUSTADO: Adicionado mapeamento para "Comprovante de Verba".
 */
function processarSolicitacao(e) {
  const PLANILHA_DESTINO_ID = "1GDjpgOJsxzKwmu5spOQK0szRoPWUXYhBj2dBg6toQi8";
  const PLANILHA_LOG_ID = "1yQnoDsS9ARhFyvT_sboVPgEtKDO5CM0sQc_XoAosJC0";
  const PASTA_ANEXOS_ID = "1fyg2UP7vOdm1VXDuAr3TQH1I3Lq4_5M4";

  // Definição das colunas de destino (Nome Exato no Cabeçalho da Aba Destino)
  const COLUNA_EMAIL = "E-mail do Solicitante";
  const COLUNA_OBSERVACAO = "Observação";
  const COLUNA_BANDEIRA = "Bandeira";
  const COLUNA_DATA = "Data da Solicitação";
  const COLUNA_HORA = "Hora";
  const COLUNA_OK_OPERACOES = "Tem ok de operações?";
  const COLUNA_ID = "ID";
  const COLUNA_ANEXO = "Arquivo Padrão de Solicitação";
  const COLUNA_EMAIL_COPIA = "E-mail Cópia";
  const COLUNA_COMPROVANTE = "Comprovante de Verba"; // <<== ADICIONADO

  const executionLog = {
    errors: [], warnings: [], info: [],
    startTime: new Date(), idSolicitacao: "N/A",
    formData: e ? JSON.stringify(e.namedValues) : 'Nenhum dado recebido'
  };

  let planilhaIdConvertida;

  try {
    if (!e || !e.namedValues) throw new Error("Dados do formulário inválidos ou não recebidos");

    const tiposSelecionados = e.namedValues["Qual tipo de solicitação você deseja registrar?"][0].split(",").map(s => s.trim()).filter(Boolean);
    const bandeirasFiltro = [...new Set((e.namedValues[COLUNA_BANDEIRA]?.[0] || "").split(",").map(s => s.trim().toLowerCase()).filter(Boolean))];
    const abasSemBandeira = ["Leve e Pague C&V", "Leve e Pague LB"];

    if (bandeirasFiltro.length === 0 && !tiposSelecionados.every(tipo => abasSemBandeira.includes(tipo))) {
      throw new Error("Nenhuma bandeira foi selecionada para solicitações que exigem essa informação.");
    }
    
    // Leitura dos valores do formulário
    const caminhoArquivo = e.namedValues[COLUNA_ANEXO]?.[0] || "";
    const emailSolicitante = e.namedValues[COLUNA_EMAIL]?.[0] || "Não informado";
    const observacao = e.namedValues[COLUNA_OBSERVACAO]?.[0] || "";
    const okOperacoes = e.namedValues[COLUNA_OK_OPERACOES]?.[0] || "";
    const idSolicitacao = e.namedValues[COLUNA_ID]?.[0] || "Sem ID";
    const emailCopia = e.namedValues[COLUNA_EMAIL_COPIA]?.[0] || "";
    const comprovanteVerba = e.namedValues[COLUNA_COMPROVANTE]?.[0] || ""; // <<== ADICIONADO: Lê o valor do AppSheet

    executionLog.idSolicitacao = idSolicitacao;
    executionLog.info.push(`Iniciando ID: ${idSolicitacao}. Filtros: Bandeiras [${bandeirasFiltro.join(", ")}]`);

    if (!caminhoArquivo) throw new Error("Nenhum arquivo anexado para processamento");

    const nomeArquivo = caminhoArquivo.replace(/^Solicitações do App_Files_\//, "");
    const arquivos = DriveApp.getFolderById(PASTA_ANEXOS_ID).getFilesByName(nomeArquivo);
    if (!arquivos.hasNext()) throw new Error(`Arquivo não encontrado: ${nomeArquivo}`);

    const arquivoOriginal = arquivos.next();
    planilhaIdConvertida = converterParaPlanilhaGoogle(arquivoOriginal.getId());
    const planilhaTemp = SpreadsheetApp.openById(planilhaIdConvertida);
    executionLog.info.push(`Arquivo aberto/convertido com sucesso.`);

    for (const tipoSelecionado of tiposSelecionados) {
      try {
        executionLog.info.push(`Processando tipo: ${tipoSelecionado}`);

        const abaOrigem = encontrarAbaPorNome(planilhaTemp, tipoSelecionado);
        const dadosOriginais = abaOrigem.getDataRange().getValues();
        const cabecalhoOrigem = dadosOriginais[0].map(h => String(h).trim());

        let dadosFiltrados = dadosOriginais.slice(1).filter(linha => linha.some(c => String(c).trim() !== ""));

        if (!abasSemBandeira.includes(tipoSelecionado)) {
          const cabecalhoOrigemLower = cabecalhoOrigem.map(h => h.toLowerCase());
          const idxBandeiraOrigem = cabecalhoOrigemLower.indexOf(COLUNA_BANDEIRA.toLowerCase());

          if (idxBandeiraOrigem === -1) {
            throw new Error(`A aba "${tipoSelecionado}" do anexo não tem a coluna "${COLUNA_BANDEIRA}".`);
          }
          
          dadosFiltrados = dadosFiltrados.filter(linha => {
            const bandeiraNaLinha = String(linha[idxBandeiraOrigem]).trim().toLowerCase();
            return bandeirasFiltro.includes(bandeiraNaLinha);
          });
        }
        
        if (dadosFiltrados.length === 0) {
          executionLog.warnings.push(`Nenhum dado válido foi encontrado na aba "${tipoSelecionado}" com os filtros aplicados`);
          continue;
        }

        executionLog.info.push(`Encontrados ${dadosFiltrados.length} registros válidos em "${tipoSelecionado}".`);
        
        const planilhaDestino = SpreadsheetApp.openById(PLANILHA_DESTINO_ID);
        let abaDestino = planilhaDestino.getSheetByName(tipoSelecionado) || planilhaDestino.insertSheet(tipoSelecionado);
        let cabecalhoDestino = abaDestino.getLastRow() === 0 ? [] : abaDestino.getRange(1, 1, 1, abaDestino.getLastColumn()).getValues()[0];

        // Se a aba destino estiver vazia, cria o cabeçalho
        if (abaDestino.getLastRow() === 0) {
          cabecalhoDestino = [...cabecalhoOrigem];
          
          // Lista de colunas extras vindas do Formulário
          const colunasDoFormulario = [
              COLUNA_EMAIL, 
              COLUNA_OBSERVACAO, 
              COLUNA_DATA, 
              COLUNA_HORA, 
              COLUNA_OK_OPERACOES, 
              COLUNA_ID, 
              COLUNA_EMAIL_COPIA, 
              COLUNA_COMPROVANTE // <<== ADICIONADO: Inclui na criação do cabeçalho se a aba for nova
          ];

          colunasDoFormulario.forEach(col => {
            if (!cabecalhoDestino.map(h => h.toLowerCase()).includes(col.toLowerCase())) {
              cabecalhoDestino.push(col);
            }
          });
          abaDestino.appendRow(cabecalhoDestino);
        }

        const agora = new Date();
        const cabecalhoOrigemLower = cabecalhoOrigem.map(h => h.toLowerCase());

        const linhasParaInserir = dadosFiltrados.map(linha => {
          return cabecalhoDestino.map(col => {
            const colNormalizada = String(col).trim().toLowerCase();
            const indiceOrigem = cabecalhoOrigemLower.indexOf(colNormalizada);

            // Mapeamento dos campos do Formulário/AppSheet para a Coluna Destino
            if (col === COLUNA_EMAIL) return emailSolicitante;
            if (col === COLUNA_OBSERVACAO) return observacao;
            if (col === COLUNA_DATA) return agora;
            if (col === COLUNA_HORA) return agora;
            if (col === COLUNA_OK_OPERACOES) return okOperacoes;
            if (col === COLUNA_ID) return idSolicitacao;
            if (col === COLUNA_EMAIL_COPIA) return emailCopia;
            if (col === COLUNA_COMPROVANTE) return comprovanteVerba; // <<== ADICIONADO: Retorna o valor lido do App
            
            // Se não for campo do formulário, tenta pegar do arquivo anexo (Excel)
            return indiceOrigem >= 0 ? linha[indiceOrigem] : "";
          });
        });

        if (linhasParaInserir.length > 0) {
          abaDestino.getRange(abaDestino.getLastRow() + 1, 1, linhasParaInserir.length, cabecalhoDestino.length).setValues(linhasParaInserir);
        }

        executionLog.info.push(`✅ ${linhasParaInserir.length} registros transferidos para ${tipoSelecionado}`);

      } catch (erroTipo) {
        executionLog.errors.push(`Erro ao processar "${tipoSelecionado}": ${erroTipo.message}`);
      }
    }

    const tiposComSucesso = tiposSelecionados.filter(tipo => !executionLog.errors.some(erro => erro.includes(tipo)));
    const statusFinal = [];
    if (tiposComSucesso.length > 0) statusFinal.push(`Processado: ${tiposComSucesso.join(", ")}`);
    if (executionLog.warnings.length > 0) statusFinal.push(`AVISO: ${executionLog.warnings.join(" | ")}`);
    if (executionLog.errors.length > 0) statusFinal.push(`ERRO: ${executionLog.errors.join(" | ")}`);

    gravarLogExecucao(executionLog, executionLog.errors.length > 0, PLANILHA_LOG_ID);
    return statusFinal.length > 0 ? statusFinal.join(" | ") : "Concluído sem dados para processar.";

  } catch (erro) {
    executionLog.errors.push(erro.message);
    gravarLogExecucao(executionLog, true, PLANILHA_LOG_ID);
    return `ERRO: ${executionLog.errors.join(" | ")}`;
  } finally {
    if (planilhaIdConvertida) {
      try {
        DriveApp.getFileById(planilhaIdConvertida).setTrashed(true);
        executionLog.info.push("Arquivo temporário limpo.");
      } catch (e) {
        executionLog.warnings.push("Falha ao limpar arquivo temporário: " + e.message);
        gravarLogExecucao(executionLog, false, PLANILHA_LOG_ID);
      }
    }
  }
}

// ... Restante das funções auxiliares (gravarLogExecucao, converterParaPlanilhaGoogle, etc.) permanecem iguais ...

function gravarLogExecucao(logData, isError, logSheetId) {
  try {
    const logSheet = SpreadsheetApp.openById(logSheetId).getSheetByName("Logs") ||
                     SpreadsheetApp.openById(logSheetId).insertSheet("Logs");

    if (logSheet.getLastRow() === 0) {
      logSheet.appendRow(["Data/Hora", "ID da Solicitação", "Tipo", "Email Solicitante", "Erros", "Warnings", "Informações", "Dados do Formulário", "Tempo de Execução"]);
    }

    const endTime = new Date();
    const executionTime = (endTime - logData.startTime) / 1000;
    const emailSolicitante = logData.formData.includes("E-mail do Solicitante") ? JSON.parse(logData.formData)["E-mail do Solicitante"][0] : "Não identificado";

    logSheet.appendRow([
      Utilities.formatDate(logData.startTime, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss"),
      logData.idSolicitacao || "N/A",
      isError ? "ERRO" : "INFO",
      emailSolicitante,
      logData.errors.join(" | ") || "Nenhum",
      logData.warnings.join(" | ") || "Nenhum",
      logData.info.join(" | ") || "Nenhum",
      logData.formData,
      `${executionTime} segundos`
    ]);
  } catch (logError) {
    console.error("Falha ao gravar log: ", logError);
  }
}

function converterParaPlanilhaGoogle(arquivoId) {
  const arquivo = DriveApp.getFileById(arquivoId);
  if (arquivo.getMimeType() === MimeType.GOOGLE_SHEETS) return arquivoId;
  const arquivoConvertido = Drive.Files.copy({}, arquivoId, {convert: true});
  return arquivoConvertido.id;
}

function encontrarAbaPorNome(planilha, nomeAba) {
  const nomeBusca = nomeAba.trim().toLowerCase();
  const abas = planilha.getSheets();
  for (const aba of abas) {
    if (aba.getName().trim().toLowerCase() === nomeBusca) return aba;
  }
  throw new Error(`Aba "${nomeAba}" não encontrada no arquivo.`);
}
