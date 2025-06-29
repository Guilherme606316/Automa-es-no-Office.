//
// ==================== Funções Utilitárias ====================
//

/**
 * Formata uma data ISO (AAAA-MM-DD) para o padrão brasileiro (DD/MM/AAAA).
 * @param {string} dataISO
 * @returns {string}
 */
function formatarDataParaBrasileiro(dataISO) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(dataISO)) return dataISO;
  const [ano, mes, dia] = dataISO.split("-");
  return `${dia}/${mes}/${ano}`;
}

/**
 * Exibe uma mensagem de status para o usuário na interface.
 * O texto fica verde para sucesso e vermelho para erros.
 * @param {string} mensagem
 * @param {boolean} erro
 */
function exibirStatus(mensagem, erro = false) {
  const status = document.getElementById("status");
  status.textContent = mensagem;
  status.style.color = erro ? "red" : "green";
}

/**
 * Restaura a interface para o estado inicial: limpa o formulário
 * e mostra novamente os botões principais ("Cadastrar" e "Atualizar").
 */
function restaurarInterface() {
  const formulario = document.getElementById("formularioInterativo");
  formulario.innerHTML = "";
  document.getElementById("botaoCadastrar").style.display = "";
  document.getElementById("botaoAtualizar").style.display = "";
}

/**
 * Permite navegação entre os elementos de input usando as setas do teclado.
 * Facilita para usuários que preferem teclado ao invés do mouse.
 * @param {HTMLElement[]} elementos
 */
function habilitarNavegacaoSetas(elementos) {
  elementos.forEach((el, idx) => {
    el.addEventListener("keydown", (e) => {
      let alvo;
      if (e.key === "ArrowDown") alvo = elementos[idx + 1];
      if (e.key === "ArrowUp") alvo = elementos[idx - 1];
      if (alvo) {
        e.preventDefault();
        alvo.focus();
      }
    });
  });
}

/**
 * Exibe o formulário para cadastrar uma nova rotina no Word.
 * Permite inserir todos os dados necessários e insere esses dados no documento Word ao confirmar.
 */
function cadastrarRotina() {
  document.getElementById("botaoAtualizar").style.display = "none";
  const formulario = document.getElementById("formularioInterativo");
  formulario.innerHTML = `
    <div class="form-group">
      <label for="Numero de três digítos">Número de 3 dígitos:</label>
      <input type="text" id="Numero de três digítos" maxlength="3"/>
      <button id="botaoSemNumero3" type="button">Sem número de 3 dígitos</button>
    </div>
    <div class="form-group">
      <label for="local">Local:</label>
      <input type="text" id="local"/>
      <button id="botaoSemLocal" type="button">Sem local</button>
    </div>
    <div class="form-group">
      <label for="vencimento">Vencimento:</label>
      <input type="date" id="vencimento"/>
      <button id="botaoSemVencimento" type="button">Sem vencimento</button>
    </div>
    <div class="form-group">
      <label for="status">Status:</label>
      <select id="status">
        <option>A EXECUTAR</option>
        <option>EXECUTADA</option>
        <option>EM EXECUÇÃO</option>
        <option>Sem status</option>
      </select>
    </div>
    <button id="botaoEnviarCadastro" type="button">Enviar</button>
  `;
  // Botões de recusar caso você saiba somente alguns valores e não todos.
  document.getElementById("botaoSemNumero3").onclick = () =>
    (document.getElementById("Numero de três digítos").value = "Sem número de 3 dígitos");
  document.getElementById("botaoSemLocal").onclick = () => (document.getElementById("local").value = "Sem local");
  document.getElementById("botaoSemVencimento").onclick = () =>
    (document.getElementById("vencimento").value = "Sem vencimento");

  // Permite navegação por setas entre campos e botões do formulário
  habilitarNavegacaoSetas([
    document.getElementById("Numero de três digítos"),
    document.getElementById("botaoSemNumero3"),
    document.getElementById("local"),
    document.getElementById("botaoSemLocal"),
    document.getElementById("vencimento"),
    document.getElementById("botaoSemVencimento"),
    document.getElementById("status"),
    document.getElementById("botaoEnviarCadastro")
  ]);

  // Quando o botão "Enviar" for clicado, os valores são validados e inseridos no Word.
  document.getElementById("botaoEnviarCadastro").onclick = async () => {
    const os = document.getElementById("Numero de três digítos").value.trim();
    const local = document.getElementById("local").value.trim();
    const vencimento = document.getElementById("vencimento").value.trim();
    const status = document.getElementById("status").value.trim();
    const vencimentoFormatado = vencimento ? formatarDataParaBrasileiro(vencimento) : "Sem vencimento";
    await Word.run(async (contexto) => {
      const corpo = contexto.document.body;
      corpo.insertParagraph(`#${os} – Rotina`, Word.InsertLocation.end).font.bold = true;
      corpo.insertParagraph("Local: ", Word.InsertLocation.end).font.bold = true;
      corpo.insertText(`${local}`, Word.InsertLocation.end).font.bold = false;
      corpo.insertParagraph("Vencimento: ", Word.InsertLocation.end).font.bold = true;
      corpo.insertText(`${vencimentoFormatado}`, Word.InsertLocation.end).font.bold = false;
      corpo.insertParagraph("Status: ", Word.InsertLocation.end).font.bold = false;
      corpo.insertText(`${status}`, Word.InsertLocation.end).font.bold = true;
      corpo.insertParagraph('', Word.InsertLocation.end);
      await contexto.sync();
    }).catch(console.error);

    exibirStatus("Cadastrado com sucesso!");
    restaurarInterface();
  };
}

/**
 * Atualiza campos pendentes no documento Word, substituindo marcadores (placeholders) como "Sem número..." pelos valores informados pelo usuário.
 */
async function atualizarPendencias() {
  document.getElementById("botaoCadastrar").style.display = "none";
  await Word.run(async (contexto) => {
    const corpo = contexto.document.body;
    const marcadores = [
      "Sem número de 3 dígitos",
      "Sem local",
      "Sem vencimento",
      "Sem status"
    ];
    for (const marcador of marcadores) {
      await Word.run(async (ctx) => {
        const resultados = ctx.document.body.search(marcador, { matchCase: false });
        resultados.load("items");
        await ctx.sync();
        if (resultados.items.length === 0) return;

        for (const item of resultados.items) {
          const intervalos = item.getTextRanges(["\n"], false);
          intervalos.load("items");
          await ctx.sync();
          const texto = intervalos.items[0]?.text.trim() || "";
          let html = "";

          // Gera o formulário correto para cada marcador pendente
          if (marcador === "Sem vencimento") {
            html = `
              <div class="form-group">
                <label>Preencha o campo de vencimento ou pule:</label>
                <input type="date" id="campoAtualizar" />
                <button id="botaoOk"   type="button">OK</button>
                <button id="botaoPular" type="button">Pular</button>
              </div>
            `;
          } else if (marcador === "Sem status") {
            html = `
              <div class="form-group">
                <label for="status">Preencha o campo de status ou pule:</label>
                <select id="campoAtualizar">
                  <option>A EXECUTAR</option>
                  <option>EXECUTADA</option>
                  <option>EM EXECUÇÃO</option>
                  <option>Sem status</option>
                </select>
                <button id="botaoOk"   type="button">OK</button>
                <button id="botaoPular" type="button">Pular</button>
              </div>
            `;
          } else if (marcador === "Sem número de 3 dígitos") {
            html = `
              <div class="form-group">
                <label>Coloque um número de 3 dígitos ou pule:</label>
                <input type="text" id="campoAtualizar" maxlength="3">
                <button id="botaoOk"   type="button">OK</button>
                <button id="botaoPular" type="button">Pular</button>
              </div>
            `;
          } else {
            html = `
              <div class="form-group">
                <label>Substituir "Sem local" ou pular:</label>
                <input type="text" id="campoAtualizar" />
                <button id="botaoOk"   type="button">OK</button>
                <button id="botaoPular" type="button">Pular</button>
              </div>
            `;
          }

          const formulario = document.getElementById("formularioInterativo");
          formulario.innerHTML = html;

          const botaoOk = document.getElementById("botaoOk");
          const botaoPular = document.getElementById("botaoPular");
          const campo = document.getElementById("campoAtualizar");
          campo.focus();

          // Permite navegar entre os campos/botões do pequeno formulário de atualização
          habilitarNavegacaoSetas([campo, botaoOk, botaoPular]);

          // Espera usuário clicar OK (com valor) ou Pular (mantém o marcador)
          const resposta = await new Promise((res) => {
            botaoOk.onclick = () => res(campo.value.trim());
            botaoPular.onclick = () => res(null);
          });

          if (resposta !== null && resposta !== "") {
            let valorParaInserir = resposta;
            if (marcador === "Sem vencimento") {
              valorParaInserir = formatarDataParaBrasileiro(resposta);
              item.insertText(valorParaInserir, Word.InsertLocation.replace).font.bold = false;
            }
            else if (marcador === 'Sem número de 3 dígitos') {
              item.insertText(valorParaInserir, Word.InsertLocation.replace).font.bold = true;
            }
            else if ( marcador === 'Sem local') {
              item.insertText(valorParaInserir, Word.InsertLocation.replace).font.bold = false;
            }
            else {
              item.insertText(valorParaInserir, Word.InsertLocation.replace).font.bold = true;
            }
            corpo.insertParagraph('', Word.InsertLocation.end)
            await ctx.sync();
          }
        }
      }).catch(console.error);
    }
    exibirStatus("Todas as pendências foram atualizadas.");
    restaurarInterface();
  });
}

/**
 * Mostra o formulário para inserir o nome do plantonista e insere o cabeçalho padrão,
 * caso ainda não existam no documento.
 * @param {Word.RequestContext} contexto
 */
async function inserirCabecalhoSeNecessario(contexto) {
  const corpo = contexto.document.body;
  const hoje = new Date().toLocaleDateString("pt-BR");
  const diaSemana = new Date().toLocaleDateString("pt-BR", { weekday: "long" });
  const diaSemanaFormatado = diaSemana.charAt(0).toUpperCase() + diaSemana.slice(1);
  const textoCabecalho = "Status de Rotina Diária";
  const textoCabecalhoCompleto = `${textoCabecalho} - ${diaSemanaFormatado} - ${hoje}`;
  const buscaCabecalho = textoCabecalho;
  const buscaNome = "NOME: ";
  const resultadosCabecalho = corpo.search(buscaCabecalho, { matchCase: false });
  const resultadosPlantonista = corpo.search(buscaNome, { matchCase: false });
  resultadosCabecalho.load("items");
  resultadosPlantonista.load("items");
  await contexto.sync();

  // Se não houver cabeçalho nem plantonista, pede o nome do plantonista e insere tudo no início
  if (resultadosCabecalho.items.length === 0 && resultadosPlantonista.items.length === 0) {
    const formulario = document.getElementById("formularioInterativo");
    // Aqui você pode tirar ou adicionar nomes ao seu relatorio.
    formulario.innerHTML = `
      <div class="form-group">
        <label for="campo_nome">Nome do autor:</label>
        <datalist id="lista_nomes">
          <option>NICOLAS</option>
          <option>CAMILE</option>
          <option>JOÃO</option>
          <option>MARCIO</option>
        </datalist>
        <input type="text" id="campo_nome" list="lista_nomes" />
        <button id="botaoEnviarNome" type="button">Enviar</button>
      <div>
    `;
    habilitarNavegacaoSetas([document.getElementById("campo_nome"), document.getElementById("botaoEnviarNome")]);

    document.getElementById("botaoEnviarNome").onclick = async () => {
      const nome = document.getElementById("campo_nome").value.trim();
      await Word.run(async (ctx) => {
        const hoje = new Date().toLocaleDateString("pt-BR");
        const diaSemanaBruto = new Date().toLocaleDateString("pt-BR", { weekday: "long" });
        const diaSemanaFormatado = diaSemanaBruto.charAt(0).toUpperCase() + diaSemanaBruto.slice(1);
        const textoNegrito = "Status do relatório";
        const textoNormal = ` - ${diaSemanaFormatado} - ${hoje}`;
        const textoNome = `NOME: ${nome}`;
        const paragrafoNome = ctx.document.body.insertParagraph(textoNome, Word.InsertLocation.start);
        paragrafoNome.font.bold = true;
        const paragrafoData = ctx.document.body.insertParagraph(textoNormal, Word.InsertLocation.start);
        paragrafoData.font.bold = false;
        const paragrafoCabecalho = ctx.document.body.insertText(textoNegrito, Word.InsertLocation.start);
        paragrafoCabecalho.font.bold = true;
        ctx.document.body.insertText("", Word.InsertLocation.end).font.bold = false;
        await ctx.sync();
      });
      document.getElementById("formularioInterativo").innerHTML = "";
    };
    await contexto.sync();
  }
}

//
// ==================== Código principal: Ponto de entrada ====================
//

Office.onReady(async (info) => {
  if (info.host === Office.HostType.Word) {
    // Insere o cabeçalho padrão no início, se necessário.
    await Word.run(async (contexto) => {
      await inserirCabecalhoSeNecessario(contexto);
    }).catch(console.error);

    // Liga os botões principais às suas funções.
    document.getElementById("botaoCadastrar").onclick = cadastrarRotina;
    document.getElementById("botaoAtualizar").onclick = atualizarPendencias;

    // Permite navegar entre os botões principais usando as setas do teclado.
    const botoesPrincipais = [document.getElementById("botaoCadastrar"), document.getElementById("botaoAtualizar")];
    habilitarNavegacaoSetas(botoesPrincipais);
  }
});
