function enviarConfirmacoesWhatsApp() {
  const calendarios = ['primary', 'email', 'gmail'];

  // Solicita a data desejada no formato dd/MM/aaaa
  const dataDesejada = "23/04/2025"; // Dia que você deseja a agenda
  const partesData = dataDesejada.split("/");
  const hoje = new Date(partesData[2], partesData[1] - 1, partesData[0]);
  
  const dataAlvo = calcularProximoDiaUtil(hoje);

  const dataInicial = new Date(dataAlvo);
  dataInicial.setHours(0, 0, 0, 0);

  const dataFinal = new Date(dataAlvo);
  dataFinal.setHours(23, 59, 59, 999);

  let agendaPorPastor = {};

  calendarios.forEach(calendarioId => {
    const eventos = CalendarApp.getCalendarById(calendarioId).getEvents(dataInicial, dataFinal);

    eventos.forEach(evento => {
      const titulo = evento.getTitle();
      const tituloLower = titulo.toLowerCase();

      // Ignorar títulos que contenham palavras irrelevantes
      const ignorarSeContem = ["reunião", "gestores", "escritório", "rede", "café"];
      if (ignorarSeContem.some(palavra => tituloLower.includes(palavra))) return;

      const hora = Utilities.formatDate(evento.getStartTime(), Session.getScriptTimeZone(), "HH:mm");
      const data = Utilities.formatDate(evento.getStartTime(), Session.getScriptTimeZone(), "dd/MM/yyyy");

      // Processa o título a partir da lista (forçando o nome do pastor para o calendário do Filipe)
      let resultado;
      if (calendarioId === "filipe.otoni@gmail.com") {
        resultado = interpretarTituloComLista(titulo, "Filipe Otoni");
      } else {
        resultado = interpretarTituloComLista(titulo);
      }

      const nomePastor = resultado.nomePastor || "SemNomePastor";
      const nomePessoaOriginal = resultado.nomePessoa || "SemNomePessoa";
      const telefone = formatarTelefone(resultado.telefone || "");
      const tipo = resultado.tipo || "pastor";

      // Limpa o nome da pessoa, removendo inclusive trechos com telefone ou o próprio nome do pastor
      const nomePessoa = limparNomePessoa(nomePessoaOriginal, nomePastor, telefone);

      // Monta a mensagem sem incluir o nome do pastor nem o telefone
      const mensagem = `Olá, tudo bem? 😊\n\n` +
        `Aqui é do CN Office. Estamos passando para confirmar seu atendimento no dia ${data} às ${hora}.\n\n` +
        `Se precisar reagendar ou tiver qualquer dúvida, estamos à disposição. Será um prazer atender você! 💛`;

      const linkZap = `https://wa.me/${telefone.replace(/\D/g, '')}?text=${encodeURIComponent(mensagem)}`;
      Logger.log(`✅ ${nomePessoa} às ${hora}:\n${linkZap}\n`);

      const chave = `${tipo}|${nomePastor}`;
      if (!agendaPorPastor[chave]) agendaPorPastor[chave] = [];
      // Aqui incluímos o telefone na lista de atendimentos
      agendaPorPastor[chave].push(`🔖 ${hora} - ${nomePessoa} ${telefone}`);
    });
  });

  for (let chave in agendaPorPastor) {
    const [tipo, nomePastor] = chave.split("|");
    const saudacao = tipo === "pastora" ? "pastora" : "pastor";
    const agendaFormatada = agendaPorPastor[chave].join("\n");
    const cabecalho = `(${nomePastor})\nBoa tarde ${saudacao}, agenda de amanhã:\n`;
    Logger.log(`${cabecalho}${agendaFormatada}`);
  }
}

function calcularProximoDiaUtil(dataAtual) {
  const dia = dataAtual.getDay();
  const novaData = new Date(dataAtual);
  if (dia === 5) novaData.setDate(dataAtual.getDate() + 3);
  else if (dia === 6) novaData.setDate(dataAtual.getDate() + 2);
  else if (dia === 0) novaData.setDate(dataAtual.getDate() + 1);
  else novaData.setDate(dataAtual.getDate() + 1);
  return novaData;
}

function interpretarTituloComLista(titulo, nomePastorForcado = "") {
  const listaPastores = ["Paulo Mota", "César Junior", "Felipe Otoni", "Marinho", "Marcus", "Filipe Otoni"];
  const listaPastoras = ["Jéssica", "Kelly", "Sayne", "Dalva", "Cida", "Nilda"];
  const todosPastores = [...listaPastores, ...listaPastoras];

  const regexTelefone = /(\+55\s?\d{2}\s?\d{4,5}-?\d{4}|\b\d{10,11}\b)/;
  const telefoneMatch = titulo.match(regexTelefone);
  let telefone = telefoneMatch ? telefoneMatch[1].replace(/\D/g, "") : "";

  if (telefone.length === 11 && !telefone.startsWith("55")) {
    telefone = "55" + telefone;
  }

  if (nomePastorForcado) {
    const tipo = listaPastoras.includes(nomePastorForcado) ? "pastora" : "pastor";
    let nomePessoa = limparNomePessoa(titulo, nomePastorForcado, telefone);
    return { nomePastor: nomePastorForcado, nomePessoa, telefone, tipo };
  }

  let nomePastor = "";
  let tipo = "";
  const tituloNormalizado = normalizarTexto(titulo);

  for (const nome of todosPastores) {
    const nomeNormalizado = normalizarTexto(nome);
    if (tituloNormalizado.includes(nomeNormalizado)) {
      nomePastor = nome;
      tipo = listaPastoras.includes(nome) ? "pastora" : "pastor";
      break;
    }
  }

  let nomePessoa = limparNomePessoa(titulo, nomePastor, telefone);
  return { nomePastor, nomePessoa, telefone, tipo };
}

function limparNomePessoa(titulo, nomePastor, telefone = "") {
  // Remove prefixos, o nome do pastor e o telefone do título,
  // além de eliminar formatações indesejadas
  let nome = titulo
    .replace(/atende:?/i, "")
    .replace(/Pr\.?|PR\.?|Pra\.?|PRA\.?/gi, "")
    .replace(nomePastor, "")
    .replace(telefone, "")
    .replace(/\+55\s?\d{2}\s?\d{4,5}-?\d{4}/g, "")
    .replace(/[:\-~()]/g, "")
    .replace(/\s+/g, " ")
    .trim();
  return nome;
}

function formatarTelefone(numero) {
  if (!numero || numero.length < 12) return numero;
  const ddd = numero.substring(2, 4);
  const parte1 = numero.substring(4, 9);
  const parte2 = numero.substring(9);
  return `+55 ${ddd} ${parte1}-${parte2}`;
}

function normalizarTexto(texto) {
  return texto
    .toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/\s+/g, " ")
    .trim();
}
