// ============================================================
// GESTÃO ESCOLAR - CODE.GS
// ============================================================

const SHEET_ID = "1-ZgtXS5Kz48NCMc-gcm7fcpnT2rsExNBi15eY0yvO-s";

const SESSION_SECONDS = 21600;


// ============================================================
// PLANILHA
// ============================================================

function getPlanilha() {
  return SpreadsheetApp.openById(SHEET_ID);
}


// ============================================================
// DO GET
// ============================================================

function doGet() {

  prepararBanco();

  return HtmlService
    .createHtmlOutputFromFile("index")
    .setTitle("Gestão Escolar")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}


// ============================================================
// BANCO
// ============================================================

function prepararBanco() {

  const ss = getPlanilha();

  criarAba(ss, "GE_USUARIOS", [
    "ID",
    "TIPO",
    "RA",
    "NOME",
    "USUARIO",
    "SENHA",
    "SERIE_ANO",
    "MATERIA"
  ]);

  criarAba(ss, "GE_ALUNOS", [
    "RA",
    "NOME",
    "SERIE_ANO"
  ]);

  criarAba(ss, "GE_PROFESSORES", [
    "RA",
    "NOME",
    "MATERIA",
    "SERIE_ANO"
  ]);

  criarAba(ss, "GE_MATERIAS", [
    "ID",
    "NOME",
    "PROFESSOR_RA",
    "PROFESSOR",
    "SERIE_ANO",
    "CONTEUDO",
    "LINKS"
  ]);

  criarAba(ss, "GE_PROVAS", [
    "ID",
    "DATA",
    "TITULO",
    "NOTA",
    "RA_ALUNO",
    "ALUNO",
    "MATERIA",
    "PROFESSOR_RA",
    "PROFESSOR",
    "SERIE_ANO",
    "DESCRICAO"
  ]);

  criarAba(ss, "GE_TRABALHOS", [
    "ID",
    "DATA",
    "TITULO",
    "NOTA",
    "RA_ALUNO",
    "ALUNO",
    "MATERIA",
    "PROFESSOR_RA",
    "PROFESSOR",
    "SERIE_ANO",
    "DESCRICAO"
  ]);

  criarAba(ss, "GE_FEEDBACKS", [
    "ID",
    "DATA",
    "TIPO",
    "ITEM_ID",
    "RA_ALUNO",
    "ALUNO",
    "PROFESSOR_RA",
    "PROFESSOR",
    "SERIE_ANO",
    "FEEDBACK"
  ]);

  return "Banco preparado!";
}


// ============================================================
// CRIAR ABA
// ============================================================

function criarAba(ss, nome, cabecalhos) {

  let aba = ss.getSheetByName(nome);

  if (!aba) {
    aba = ss.insertSheet(nome);
  }

  if (aba.getLastRow() === 0) {

    aba
      .getRange(1, 1, 1, cabecalhos.length)
      .setValues([cabecalhos]);

    aba
      .getRange(1, 1, 1, cabecalhos.length)
      .setFontWeight("bold");
  }
}


// ============================================================
// UTILIDADES
// ============================================================

function gerarID() {
  return Utilities.getUuid();
}


function texto(v) {
  return String(v == null ? "" : v).trim();
}


function normalizar(v) {
  return texto(v).toLowerCase();
}


// ============================================================
// CADASTRO
// ============================================================

function cadastrarUsuario(dados) {

  prepararBanco();

  const tipo = texto(dados.tipo).toUpperCase();
  const ra = texto(dados.ra);
  const nome = texto(dados.nome);
  const usuario = normalizar(dados.usuario);
  const senha = texto(dados.senha);
  const serie = texto(dados.serieAno);
  const materia = texto(dados.materia);

  if (!ra)
    throw new Error("Informe o RA/ID.");

  if (!nome)
    throw new Error("Informe o nome.");

  if (!usuario)
    throw new Error("Informe o usuário.");

  if (!senha || senha.length < 6)
    throw new Error("A senha precisa ter pelo menos 6 caracteres.");

  if (!serie)
    throw new Error("Selecione a Série/Ano.");

  if (tipo !== "ALUNO" && tipo !== "PROFESSOR")
    throw new Error("Tipo de conta inválido.");

  if (tipo === "PROFESSOR" && !materia)
    throw new Error("Informe a matéria do professor.");

  const ss = getPlanilha();

  const usuarios = ss.getSheetByName("GE_USUARIOS");
  const alunos = ss.getSheetByName("GE_ALUNOS");
  const professores = ss.getSheetByName("GE_PROFESSORES");

  const registros = usuarios.getDataRange().getValues();

  for (let i = 1; i < registros.length; i++) {

    if (texto(registros[i][2]) === ra) {
      throw new Error("Este RA/ID já está cadastrado.");
    }

    if (normalizar(registros[i][4]) === usuario) {
      throw new Error("Este usuário já está cadastrado.");
    }
  }

  const id = gerarID();

  usuarios.appendRow([
    id,
    tipo,
    ra,
    nome,
    usuario,
    senha,
    serie,
    materia
  ]);

  if (tipo === "ALUNO") {

    alunos.appendRow([
      ra,
      nome,
      serie
    ]);

  } else {

    professores.appendRow([
      ra,
      nome,
      materia,
      serie
    ]);
  }

  return {
    sucesso: true,
    mensagem: "Conta criada com sucesso!"
  };
}


// ============================================================
// LOGIN
// ============================================================

function login(dados) {

  prepararBanco();

  const usuario = normalizar(dados.usuario);
  const senha = texto(dados.senha);
  const tipo = texto(dados.tipo).toUpperCase();

  if (!usuario || !senha) {
    throw new Error("Informe usuário e senha.");
  }

  const aba =
    getPlanilha().getSheetByName("GE_USUARIOS");

  const registros =
    aba.getDataRange().getValues();

  for (let i = 1; i < registros.length; i++) {

    const usuarioPlanilha =
      normalizar(registros[i][4]);

    const senhaPlanilha =
      texto(registros[i][5]);

    const tipoPlanilha =
      texto(registros[i][1]).toUpperCase();

    if (
      usuarioPlanilha === usuario &&
      senhaPlanilha === senha &&
      tipoPlanilha === tipo
    ) {

      const usuarioLogado = {

        id: registros[i][0],

        tipo: tipoPlanilha,

        ra: texto(registros[i][2]),

        nome: texto(registros[i][3]),

        usuario: texto(registros[i][4]),

        serieAno: texto(registros[i][6]),

        materia: texto(registros[i][7])
      };

      const token =
        criarSessao(usuarioLogado);

      return {
        sucesso: true,
        token: token,
        usuario: usuarioLogado
      };
    }
  }

  return {
    sucesso: false,
    mensagem: "Usuário, senha ou tipo de conta incorreto."
  };
}


// ============================================================
// SESSÃO
// ============================================================

function criarSessao(usuario) {

  const token = gerarID();

  CacheService
    .getScriptCache()
    .put(
      "GE_SESSION_" + token,
      JSON.stringify(usuario),
      SESSION_SECONDS
    );

  return token;
}


function obterSessao(token) {

  if (!token) {
    throw new Error("Sessão não encontrada.");
  }

  const data =
    CacheService
      .getScriptCache()
      .get("GE_SESSION_" + token);

  if (!data) {
    throw new Error("Sua sessão expirou. Faça login novamente.");
  }

  return JSON.parse(data);
}


function obterUsuarioLogado(token) {
  return obterSessao(token);
}


function logout(token) {

  if (token) {

    CacheService
      .getScriptCache()
      .remove(
        "GE_SESSION_" + token
      );
  }

  return true;
}


// ============================================================
// ALUNOS
// ============================================================

function encontrarAluno(ra) {

  const aba =
    getPlanilha().getSheetByName("GE_ALUNOS");

  const dados =
    aba.getDataRange().getValues();

  const buscado = texto(ra);

  for (let i = 1; i < dados.length; i++) {

    if (
      texto(dados[i][0]) === buscado
    ) {

      return {

        ra: texto(dados[i][0]),

        nome: texto(dados[i][1]),

        serieAno: texto(dados[i][2])
      };
    }
  }

  return null;
}


function listarAlunos(token) {

  const usuario =
    obterSessao(token);

  if (usuario.tipo !== "PROFESSOR") {
    throw new Error(
      "Acesso permitido somente para professores."
    );
  }

  const dados =
    getPlanilha()
      .getSheetByName("GE_ALUNOS")
      .getDataRange()
      .getDisplayValues();

  return dados
    .slice(1)
    .filter(function(linha) {

      return (
        texto(linha[2]) ===
        texto(usuario.serieAno)
      );

    })
    .map(function(linha) {

      return {

        ra: linha[0],

        nome: linha[1],

        serieAno: linha[2]
      };

    });
}


// ============================================================
// MATÉRIAS
// ============================================================

function cadastrarMateria(token, dados) {

  const usuario =
    obterSessao(token);

  if (usuario.tipo !== "PROFESSOR") {
    throw new Error(
      "Somente professores podem cadastrar matérias."
    );
  }

  const nome =
    texto(dados.nome);

  const conteudo =
    texto(dados.conteudo);

  const links =
    texto(dados.links);

  if (!nome) {
    throw new Error(
      "Informe o nome da matéria."
    );
  }

  if (!usuario.serieAno) {
    throw new Error(
      "O professor não possui Série/Ano."
    );
  }

  getPlanilha()
    .getSheetByName("GE_MATERIAS")
    .appendRow([

      gerarID(),

      nome,

      usuario.ra,

      usuario.nome,

      usuario.serieAno,

      conteudo,

      links

    ]);

  return {

    sucesso: true,

    mensagem:
      "Matéria publicada com sucesso!"
  };
}


function listarMaterias(token) {

  const usuario =
    obterSessao(token);

  const dados =
    getPlanilha()
      .getSheetByName("GE_MATERIAS")
      .getDataRange()
      .getDisplayValues();

  return dados
    .slice(1)
    .map(function(linha, index) {

      return {

        linha: index + 2,

        id: linha[0],

        nome: linha[1],

        professorRA: linha[2],

        professor: linha[3],

        serieAno: linha[4],

        conteudo: linha[5],

        links: linha[6]
      };

    })
    .filter(function(item) {

      if (usuario.tipo === "ALUNO") {

        return (
          texto(item.serieAno) ===
          texto(usuario.serieAno)
        );
      }

      return (
        texto(item.professorRA) ===
        texto(usuario.ra)
      );
    });
}


function atualizarMateria(token, dados) {

  const usuario =
    obterSessao(token);

  if (usuario.tipo !== "PROFESSOR") {
    throw new Error(
      "Somente professores podem editar."
    );
  }

  const linha =
    Number(dados.linha);

  const aba =
    getPlanilha()
      .getSheetByName("GE_MATERIAS");

  const registros =
    aba.getDataRange().getValues();

  if (
    linha < 2 ||
    linha > registros.length
  ) {
    throw new Error(
      "Matéria não encontrada."
    );
  }

  if (
    texto(registros[linha - 1][2]) !==
    texto(usuario.ra)
  ) {
    throw new Error(
      "Você só pode editar suas próprias matérias."
    );
  }

  aba.getRange(linha, 2)
    .setValue(texto(dados.nome));

  aba.getRange(linha, 5)
    .setValue(usuario.serieAno);

  aba.getRange(linha, 6)
    .setValue(texto(dados.conteudo));

  aba.getRange(linha, 7)
    .setValue(texto(dados.links));

  return {

    sucesso: true,

    mensagem:
      "Matéria atualizada!"
  };
}


function excluirMateria(token, linha) {

  const usuario =
    obterSessao(token);

  if (usuario.tipo !== "PROFESSOR") {
    throw new Error(
      "Somente professores podem excluir."
    );
  }

  const numero =
    Number(linha);

  const aba =
    getPlanilha()
      .getSheetByName("GE_MATERIAS");

  const dados =
    aba.getDataRange().getValues();

  if (
    numero < 2 ||
    numero > dados.length
  ) {
    throw new Error(
      "Matéria inválida."
    );
  }

  if (
    texto(dados[numero - 1][2]) !==
    texto(usuario.ra)
  ) {
    throw new Error(
      "Você só pode excluir suas próprias matérias."
    );
  }

  aba.deleteRow(numero);

  return {

    sucesso: true,

    mensagem:
      "Matéria excluída!"
  };
}


// ============================================================
// PROVAS
// ============================================================

function cadastrarProva(token, dados) {

  const usuario =
    obterSessao(token);

  if (usuario.tipo !== "PROFESSOR") {
    throw new Error(
      "Somente professores podem cadastrar provas."
    );
  }

  const aluno =
    encontrarAluno(dados.alunoRA);

  if (!aluno) {
    throw new Error(
      "RA do aluno não encontrado."
    );
  }

  if (
    texto(aluno.serieAno) !==
    texto(usuario.serieAno)
  ) {
    throw new Error(
      "O aluno pertence a outra Série/Ano."
    );
  }

  if (!dados.data) {
    throw new Error(
      "Informe a data."
    );
  }

  if (!dados.titulo) {
    throw new Error(
      "Informe o título."
    );
  }

  if (!dados.materia) {
    throw new Error(
      "Informe a matéria."
    );
  }

  let nota = texto(dados.nota);

  if (nota !== "") {

    const numeroNota =
      Number(nota);

    if (
      isNaN(numeroNota) ||
      numeroNota < 0 ||
      numeroNota > 10
    ) {
      throw new Error(
        "A nota deve estar entre 0 e 10."
      );
    }
  }

  const id =
    gerarID();

  getPlanilha()
    .getSheetByName("GE_PROVAS")
    .appendRow([

      id,

      texto(dados.data),

      texto(dados.titulo),

      nota,

      aluno.ra,

      aluno.nome,

      texto(dados.materia),

      usuario.ra,

      usuario.nome,

      aluno.serieAno,

      texto(dados.descricao)

    ]);

  return {

    sucesso: true,

    mensagem:
      "Prova cadastrada com sucesso!"
  };
}


function listarProvas(token) {

  const usuario =
    obterSessao(token);

  const dados =
    getPlanilha()
      .getSheetByName("GE_PROVAS")
      .getDataRange()
      .getDisplayValues();

  return dados
    .slice(1)
    .filter(function(linha) {

      if (usuario.tipo === "ALUNO") {

        return (
          texto(linha[4]) ===
          texto(usuario.ra)
        );
      }

      return (
        texto(linha[7]) ===
        texto(usuario.ra)
      );
    })
    .map(function(linha, index) {

      return {

        linha: index + 2,

        id: linha[0],

        data: linha[1],

        titulo: linha[2],

        nota: linha[3],

        alunoRA: linha[4],

        aluno: linha[5],

        materia: linha[6],

        professorRA: linha[7],

        professor: linha[8],

        serieAno: linha[9],

        descricao: linha[10]
      };
    });
}


function excluirProva(token, linha) {

  const usuario =
    obterSessao(token);

  if (usuario.tipo !== "PROFESSOR") {
    throw new Error(
      "Somente professores podem excluir."
    );
  }

  const numero =
    Number(linha);

  const aba =
    getPlanilha()
      .getSheetByName("GE_PROVAS");

  const dados =
    aba.getDataRange().getValues();

  if (
    numero < 2 ||
    numero > dados.length
  ) {
    throw new Error(
      "Prova inválida."
    );
  }

  if (
    texto(dados[numero - 1][7]) !==
    texto(usuario.ra)
  ) {
    throw new Error(
      "Você só pode excluir suas próprias provas."
    );
  }

  aba.deleteRow(numero);

  return {

    sucesso: true,

    mensagem:
      "Prova excluída!"
  };
}


// ============================================================
// TRABALHOS
// ============================================================

function cadastrarTrabalho(token, dados) {

  const usuario =
    obterSessao(token);

  if (usuario.tipo !== "PROFESSOR") {
    throw new Error(
      "Somente professores podem cadastrar trabalhos."
    );
  }

  const aluno =
    encontrarAluno(dados.alunoRA);

  if (!aluno) {
    throw new Error(
      "RA do aluno não encontrado."
    );
  }

  if (
    texto(aluno.serieAno) !==
    texto(usuario.serieAno)
  ) {
    throw new Error(
      "O aluno pertence a outra Série/Ano."
    );
  }

  if (!dados.data) {
    throw new Error(
      "Informe a data."
    );
  }

  if (!dados.titulo) {
    throw new Error(
      "Informe o título."
    );
  }

  if (!dados.materia) {
    throw new Error(
      "Informe a matéria."
    );
  }

  let nota =
    texto(dados.nota);

  if (nota !== "") {

    const numeroNota =
      Number(nota);

    if (
      isNaN(numeroNota) ||
      numeroNota < 0 ||
      numeroNota > 10
    ) {
      throw new Error(
        "A nota deve estar entre 0 e 10."
      );
    }
  }

  const id =
    gerarID();

  getPlanilha()
    .getSheetByName("GE_TRABALHOS")
    .appendRow([

      id,

      texto(dados.data),

      texto(dados.titulo),

      nota,

      aluno.ra,

      aluno.nome,

      texto(dados.materia),

      usuario.ra,

      usuario.nome,

      aluno.serieAno,

      texto(dados.descricao)

    ]);

  return {

    sucesso: true,

    mensagem:
      "Trabalho cadastrado com sucesso!"
  };
}


function listarTrabalhos(token) {

  const usuario =
    obterSessao(token);

  const dados =
    getPlanilha()
      .getSheetByName("GE_TRABALHOS")
      .getDataRange()
      .getDisplayValues();

  return dados
    .slice(1)
    .filter(function(linha) {

      if (usuario.tipo === "ALUNO") {

        return (
          texto(linha[4]) ===
          texto(usuario.ra)
        );
      }

      return (
        texto(linha[7]) ===
        texto(usuario.ra)
      );
    })
    .map(function(linha, index) {

      return {

        linha: index + 2,

        id: linha[0],

        data: linha[1],

        titulo: linha[2],

        nota: linha[3],

        alunoRA: linha[4],

        aluno: linha[5],

        materia: linha[6],

        professorRA: linha[7],

        professor: linha[8],

        serieAno: linha[9],

        descricao: linha[10]
      };
    });
}


function excluirTrabalho(token, linha) {

  const usuario =
    obterSessao(token);

  if (usuario.tipo !== "PROFESSOR") {
    throw new Error(
      "Somente professores podem excluir."
    );
  }

  const numero =
    Number(linha);

  const aba =
    getPlanilha()
      .getSheetByName("GE_TRABALHOS");

  const dados =
    aba.getDataRange().getValues();

  if (
    numero < 2 ||
    numero > dados.length
  ) {
    throw new Error(
      "Trabalho inválido."
    );
  }

  if (
    texto(dados[numero - 1][7]) !==
    texto(usuario.ra)
  ) {
    throw new Error(
      "Você só pode excluir seus próprios trabalhos."
    );
  }

  aba.deleteRow(numero);

  return {

    sucesso: true,

    mensagem:
      "Trabalho excluído!"
  };
}


// ============================================================
// FEEDBACKS
// ============================================================

function cadastrarFeedback(token, dados) {

  const usuario =
    obterSessao(token);

  if (usuario.tipo !== "ALUNO") {

    throw new Error(
      "Somente alunos podem enviar feedback."
    );
  }

  const tipo =
    texto(dados.tipo).toUpperCase();

  const itemID =
    texto(dados.itemID);

  const feedback =
    texto(dados.feedback);

  if (
    tipo !== "PROVA" &&
    tipo !== "TRABALHO"
  ) {
    throw new Error(
      "Tipo de feedback inválido."
    );
  }

  if (!itemID) {
    throw new Error(
      "Item não encontrado."
    );
  }

  if (!feedback) {
    throw new Error(
      "Digite seu feedback."
    );
  }

  if (feedback.length > 1000) {
    throw new Error(
      "O feedback pode ter no máximo 1000 caracteres."
    );
  }

  const nomeAba =
    tipo === "PROVA"
      ? "GE_PROVAS"
      : "GE_TRABALHOS";

  const aba =
    getPlanilha()
      .getSheetByName(nomeAba);

  const dadosAba =
    aba.getDataRange()
      .getValues();

  let registro = null;

  for (
    let i = 1;
    i < dadosAba.length;
    i++
  ) {

    if (
      texto(dadosAba[i][0]) === itemID &&
      texto(dadosAba[i][4]) ===
      texto(usuario.ra)
    ) {

      registro = dadosAba[i];

      break;
    }
  }

  if (!registro) {

    throw new Error(
      "Você não pode enviar feedback para este registro."
    );
  }

  const professorRA =
    texto(registro[7]);

  const professor =
    texto(registro[8]);

  const serie =
    texto(registro[9]);

  const feedbackAba =
    getPlanilha()
      .getSheetByName("GE_FEEDBACKS");

  const registrosFeedback =
    feedbackAba
      .getDataRange()
      .getValues();

  let encontrou = false;

  for (
    let i = 1;
    i < registrosFeedback.length;
    i++
  ) {

    if (
      texto(registrosFeedback[i][2]) === tipo &&
      texto(registrosFeedback[i][3]) === itemID &&
      texto(registrosFeedback[i][4]) ===
      texto(usuario.ra)
    ) {

      feedbackAba
        .getRange(i + 1, 2)
        .setValue(
          new Date()
        );

      feedbackAba
        .getRange(i + 1, 10)
        .setValue(feedback);

      encontrou = true;

      break;
    }
  }

  if (!encontrou) {

    feedbackAba.appendRow([

      gerarID(),

      new Date(),

      tipo,

      itemID,

      usuario.ra,

      usuario.nome,

      professorRA,

      professor,

      serie,

      feedback

    ]);
  }

  return {

    sucesso: true,

    mensagem:
      "Feedback enviado para o professor!"
  };
}


// ============================================================
// LISTAR FEEDBACKS
// ============================================================

function listarFeedbacks(token) {

  const usuario =
    obterSessao(token);

  const aba =
    getPlanilha()
      .getSheetByName("GE_FEEDBACKS");

  const dados =
    aba.getDataRange()
      .getDisplayValues();

  return dados
    .slice(1)
    .filter(function(linha) {

      if (usuario.tipo === "ALUNO") {

        return (
          texto(linha[4]) ===
          texto(usuario.ra)
        );
      }

      return (
        texto(linha[6]) ===
        texto(usuario.ra)
      );
    })
    .map(function(linha, index) {

      return {

        linha: index + 2,

        id: linha[0],

        data: linha[1],

        tipo: linha[2],

        itemID: linha[3],

        alunoRA: linha[4],

        aluno: linha[5],

        professorRA: linha[6],

        professor: linha[7],

        serieAno: linha[8],

        feedback: linha[9]
      };
    });
}


// ============================================================
// ESTATÍSTICAS
// ============================================================

function obterEstatisticas(token) {

  const usuario =
    obterSessao(token);

  const materias =
    listarMaterias(token);

  const provas =
    listarProvas(token);

  const trabalhos =
    listarTrabalhos(token);

  let alunos = [];

  if (usuario.tipo === "PROFESSOR") {
    alunos =
      listarAlunos(token);
  }

  return {

    materias: materias.length,

    provas: provas.length,

    trabalhos: trabalhos.length,

    alunos: alunos.length
  };
}

<!DOCTYPE html>
<html lang="pt-BR">

<head>

<meta charset="UTF-8">

<meta name="viewport"
content="width=device-width, initial-scale=1">

<base target="_top">

<title>Gestão Escolar</title>

<style>

/* ==========================================================
   RESET
========================================================== */

* {
  box-sizing: border-box;
  margin: 0;
  padding: 0;
}

body {
  font-family:
    Arial,
    Helvetica,
    sans-serif;

  background: #f1f5f9;

  color: #0f172a;
}

.hidden {
  display: none !important;
}

button,
input,
select,
textarea {
  font-family: inherit;
}


/* ==========================================================
   LOGIN
========================================================== */

.login-page {

  min-height: 100vh;

  display: flex;

  justify-content: center;

  align-items: center;

  padding: 20px;

  background:
    linear-gradient(
      135deg,
      #020617,
      #1e3a8a,
      #2563eb
    );
}

.login-card {

  width: 100%;

  max-width: 440px;

  background: white;

  padding: 35px;

  border-radius: 24px;

  box-shadow:
    0 30px 80px
    rgba(0,0,0,.30);
}

.logo {

  width: 72px;

  height: 72px;

  margin:
    0 auto 20px;

  display: flex;

  justify-content: center;

  align-items: center;

  border-radius: 20px;

  background:
    linear-gradient(
      135deg,
      #2563eb,
      #7c3aed
    );

  color: white;

  font-size: 34px;
}

h1 {

  text-align: center;

  margin-bottom: 8px;
}

.subtitle {

  text-align: center;

  color: #64748b;

  margin-bottom: 25px;
}


/* ==========================================================
   TYPE
========================================================== */

.type-buttons {

  display: grid;

  grid-template-columns:
    1fr 1fr;

  gap: 8px;

  margin-bottom: 20px;
}

.type-buttons button {

  padding: 12px;

  border:
    1px solid #cbd5e1;

  background: white;

  border-radius: 10px;

  cursor: pointer;

  font-weight: bold;
}

.type-buttons button.active {

  background: #eff6ff;

  color: #2563eb;

  border-color: #2563eb;
}


/* ==========================================================
   FORM
========================================================== */

label {

  display: block;

  margin:
    13px 0 6px;

  font-size: 14px;

  font-weight: bold;
}

input,
select,
textarea {

  width: 100%;

  padding: 12px;

  border:
    1px solid #cbd5e1;

  border-radius: 10px;

  outline: none;

  font-size: 15px;
}

textarea {

  min-height: 120px;

  resize: vertical;
}

input:focus,
select:focus,
textarea:focus {

  border-color: #2563eb;

  box-shadow:
    0 0 0 3px
    rgba(37,99,235,.12);
}


/* ==========================================================
   BUTTONS
========================================================== */

.btn {

  border: none;

  padding:
    12px 16px;

  border-radius: 10px;

  cursor: pointer;

  font-weight: bold;
}

.btn-primary {

  background: #2563eb;

  color: white;
}

.btn-primary:hover {
  background: #1d4ed8;
}

.btn-green {

  background: #16a34a;

  color: white;
}

.btn-red {

  background: #dc2626;

  color: white;
}

.btn-gray {

  background: #64748b;

  color: white;
}

.full {

  width: 100%;

  margin-top: 15px;
}

.link-button {

  width: 100%;

  border: none;

  background: transparent;

  color: #2563eb;

  padding: 12px;

  cursor: pointer;

  font-weight: bold;
}


/* ==========================================================
   MESSAGE
========================================================== */

.message {

  margin-top: 12px;

  padding: 11px;

  border-radius: 8px;

  display: none;

  font-size: 14px;
}

.message.show {
  display: block;
}

.message.error {

  background: #fee2e2;

  color: #991b1b;
}

.message.success {

  background: #dcfce7;

  color: #166534;
}


/* ==========================================================
   APP
========================================================== */

#app {
  display: none;
}

.layout {

  min-height: 100vh;

  display: flex;
}


/* ==========================================================
   SIDEBAR
========================================================== */

.sidebar {

  width: 255px;

  position: fixed;

  top: 0;

  left: 0;

  bottom: 0;

  background: #0f172a;

  color: white;

  padding: 20px;

  z-index: 30;
}

.brand {

  display: flex;

  align-items: center;

  gap: 10px;

  padding-bottom: 22px;

  border-bottom:
    1px solid
    rgba(255,255,255,.1);
}

.brand-icon {

  width: 42px;

  height: 42px;

  display: flex;

  justify-content: center;

  align-items: center;

  border-radius: 10px;

  background: #2563eb;
}

.brand-text {
  font-weight: bold;
}

.brand-small {

  display: block;

  color: #94a3b8;

  font-size: 11px;

  margin-top: 3px;
}

.section-title {

  color: #64748b;

  font-size: 11px;

  font-weight: bold;

  margin:
    22px 8px 8px;

  text-transform: uppercase;
}

.nav {

  width: 100%;

  padding: 12px;

  margin-bottom: 5px;

  background: transparent;

  color: #cbd5e1;

  border: none;

  border-radius: 9px;

  text-align: left;

  cursor: pointer;

  font-weight: bold;
}

.nav:hover,
.nav.active {

  background: #1d4ed8;

  color: white;
}


/* ==========================================================
   MAIN
========================================================== */

.main {

  margin-left: 255px;

  width:
    calc(100% - 255px);
}

.top {

  height: 72px;

  padding:
    0 25px;

  display: flex;

  align-items: center;

  justify-content: space-between;

  background: white;

  border-bottom:
    1px solid #e2e8f0;

  position: sticky;

  top: 0;

  z-index: 20;
}

.user {

  display: flex;

  align-items: center;

  gap: 10px;
}

.avatar {

  width: 40px;

  height: 40px;

  border-radius: 50%;

  display: flex;

  justify-content: center;

  align-items: center;

  background: #2563eb;

  color: white;

  font-weight: bold;
}

.user-name {
  font-weight: bold;
}

.user-type {

  display: block;

  font-size: 11px;

  color: #64748b;
}


/* ==========================================================
   CONTENT
========================================================== */

.content {

  padding: 25px;

  max-width: 1450px;

  margin: auto;
}

.screen {
  display: none;
}

.screen.active {
  display: block;
}


/* ==========================================================
   HERO
========================================================== */

.hero {

  padding: 30px;

  border-radius: 18px;

  color: white;

  margin-bottom: 20px;

  background:
    linear-gradient(
      135deg,
      #2563eb,
      #4f46e5
    );
}

.hero h2 {

  text-align: left;

  margin-bottom: 8px;
}

.hero p {
  color: #dbeafe;
}


/* ==========================================================
   STATS
========================================================== */

.stats {

  display: grid;

  grid-template-columns:
    repeat(
      auto-fit,
      minmax(180px,1fr)
    );

  gap: 15px;

  margin-bottom: 20px;
}

.stat {

  padding: 20px;

  border-radius: 15px;

  background: white;

  box-shadow:
    0 5px 18px
    rgba(15,23,42,.06);
}

.stat-icon {
  font-size: 28px;
}

.stat-number {

  font-size: 25px;

  font-weight: bold;

  margin-top: 10px;
}

.stat-label {

  color: #64748b;

  font-size: 13px;
}


/* ==========================================================
   CARD
========================================================== */

.card {

  background: white;

  padding: 22px;

  border-radius: 15px;

  margin-bottom: 20px;

  box-shadow:
    0 5px 18px
    rgba(15,23,42,.06);
}

.card-header {

  display: flex;

  align-items: center;

  justify-content: space-between;

  gap: 10px;

  margin-bottom: 18px;
}

.card-header h2 {

  text-align: left;

  font-size: 19px;
}


/* ==========================================================
   GRID
========================================================== */

.grid {

  display: grid;

  grid-template-columns:
    repeat(
      auto-fit,
      minmax(280px,1fr)
    );

  gap: 15px;
}


/* ==========================================================
   MATÉRIA
========================================================== */

.subject {

  border:
    1px solid #e2e8f0;

  border-radius: 14px;

  overflow: hidden;

  background: white;
}

.subject-head {

  padding: 18px;

  background: #eff6ff;
}

.subject-head h3 {

  color: #1e3a8a;

  margin-bottom: 6px;
}

.subject-meta {

  color: #64748b;

  font-size: 12px;
}

.subject-body {
  padding: 18px;
}

.subject-content {

  background: #f8fafc;

  padding: 13px;

  border-radius: 8px;

  white-space: pre-line;

  line-height: 1.6;
}

.links {

  border-top:
    1px solid #e2e8f0;

  margin-top: 15px;

  padding-top: 15px;
}

.study-link {

  display: block;

  padding: 10px;

  margin-top: 7px;

  border-radius: 8px;

  text-decoration: none;

  background: #eff6ff;

  color: #1d4ed8;

  word-break: break-all;

  font-size: 13px;
}

.study-link:hover {
  background: #dbeafe;
}


/* ==========================================================
   ITEMS
========================================================== */

.item {

  border:
    1px solid #e2e8f0;

  border-radius: 12px;

  padding: 18px;

  background: white;
}

.item h3 {

  color: #1e3a8a;

  margin-bottom: 10px;
}

.item p {

  color: #475569;

  font-size: 14px;

  margin: 6px 0;
}

.note {

  display: inline-block;

  margin-top: 8px;

  padding:
    6px 10px;

  border-radius: 15px;

  background: #dcfce7;

  color: #166534;

  font-weight: bold;

  font-size: 13px;
}


/* ==========================================================
   FEEDBACK
========================================================== */

.feedback-box {

  margin-top: 18px;

  padding-top: 16px;

  border-top:
    1px solid #e2e8f0;
}

.feedback-title {

  font-weight: bold;

  margin-bottom: 8px;
}

.feedback-input {

  min-height: 90px;

  margin-bottom: 8px;
}

.feedback-enviado {

  background: #f0fdf4;

  border:
    1px solid #bbf7d0;

  padding: 12px;

  border-radius: 10px;

  margin-top: 12px;

  color: #166534;

  white-space: pre-line;
}


/* ==========================================================
   EMPTY
========================================================== */

.empty {

  grid-column: 1 / -1;

  text-align: center;

  padding: 40px;

  color: #64748b;
}


/* ==========================================================
   MOBILE
========================================================== */

#menuMobile {
  display: none;
}

@media(max-width:850px) {

  .sidebar {

    transform:
      translateX(-100%);

    transition: .2s;
  }

  .sidebar.open {

    transform:
      translateX(0);
  }

  .main {

    margin-left: 0;

    width: 100%;
  }

  #menuMobile {

    display: block;

    border: none;

    background: #f1f5f9;

    padding: 8px 12px;

    border-radius: 8px;

    cursor: pointer;
  }
}

@media(max-width:600px) {

  .content {
    padding: 15px;
  }

  .top {
    padding: 0 15px;
  }

  .user-type {
    display: none;
  }

}

</style>

</head>

<body>


<!-- ==========================================================
     LOGIN
========================================================== -->

<div id="loginPage" class="login-page">

  <div class="login-card">

    <div class="logo">
      🏫
    </div>

    <h1>
      Gestão Escolar
    </h1>

    <p class="subtitle">
      Acesse sua conta
    </p>

    <div class="type-buttons">

      <button id="loginAluno"
      class="active">
        👨‍🎓 Aluno
      </button>

      <button id="loginProfessor">
        👨‍🏫 Professor
      </button>

    </div>

    <label>
      Usuário
    </label>

    <input
      id="loginUsuario"
      placeholder="Digite seu usuário">

    <label>
      Senha
    </label>

    <input
      id="loginSenha"
      type="password"
      placeholder="Digite sua senha">

    <button
      id="entrarBtn"
      class="btn btn-primary full">
      🔐 Entrar
    </button>

    <button
      id="cadastroBtn"
      class="link-button">
      📝 Criar uma conta
    </button>

    <div
      id="loginMsg"
      class="message">
    </div>

  </div>

</div>


<!-- ==========================================================
     CADASTRO
========================================================== -->

<div
  id="registerPage"
  class="login-page hidden">

  <div class="login-card">

    <div class="logo">
      📝
    </div>

    <h1>
      Criar conta
    </h1>

    <p class="subtitle">
      Escolha o tipo de usuário
    </p>

    <div class="type-buttons">

      <button
        id="cadAluno"
        class="active">
        👨‍🎓 Aluno
      </button>

      <button
        id="cadProfessor">
        👨‍🏫 Professor
      </button>

    </div>

    <label>
      RA / ID
    </label>

    <input
      id="cadRA"
      placeholder="RA do aluno ou ID do professor">

    <label>
      Nome completo
    </label>

    <input
      id="cadNome"
      placeholder="Digite seu nome">

    <label>
      Usuário
    </label>

    <input
      id="cadUsuario"
      placeholder="Escolha seu usuário">

    <label>
      Senha
    </label>

    <input
      id="cadSenha"
      type="password"
      placeholder="Mínimo 6 caracteres">

    <label>
      Série / Ano
    </label>

    <select id="cadSerie">

      <option value="">
        Selecione
      </option>

      <option>6º Ano</option>
      <option>7º Ano</option>
      <option>8º Ano</option>
      <option>9º Ano</option>

      <option>1ª Série</option>
      <option>2ª Série</option>
      <option>3ª Série</option>

    </select>

    <div
      id="cadMateriaBox"
      class="hidden">

      <label>
        Matéria
      </label>

      <input
        id="cadMateria"
        placeholder="Ex: Matemática">

    </div>

    <button
      id="salvarCadastro"
      class="btn btn-green full">
      ✅ Criar conta
    </button>

    <button
      id="voltarLogin"
      class="btn btn-gray full">
      ← Voltar
    </button>

    <div
      id="cadMsg"
      class="message">
    </div>

  </div>

</div>


<!-- ==========================================================
     APP
========================================================== -->

<div id="app">

  <div class="layout">

    <aside
      id="sidebar"
      class="sidebar">

      <div class="brand">

        <div class="brand-icon">
          🏫
        </div>

        <div>

          <div class="brand-text">
            Gestão Escolar
          </div>

          <span class="brand-small">
            Sistema acadêmico
          </span>

        </div>

      </div>

      <div class="section-title">
        Principal
      </div>

      <button
        class="nav active"
        data-screen="dashboard">
        🏠 Dashboard
      </button>

      <button
        class="nav"
        data-screen="materias">
        📚 Matérias
      </button>

      <div class="section-title">
        Acadêmico
      </div>

      <button
        class="nav"
        data-screen="provas">
        📝 Provas e Notas
      </button>

      <button
        class="nav"
        data-screen="trabalhos">
        📂 Trabalhos
      </button>

      <button
        id="feedbackNav"
        class="nav"
        data-screen="feedbacks">
        💬 Feedbacks
      </button>

      <button
        id="alunosNav"
        class="nav"
        data-screen="alunos">
        👨‍🎓 Alunos
      </button>

      <div class="section-title">
        Conta
      </div>

      <button
        id="logoutBtn"
        class="nav">
        🚪 Sair
      </button>

    </aside>


    <main class="main">

      <div class="top">

        <button id="menuMobile">
          ☰
        </button>

        <strong id="pageTitle">
          Dashboard
        </strong>

        <div class="user">

          <div
            id="avatar"
            class="avatar">
            U
          </div>

          <div>

            <div
              id="userName"
              class="user-name">
              Usuário
            </div>

            <span
              id="userType"
              class="user-type">
              Conta
            </span>

          </div>

        </div>

      </div>


      <div class="content">


        <!-- ==================================================
             DASHBOARD
        ================================================== -->

        <section
          id="screenDashboard"
          class="screen active">

          <div class="hero">

            <h2 id="hello">
              Olá!
            </h2>

            <p id="heroText">
              Bem-vindo ao sistema.
            </p>

          </div>

          <div class="stats">

            <div class="stat">

              <div class="stat-icon">
                📚
              </div>

              <div
                id="materiaCount"
                class="stat-number">
                0
              </div>

              <div class="stat-label">
                Matérias
              </div>

            </div>

            <div class="stat">

              <div class="stat-icon">
                📝
              </div>

              <div
                id="provaCount"
                class="stat-number">
                0
              </div>

              <div class="stat-label">
                Provas
              </div>

            </div>

            <div class="stat">

              <div class="stat-icon">
                📂
              </div>

              <div
                id="trabalhoCount"
                class="stat-number">
                0
              </div>

              <div class="stat-label">
                Trabalhos
              </div>

            </div>

            <div
              id="alunoStat"
              class="stat">

              <div class="stat-icon">
                👨‍🎓
              </div>

              <div
                id="alunoCount"
                class="stat-number">
                0
              </div>

              <div class="stat-label">
                Alunos
              </div>

            </div>

          </div>

        </section>


        <!-- ==================================================
             MATÉRIAS
        ================================================== -->

        <section
          id="screenMaterias"
          class="screen">

          <div
            id="materiaForm"
            class="card">

            <div class="card-header">

              <div>

                <h2>
                  📚 Nova matéria
                </h2>

                <p>
                  Publique conteúdo e links para os alunos.
                </p>

              </div>

            </div>

            <label>
              Nome da matéria
            </label>

            <input
              id="materiaNome"
              placeholder="Ex: Matemática">

            <label>
              Conteúdo da aula
            </label>

            <textarea
              id="materiaConteudo"
              placeholder="Digite o conteúdo da aula...">
            </textarea>

            <label>
              Links de estudo
            </label>

            <textarea
              id="materiaLinks"
              placeholder="Um link por linha">
            </textarea>

            <button
              id="materiaSalvar"
              class="btn btn-primary">
              💾 Publicar matéria
            </button>

          </div>


          <div class="card">

            <div class="card-header">

              <h2>
                📚 Matérias
              </h2>

              <button
                id="materiaAtualizar"
                class="btn btn-gray">
                🔄
              </button>

            </div>

            <div
              id="materiasGrid"
              class="grid">
            </div>

          </div>

        </section>


        <!-- ==================================================
             PROVAS
        ================================================== -->

        <section
          id="screenProvas"
          class="screen">

          <div
            id="provaForm"
            class="card">

            <h2>
              📝 Lançar nota da prova
            </h2>

            <label>
              Data
            </label>

            <input
              id="provaData"
              type="date">

            <label>
              Título
            </label>

            <input
              id="provaTitulo"
              placeholder="Ex: Prova de Matemática">

            <label>
              Aluno
            </label>

            <select id="provaAluno">
              <option value="">
                Selecione o aluno
              </option>
            </select>

            <label>
              Matéria
            </label>

            <input
              id="provaMateria"
              placeholder="Ex: Matemática">

            <label>
              Nota
            </label>

            <input
              id="provaNota"
              type="number"
              min="0"
              max="10"
              step="0.01"
              placeholder="0 a 10">

            <label>
              Descrição
            </label>

            <textarea
              id="provaDescricao">
            </textarea>

            <button
              id="provaSalvar"
              class="btn btn-primary">
              💾 Lançar nota
            </button>

          </div>


          <div class="card">

            <div class="card-header">

              <h2>
                📝 Provas e Notas
              </h2>

              <button
                id="provaAtualizar"
                class="btn btn-gray">
                🔄
              </button>

            </div>

            <div
              id="provasGrid"
              class="grid">
            </div>

          </div>

        </section>


        <!-- ==================================================
             TRABALHOS
        ================================================== -->

        <section
          id="screenTrabalhos"
          class="screen">

          <div
            id="trabalhoForm"
            class="card">

            <h2>
              📂 Novo trabalho
            </h2>

            <label>
              Data
            </label>

            <input
              id="trabalhoData"
              type="date">

            <label>
              Título
            </label>

            <input
              id="trabalhoTitulo"
              placeholder="Ex: Trabalho de História">

            <label>
              Aluno
            </label>

            <select id="trabalhoAluno">
              <option value="">
                Selecione o aluno
              </option>
            </select>

            <label>
              Matéria
            </label>

            <input
              id="trabalhoMateria"
              placeholder="Ex: História">

            <label>
              Nota
            </label>

            <input
              id="trabalhoNota"
              type="number"
              min="0"
              max="10"
              step="0.01">

            <label>
              Descrição
            </label>

            <textarea
              id="trabalhoDescricao">
            </textarea>

            <button
              id="trabalhoSalvar"
              class="btn btn-primary">
              💾 Cadastrar trabalho
            </button>

          </div>


          <div class="card">

            <div class="card-header">

              <h2>
                📂 Trabalhos
              </h2>

              <button
                id="trabalhoAtualizar"
                class="btn btn-gray">
                🔄
              </button>

            </div>

            <div
              id="trabalhosGrid"
              class="grid">
            </div>

          </div>

        </section>


        <!-- ==================================================
             FEEDBACKS
        ================================================== -->

        <section
          id="screenFeedbacks"
          class="screen">

          <div class="card">

            <div class="card-header">

              <div>

                <h2>
                  💬 Feedbacks
                </h2>

                <p>
                  Feedbacks enviados pelos alunos.
                </p>

              </div>

              <button
                id="feedbackAtualizar"
                class="btn btn-gray">
                🔄
              </button>

            </div>

            <div
              id="feedbacksGrid"
              class="grid">
            </div>

          </div>

        </section>


        <!-- ==================================================
             ALUNOS
        ================================================== -->

        <section
          id="screenAlunos"
          class="screen">

          <div class="card">

            <div class="card-header">

              <h2>
                👨‍🎓 Alunos da minha Série/Ano
              </h2>

              <button
                id="alunosAtualizar"
                class="btn btn-gray">
                🔄
              </button>

            </div>

            <div
              id="alunosGrid"
              class="grid">
            </div>

          </div>

        </section>


      </div>

    </main>

  </div>

</div>


<script>

// ==========================================================
// VARIÁVEIS
// ==========================================================

let loginTipo = "ALUNO";

let cadastroTipo = "ALUNO";

let token = null;

let usuario = null;


// ==========================================================
// UTILIDADES
// ==========================================================

function el(id) {
  return document.getElementById(id);
}


function mensagem(id, texto, sucesso) {

  const box = el(id);

  box.textContent = texto;

  box.className =
    "message show " +
    (sucesso ? "success" : "error");
}


function escapeHTML(value) {

  return String(value || "")

    .replace(/&/g, "&amp;")

    .replace(/</g, "&lt;")

    .replace(/>/g, "&gt;")

    .replace(/"/g, "&quot;")

    .replace(/'/g, "&#039;");
}


// ==========================================================
// LOGIN TIPO
// ==========================================================

el("loginAluno").onclick =
function() {

  loginTipo = "ALUNO";

  el("loginAluno")
    .classList.add("active");

  el("loginProfessor")
    .classList.remove("active");
};


el("loginProfessor").onclick =
function() {

  loginTipo = "PROFESSOR";

  el("loginProfessor")
    .classList.add("active");

  el("loginAluno")
    .classList.remove("active");
};


// ==========================================================
// CADASTRO
// ==========================================================

el("cadastroBtn").onclick =
function() {

  el("loginPage")
    .classList.add("hidden");

  el("registerPage")
    .classList.remove("hidden");
};


el("voltarLogin").onclick =
function() {

  el("registerPage")
    .classList.add("hidden");

  el("loginPage")
    .classList.remove("hidden");
};


el("cadAluno").onclick =
function() {

  cadastroTipo = "ALUNO";

  el("cadAluno")
    .classList.add("active");

  el("cadProfessor")
    .classList.remove("active");

  el("cadMateriaBox")
    .classList.add("hidden");
};


el("cadProfessor").onclick =
function() {

  cadastroTipo = "PROFESSOR";

  el("cadProfessor")
    .classList.add("active");

  el("cadAluno")
    .classList.remove("active");

  el("cadMateriaBox")
    .classList.remove("hidden");
};


// ==========================================================
// SALVAR CADASTRO
// ==========================================================

el("salvarCadastro").onclick =
function() {

  const dados = {

    ra:
      el("cadRA").value.trim(),

    nome:
      el("cadNome").value.trim(),

    usuario:
      el("cadUsuario").value.trim(),

    senha:
      el("cadSenha").value,

    serieAno:
      el("cadSerie").value,

    materia:
      el("cadMateria").value.trim(),

    tipo:
      cadastroTipo
  };

  if (!dados.ra) {
    mensagem(
      "cadMsg",
      "Informe o RA/ID.",
      false
    );
    return;
  }

  if (!dados.nome) {
    mensagem(
      "cadMsg",
      "Informe o nome.",
      false
    );
    return;
  }

  if (!dados.usuario) {
    mensagem(
      "cadMsg",
      "Informe o usuário.",
      false
    );
    return;
  }

  if (dados.senha.length < 6) {
    mensagem(
      "cadMsg",
      "A senha precisa ter pelo menos 6 caracteres.",
      false
    );
    return;
  }

  if (!dados.serieAno) {
    mensagem(
      "cadMsg",
      "Selecione a Série/Ano.",
      false
    );
    return;
  }

  if (
    cadastroTipo === "PROFESSOR" &&
    !dados.materia
  ) {

    mensagem(
      "cadMsg",
      "Informe a matéria.",
      false
    );

    return;
  }

  mensagem(
    "cadMsg",
    "Criando conta...",
    true
  );

  google.script.run

    .withSuccessHandler(
      function(resultado) {

        mensagem(
          "cadMsg",
          resultado.mensagem,
          resultado.sucesso
        );

        if (resultado.sucesso) {

          el("cadRA").value = "";

          el("cadNome").value = "";

          el("cadUsuario").value = "";

          el("cadSenha").value = "";

          el("cadSerie").value = "";

          el("cadMateria").value = "";

          setTimeout(
            function() {

              el("registerPage")
                .classList.add("hidden");

              el("loginPage")
                .classList.remove("hidden");

            },
            1200
          );
        }
      }
    )

    .withFailureHandler(
      function(error) {

        mensagem(
          "cadMsg",
          error.message,
          false
        );
      }
    )

    .cadastrarUsuario(dados);
};


// ==========================================================
// LOGIN
// ==========================================================

el("entrarBtn").onclick =
fazerLogin;


el("loginSenha").onkeydown =
function(event) {

  if (event.key === "Enter") {
    fazerLogin();
  }
};


function fazerLogin() {

  const usuarioDigitado =
    el("loginUsuario")
      .value
      .trim();

  const senha =
    el("loginSenha")
      .value;

  if (
    !usuarioDigitado ||
    !senha
  ) {

    mensagem(
      "loginMsg",
      "Informe usuário e senha.",
      false
    );

    return;
  }

  mensagem(
    "loginMsg",
    "Entrando...",
    true
  );

  google.script.run

    .withSuccessHandler(
      function(resultado) {

        if (!resultado.sucesso) {

          mensagem(
            "loginMsg",
            resultado.mensagem,
            false
          );

          return;
        }

        token =
          resultado.token;

        usuario =
          resultado.usuario;

        localStorage.setItem(
          "GE_TOKEN",
          token
        );

        abrirApp();
      }
    )

    .withFailureHandler(
      function(error) {

        mensagem(
          "loginMsg",
          error.message,
          false
        );
      }
    )

    .login({

      usuario:
        usuarioDigitado,

      senha:
        senha,

      tipo:
        loginTipo
    });
}


// ==========================================================
// ABRIR APP
// ==========================================================

function abrirApp() {

  el("loginPage")
    .classList.add("hidden");

  el("registerPage")
    .classList.add("hidden");

  el("app")
    .style.display = "block";

  el("userName")
    .textContent =
    usuario.nome;

  el("userType")
    .textContent =
    usuario.tipo +
    " • " +
    usuario.serieAno;

  el("avatar")
    .textContent =
    usuario.nome
      .charAt(0)
      .toUpperCase();

  el("hello")
    .textContent =
    "Olá, " +
    usuario.nome +
    " 👋";

  el("heroText")
    .textContent =

    usuario.tipo === "PROFESSOR"

      ? "Gerencie suas matérias, notas, provas e trabalhos."

      : "Veja suas notas, matérias, trabalhos e envie feedback aos professores.";

  const professor =
    usuario.tipo === "PROFESSOR";

  el("materiaForm")
    .style.display =
    professor
      ? "block"
      : "none";

  el("provaForm")
    .style.display =
    professor
      ? "block"
      : "none";

  el("trabalhoForm")
    .style.display =
    professor
      ? "block"
      : "none";

  el("alunosNav")
    .style.display =
    professor
      ? "block"
      : "none";

  el("alunoStat")
    .style.display =
    professor
      ? "block"
      : "none";

  carregarMaterias();

  carregarProvas();

  carregarTrabalhos();

  carregarFeedbacks();

  if (professor) {
    carregarAlunos();
  }
}


// ==========================================================
// NAVEGAÇÃO
// ==========================================================

document
  .querySelectorAll(".nav[data-screen]")
  .forEach(
    function(button) {

      button.onclick =
      function() {

        document
          .querySelectorAll(".screen")
          .forEach(
            function(screen) {

              screen
                .classList
                .remove("active");
            }
          );

        const nome =
          this.dataset.screen;

        el(
          "screen" +
          nome.charAt(0).toUpperCase() +
          nome.slice(1)
        )
        .classList
        .add("active");

        document
          .querySelectorAll(".nav")
          .forEach(
            function(nav) {

              nav.classList
                .remove("active");
            }
          );

        this.classList.add("active");

        const titulos = {

          dashboard:
            "Dashboard",

          materias:
            "Matérias",

          provas:
            "Provas e Notas",

          trabalhos:
            "Trabalhos",

          feedbacks:
            "Feedbacks",

          alunos:
            "Alunos"
        };

        el("pageTitle")
          .textContent =
          titulos[nome];

        if (nome === "materias")
          carregarMaterias();

        if (nome === "provas") {

          carregarProvas();

          if (usuario.tipo === "PROFESSOR") {
            carregarAlunos();
          }
        }

        if (nome === "trabalhos") {

          carregarTrabalhos();

          if (usuario.tipo === "PROFESSOR") {
            carregarAlunos();
          }
        }

        if (nome === "feedbacks")
          carregarFeedbacks();

        if (nome === "alunos")
          carregarAlunos();

      };
    }
  );


// ==========================================================
// MATÉRIAS
// ==========================================================

el("materiaSalvar").onclick =
function() {

  const dados = {

    nome:
      el("materiaNome")
        .value.trim(),

    conteudo:
      el("materiaConteudo")
        .value.trim(),

    links:
      el("materiaLinks")
        .value.trim()
  };

  if (!dados.nome) {

    alert(
      "Informe o nome da matéria."
    );

    return;
  }

  google.script.run

    .withSuccessHandler(
      function(resultado) {

        alert(
          resultado.mensagem
        );

        el("materiaNome").value = "";

        el("materiaConteudo").value = "";

        el("materiaLinks").value = "";

        carregarMaterias();
      }
    )

    .withFailureHandler(
      mostrarErro
    )

    .cadastrarMateria(
      token,
      dados
    );
};


el("materiaAtualizar").onclick =
carregarMaterias;


function carregarMaterias() {

  google.script.run

    .withSuccessHandler(
      function(materias) {

        el("materiaCount")
          .textContent =
          materias.length;

        const grid =
          el("materiasGrid");

        grid.innerHTML = "";

        if (!materias.length) {

          grid.innerHTML =
            '<div class="empty">' +
            "📚<br><br>" +
            "Nenhuma matéria encontrada." +
            "</div>";

          return;
        }

        materias.forEach(
          function(materia) {

            let linksHtml = "";

            const links =
              String(
                materia.links || ""
              )
              .split(/\n|,/)
              .map(function(x) {
                return x.trim();
              })
              .filter(Boolean);

            if (links.length) {

              linksHtml =
                '<div class="links">' +

                "<strong>🔗 Materiais de estudo</strong>" +

                links
                  .map(function(link) {

                    const url =
                      /^https?:\/\//i.test(link)
                        ? link
                        : "https://" + link;

                    return (

                      '<a class="study-link" ' +

                      'href="' +
                      escapeHTML(url) +
                      '" ' +

                      'target="_blank">' +

                      "🌐 " +

                      escapeHTML(link) +

                      "</a>"
                    );

                  })
                  .join("") +

                "</div>";
            }

            let botoes = "";

            if (
              usuario.tipo === "PROFESSOR" &&
              String(materia.professorRA) ===
              String(usuario.ra)
            ) {

              botoes =

                '<div style="margin-top:15px;display:flex;gap:8px">' +

                '<button class="btn btn-primary" ' +

                'onclick="editarMateria(\'' +
                escapeHTML(materia.linha) +
                '\')">' +

                "✏️ Editar" +

                "</button>" +

                '<button class="btn btn-red" ' +

                'onclick="excluirMateria(\'' +
                escapeHTML(materia.linha) +
                '\')">' +

                "🗑 Excluir" +

                "</button>" +

                "</div>";
            }

            const card =
              document.createElement("div");

            card.className =
              "subject";

            card.innerHTML =

              '<div class="subject-head">' +

              "<h3>" +

              "📚 " +

              escapeHTML(materia.nome) +

              "</h3>" +

              '<div class="subject-meta">' +

              "Professor: " +

              escapeHTML(materia.professor) +

              "<br>" +

              "Série/Ano: " +

              escapeHTML(materia.serieAno) +

              "</div>" +

              "</div>" +

              '<div class="subject-body">' +

              '<div class="subject-content">' +

              "<strong>📖 Conteúdo da aula</strong>" +

              "<br><br>" +

              escapeHTML(
                materia.conteudo ||
                "Nenhum conteúdo cadastrado."
              ) +

              "</div>" +

              linksHtml +

              botoes +

              "</div>";

            grid.appendChild(card);
          }
        );
      }
    )

    .withFailureHandler(
      mostrarErro
    )

    .listarMaterias(token);
}


function editarMateria(linha) {

  const nome =
    prompt(
      "Nome da matéria:"
    );

  if (nome === null)
    return;

  const conteudo =
    prompt(
      "Conteúdo da aula:"
    );

  if (conteudo === null)
    return;

  const links =
    prompt(
      "Links, um por linha:"
    );

  if (links === null)
    return;

  google.script.run

    .withSuccessHandler(
      function(resultado) {

        alert(
          resultado.mensagem
        );

        carregarMaterias();
      }
    )

    .withFailureHandler(
      mostrarErro
    )

    .atualizarMateria(
      token,
      {
        linha:
          Number(linha),

        nome:
          nome,

        conteudo:
          conteudo,

        links:
          links
      }
    );
}


function excluirMateria(linha) {

  if (
    !confirm(
      "Deseja excluir esta matéria?"
    )
  ) {
    return;
  }

  google.script.run

    .withSuccessHandler(
      function(resultado) {

        alert(
          resultado.mensagem
        );

        carregarMaterias();
      }
    )

    .withFailureHandler(
      mostrarErro
    )

    .excluirMateria(
      token,
      Number(linha)
    );
}


// ==========================================================
// ALUNOS
// ==========================================================

function carregarAlunos() {

  if (usuario.tipo !== "PROFESSOR")
    return;

  google.script.run

    .withSuccessHandler(
      function(alunos) {

        el("alunoCount")
          .textContent =
          alunos.length;

        const grid =
          el("alunosGrid");

        grid.innerHTML = "";

        preencherSelectAlunos(
          "provaAluno",
          alunos
        );

        preencherSelectAlunos(
          "trabalhoAluno",
          alunos
        );

        if (!alunos.length) {

          grid.innerHTML =
            '<div class="empty">' +
            "Nenhum aluno cadastrado nesta Série/Ano." +
            "</div>";

          return;
        }

        alunos.forEach(
          function(aluno) {

            const card =
              document.createElement("div");

            card.className =
              "item";

            card.innerHTML =

              "<h3>👨‍🎓 " +

              escapeHTML(aluno.nome) +

              "</h3>" +

              "<p><strong>RA:</strong> " +

              escapeHTML(aluno.ra) +

              "</p>" +

              "<p><strong>Série/Ano:</strong> " +

              escapeHTML(aluno.serieAno) +

              "</p>";

            grid.appendChild(card);
          }
        );
      }
    )

    .withFailureHandler(
      mostrarErro
    )

    .listarAlunos(token);
}


function preencherSelectAlunos(id, alunos) {

  const select =
    el(id);

  if (!select)
    return;

  select.innerHTML =
    '<option value="">Selecione o aluno</option>';

  alunos.forEach(
    function(aluno) {

      const option =
        document.createElement("option");

      option.value =
        aluno.ra;

      option.textContent =
        aluno.nome +
        " — RA: " +
        aluno.ra;

      select.appendChild(option);
    }
  );
}


el("alunosAtualizar").onclick =
carregarAlunos;


// ==========================================================
// PROVAS
// ==========================================================

el("provaSalvar").onclick =
function() {

  const dados = {

    data:
      el("provaData")
        .value,

    titulo:
      el("provaTitulo")
        .value.trim(),

    alunoRA:
      el("provaAluno")
        .value,

    materia:
      el("provaMateria")
        .value.trim(),

    nota:
      el("provaNota")
        .value,

    descricao:
      el("provaDescricao")
        .value.trim()
  };

  if (!dados.data) {

    alert("Informe a data.");

    return;
  }

  if (!dados.titulo) {

    alert("Informe o título.");

    return;
  }

  if (!dados.alunoRA) {

    alert("Selecione o aluno.");

    return;
  }

  if (!dados.materia) {

    alert("Informe a matéria.");

    return;
  }

  if (dados.nota === "") {

    alert("Informe a nota.");

    return;
  }

  google.script.run

    .withSuccessHandler(
      function(resultado) {

        alert(
          resultado.mensagem
        );

        el("provaData").value = "";

        el("provaTitulo").value = "";

        el("provaAluno").value = "";

        el("provaMateria").value = "";

        el("provaNota").value = "";

        el("provaDescricao").value = "";

        carregarProvas();
      }
    )

    .withFailureHandler(
      mostrarErro
    )

    .cadastrarProva(
      token,
      dados
    );
};


el("provaAtualizar").onclick =
carregarProvas;


function carregarProvas() {

  google.script.run

    .withSuccessHandler(
      function(provas) {

        el("provaCount")
          .textContent =
          provas.length;

        const grid =
          el("provasGrid");

        grid.innerHTML = "";

        if (!provas.length) {

          grid.innerHTML =
            '<div class="empty">' +

            "📝<br><br>" +

            "Nenhuma prova encontrada." +

            "</div>";

          return;
        }

        google.script.run

          .withSuccessHandler(
            function(feedbacks) {

              provas.forEach(
                function(prova) {

                  const card =
                    document.createElement("div");

                  card.className =
                    "item";

                  const feedback =
                    feedbacks.find(
                      function(f) {

                        return (
                          f.itemID ===
                          prova.id
                        );
                      }
                    );

                  const nota =
                    prova.nota !== ""

                      ? '<span class="note">⭐ Nota: ' +
                        escapeHTML(prova.nota) +
                        "</span>"

                      : '<span class="note">⏳ Nota não lançada</span>';

                  let feedbackHTML = "";

                  if (
                    usuario.tipo === "ALUNO"
                  ) {

                    feedbackHTML =

                      '<div class="feedback-box">' +

                      '<div class="feedback-title">' +

                      "💬 Seu feedback para o professor" +

                      "</div>" +

                      '<textarea ' +

                      'id="feedback_' +
                      escapeHTML(prova.id) +
                      '" ' +

                      'class="feedback-input" ' +

                      'placeholder="Escreva aqui sua opinião sobre a prova, dificuldade, conteúdo, correção etc.">' +

                      escapeHTML(
                        feedback
                          ? feedback.feedback
                          : ""
                      ) +

                      "</textarea>" +

                      '<button class="btn btn-primary" ' +

                      'onclick="enviarFeedback(\'' +
                      escapeHTML(prova.id) +
                      '\',\'PROVA\')">' +

                      "📨 Enviar feedback" +

                      "</button>";

                    if (feedback) {

                      feedbackHTML +=

                        '<div class="feedback-enviado">' +

                        "✅ Feedback enviado ao professor." +

                        "</div>";
                    }

                    feedbackHTML +=
                      "</div>";
                  }

                  let botao = "";

                  if (
                    usuario.tipo ===
                    "PROFESSOR"
                  ) {

                    botao =

                      '<br><button class="btn btn-red" ' +

                      'onclick="excluirProva(\'' +

                      escapeHTML(prova.linha) +

                      '\')">' +

                      "🗑 Excluir" +

                      "</button>";
                  }

                  card.innerHTML =

                    "<h3>" +

                    "📝 " +

                    escapeHTML(prova.titulo) +

                    "</h3>" +

                    "<p><strong>Data:</strong> " +

                    escapeHTML(prova.data) +

                    "</p>" +

                    "<p><strong>Aluno:</strong> " +

                    escapeHTML(prova.aluno) +

                    "</p>" +

                    "<p><strong>RA:</strong> " +

                    escapeHTML(prova.alunoRA) +

                    "</p>" +

                    "<p><strong>Matéria:</strong> " +

                    escapeHTML(prova.materia) +

                    "</p>" +

                    "<p><strong>Professor:</strong> " +

                    escapeHTML(prova.professor) +

                    "</p>" +

                    nota +

                    "<p><strong>Descrição:</strong> " +

                    escapeHTML(
                      prova.descricao || "-"
                    ) +

                    "</p>" +

                    feedbackHTML +

                    botao;

                  grid.appendChild(card);
                }
              );
            }
          )

          .withFailureHandler(
            mostrarErro
          )

          .listarFeedbacks(token);
      }
    )

    .withFailureHandler(
      mostrarErro
    )

    .listarProvas(token);
}


function excluirProva(linha) {

  if (
    !confirm(
      "Excluir esta prova?"
    )
  ) {
    return;
  }

  google.script.run

    .withSuccessHandler(
      function(resultado) {

        alert(
          resultado.mensagem
        );

        carregarProvas();
      }
    )

    .withFailureHandler(
      mostrarErro
    )

    .excluirProva(
      token,
      Number(linha)
    );
}


// ==========================================================
// TRABALHOS
// ==========================================================

el("trabalhoSalvar").onclick =
function() {

  const dados = {

    data:
      el("trabalhoData")
        .value,

    titulo:
      el("trabalhoTitulo")
        .value.trim(),

    alunoRA:
      el("trabalhoAluno")
        .value,

    materia:
      el("trabalhoMateria")
        .value.trim(),

    nota:
      el("trabalhoNota")
        .value,

    descricao:
      el("trabalhoDescricao")
        .value.trim()
  };

  if (!dados.data) {

    alert("Informe a data.");

    return;
  }

  if (!dados.titulo) {

    alert("Informe o título.");

    return;
  }

  if (!dados.alunoRA) {

    alert("Selecione o aluno.");

    return;
  }

  if (!dados.materia) {

    alert("Informe a matéria.");

    return;
  }

  google.script.run

    .withSuccessHandler(
      function(resultado) {

        alert(
          resultado.mensagem
        );

        carregarTrabalhos();
      }
    )

    .withFailureHandler(
      mostrarErro
    )

    .cadastrarTrabalho(
      token,
      dados
    );
};


el("trabalhoAtualizar").onclick =
carregarTrabalhos;


function carregarTrabalhos() {

  google.script.run

    .withSuccessHandler(
      function(trabalhos) {

        el("trabalhoCount")
          .textContent =
          trabalhos.length;

        const grid =
          el("trabalhosGrid");

        grid.innerHTML = "";

        if (!trabalhos.length) {

          grid.innerHTML =
            '<div class="empty">' +

            "📂<br><br>" +

            "Nenhum trabalho encontrado." +

            "</div>";

          return;
        }

        google.script.run

          .withSuccessHandler(
            function(feedbacks) {

              trabalhos.forEach(
                function(trabalho) {

                  const card =
                    document.createElement("div");

                  card.className =
                    "item";

                  const feedback =
                    feedbacks.find(
                      function(f) {

                        return (
                          f.itemID ===
                          trabalho.id
                        );
                      }
                    );

                  const nota =
                    trabalho.nota !== ""

                      ? '<span class="note">⭐ Nota: ' +
                        escapeHTML(trabalho.nota) +
                        "</span>"

                      : '<span class="note">⏳ Nota não lançada</span>';

                  let feedbackHTML = "";

                  if (
                    usuario.tipo === "ALUNO"
                  ) {

                    feedbackHTML =

                      '<div class="feedback-box">' +

                      '<div class="feedback-title">' +

                      "💬 Seu feedback para o professor" +

                      "</div>" +

                      '<textarea ' +

                      'id="feedback_' +
                      escapeHTML(trabalho.id) +
                      '" ' +

                      'class="feedback-input" ' +

                      'placeholder="Escreva seu feedback...">' +

                      escapeHTML(
                        feedback
                          ? feedback.feedback
                          : ""
                      ) +

                      "</textarea>" +

                      '<button class="btn btn-primary" ' +

                      'onclick="enviarFeedback(\'' +
                      escapeHTML(trabalho.id) +
                      '\',\'TRABALHO\')">' +

                      "📨 Enviar feedback" +

                      "</button>" +

                      "</div>";
                  }

                  let botao = "";

                  if (
                    usuario.tipo ===
                    "PROFESSOR"
                  ) {

                    botao =

                      '<br><button class="btn btn-red" ' +

                      'onclick="excluirTrabalho(\'' +

                      escapeHTML(
                        trabalho.linha
                      ) +

                      '\')">' +

                      "🗑 Excluir" +

                      "</button>";
                  }

                  card.innerHTML =

                    "<h3>" +

                    "📂 " +

                    escapeHTML(
                      trabalho.titulo
                    ) +

                    "</h3>" +

                    "<p><strong>Data:</strong> " +

                    escapeHTML(
                      trabalho.data
                    ) +

                    "</p>" +

                    "<p><strong>Aluno:</strong> " +

                    escapeHTML(
                      trabalho.aluno
                    ) +

                    "</p>" +

                    "<p><strong>Matéria:</strong> " +

                    escapeHTML(
                      trabalho.materia
                    ) +

                    "</p>" +

                    "<p><strong>Professor:</strong> " +

                    escapeHTML(
                      trabalho.professor
                    ) +

                    "</p>" +

                    nota +

                    "<p><strong>Descrição:</strong> " +

                    escapeHTML(
                      trabalho.descricao || "-"
                    ) +

                    "</p>" +

                    feedbackHTML +

                    botao;

                  grid.appendChild(card);
                }
              );
            }
          )

          .withFailureHandler(
            mostrarErro
          )

          .listarFeedbacks(token);
      }
    )

    .withFailureHandler(
      mostrarErro
    )

    .listarTrabalhos(token);
}


function excluirTrabalho(linha) {

  if (
    !confirm(
      "Excluir este trabalho?"
    )
  ) {
    return;
  }

  google.script.run

    .withSuccessHandler(
      function(resultado) {

        alert(
          resultado.mensagem
        );

        carregarTrabalhos();
      }
    )

    .withFailureHandler(
      mostrarErro
    )

    .excluirTrabalho(
      token,
      Number(linha)
    );
}


// ==========================================================
// ENVIAR FEEDBACK
// ==========================================================

function enviarFeedback(itemID, tipo) {

  const campo =
    el("feedback_" + itemID);

  if (!campo) {
    alert(
      "Campo de feedback não encontrado."
    );
    return;
  }

  const feedback =
    campo.value.trim();

  if (!feedback) {

    alert(
      "Digite seu feedback antes de enviar."
    );

    campo.focus();

    return;
  }

  google.script.run

    .withSuccessHandler(
      function(resultado) {

        alert(
          resultado.mensagem
        );

        carregarProvas();

        carregarTrabalhos();

        carregarFeedbacks();
      }
    )

    .withFailureHandler(
      mostrarErro
    )

    .cadastrarFeedback(
      token,
      {
        itemID:
          itemID,

        tipo:
          tipo,

        feedback:
          feedback
      }
    );
}


// ==========================================================
// FEEDBACKS
// ==========================================================

el("feedbackAtualizar").onclick =
carregarFeedbacks;


function carregarFeedbacks() {

  google.script.run

    .withSuccessHandler(
      function(feedbacks) {

        const grid =
          el("feedbacksGrid");

        grid.innerHTML = "";

        if (!feedbacks.length) {

          grid.innerHTML =

            '<div class="empty">' +

            "💬<br><br>" +

            "Nenhum feedback encontrado." +

            "</div>";

          return;
        }

        feedbacks
          .slice()
          .reverse()
          .forEach(
            function(feedback) {

              const card =
                document.createElement("div");

              card.className =
                "item";

              card.innerHTML =

                "<h3>" +

                "💬 Feedback de " +

                escapeHTML(
                  feedback.aluno
                ) +

                "</h3>" +

                "<p><strong>RA:</strong> " +

                escapeHTML(
                  feedback.alunoRA
                ) +

                "</p>" +

                "<p><strong>Tipo:</strong> " +

                escapeHTML(
                  feedback.tipo
                ) +

                "</p>" +

                "<p><strong>Data:</strong> " +

                escapeHTML(
                  feedback.data
                ) +

                "</p>" +

                '<div class="feedback-enviado">' +

                escapeHTML(
                  feedback.feedback
                ) +

                "</div>";

              grid.appendChild(card);
            }
          );
      }
    )

    .withFailureHandler(
      mostrarErro
    )

    .listarFeedbacks(token);
}


// ==========================================================
// MOBILE
// ==========================================================

el("menuMobile").onclick =
function() {

  el("sidebar")
    .classList
    .toggle("open");
};


// ==========================================================
// LOGOUT
// ==========================================================

el("logoutBtn").onclick =
function() {

  if (
    !confirm(
      "Deseja realmente sair?"
    )
  ) {
    return;
  }

  google.script.run

    .withSuccessHandler(
      function() {

        localStorage.removeItem(
          "GE_TOKEN"
        );

        location.reload();
      }
    )

    .withFailureHandler(
      function() {

        localStorage.removeItem(
          "GE_TOKEN"
        );

        location.reload();
      }
    )

    .logout(token);
};


// ==========================================================
// ERRO
// ==========================================================

function mostrarErro(error) {

  alert(
    "❌ " +

    (
      error &&
      error.message
        ? error.message
        : String(error)
    )
  );
}


// ==========================================================
// SESSÃO SALVA
// ==========================================================

window.addEventListener(
  "load",
  function() {

    const salvo =
      localStorage.getItem(
        "GE_TOKEN"
      );

    if (!salvo) {
      return;
    }

    google.script.run

      .withSuccessHandler(
        function(usuarioRetornado) {

          token =
            salvo;

          usuario =
            usuarioRetornado;

          abrirApp();
        }
      )

      .withFailureHandler(
        function() {

          localStorage.removeItem(
            "GE_TOKEN"
          );
        }
      )

      .obterUsuarioLogado(
        salvo
      );
  }
);

</script>

</body>
</html>

<style>

* {
  box-sizing: border-box;
}

body {
  margin: 0;
  font-family: Arial, sans-serif;
  background: #f1f5f9;
  color: #172033;
}


/* LOGIN */

.tela {
  min-height: 100vh;

  display: flex;

  align-items: center;

  justify-content: center;

  padding: 20px;
}

.login-box {

  width: 100%;

  max-width: 420px;

  background: white;

  padding: 35px;

  border-radius: 15px;

  box-shadow:
    0 10px 30px
    rgba(0,0,0,.12);

  text-align: center;
}

.login-box h1 {
  font-size: 45px;
  margin-bottom: 5px;
}

.login-box h2 {
  margin-bottom: 5px;
}

.login-box p {
  color: #64748b;
}


input,
select,
textarea {

  width: 100%;

  padding: 12px;

  margin-top: 10px;

  border: 1px solid #cbd5e1;

  border-radius: 8px;

  font-family: Arial;
}

textarea {
  min-height: 90px;
  resize: vertical;
}


button {

  border: 0;

  border-radius: 8px;

  padding: 12px 18px;

  margin-top: 10px;

  background: #2563eb;

  color: white;

  font-weight: bold;

  cursor: pointer;
}

button:hover {
  background: #1d4ed8;
}

.btn-secundario {
  background: #64748b;
}

.btn-secundario:hover {
  background: #475569;
}


/* APP */

#app header {

  background: #172554;

  color: white;

  padding: 18px 25px;

  display: flex;

  justify-content: space-between;

  align-items: center;
}

#app header span {
  margin-left: 15px;
  color: #bfdbfe;
}

nav {

  background: white;

  padding: 10px;

  display: flex;

  gap: 8px;

  flex-wrap: wrap;

  box-shadow:
    0 2px 8px
    rgba(0,0,0,.08);
}

nav button {
  margin: 0;
}

main {
  padding: 30px;
  max-width: 1200px;
  margin: auto;
}


/* CARDS */

.cards {

  display: grid;

  grid-template-columns:
    repeat(
      auto-fit,
      minmax(180px, 1fr)
    );

  gap: 20px;

  margin-top: 25px;
}

.card {

  background: white;

  padding: 25px;

  border-radius: 12px;

  box-shadow:
    0 3px 12px
    rgba(0,0,0,.08);
}

.card h3 {
  color: #64748b;
}

.card strong {

  display: block;

  font-size: 35px;

  color: #2563eb;

  margin-top: 10px;
}


/* FORM */

.form-box {

  background: white;

  padding: 25px;

  margin-top: 25px;

  border-radius: 12px;

  box-shadow:
    0 3px 12px
    rgba(0,0,0,.08);
}


/* LISTAS */

.item {

  background: white;

  padding: 20px;

  margin-top: 12px;

  border-radius: 10px;

  box-shadow:
    0 2px 8px
    rgba(0,0,0,.06);
}

.item h3 {
  margin-top: 0;
}

.item small {
  color: #64748b;
}


/* CHAT */

.chat {

  background: white;

  border-radius: 12px;

  padding: 20px;

  margin-top: 15px;

  min-height: 300px;

  max-height: 450px;

  overflow-y: auto;
}

.mensagem {

  padding: 10px 14px;

  margin: 8px 0;

  border-radius: 10px;

  max-width: 75%;

  background: #e2e8f0;
}

.mensagem.eu {

  margin-left: auto;

  background: #2563eb;

  color: white;
}

.chat-envio {

  display: flex;

  gap: 10px;

  margin-top: 10px;
}

.chat-envio input {
  margin: 0;
}

.chat-envio button {
  margin: 0;
}


/* ESCONDER */

.escondido {
  display: none !important;
}


/* LOADING */

#loading {

  display: none;

  position: fixed;

  inset: 0;

  background: rgba(255,255,255,.8);

  z-index: 9999;

  align-items: center;

  justify-content: center;
}

#loading.ativo {
  display: flex;
}

.spinner {

  width: 45px;

  height: 45px;

  border: 5px solid #dbeafe;

  border-top-color: #2563eb;

  border-radius: 50%;

  animation:
    girar .8s linear infinite;
}

@keyframes girar {

  to {
    transform: rotate(360deg);
  }

}


/* MOBILE */

@media(max-width:700px) {

  main {
    padding: 15px;
  }

  #app header {
    flex-direction: column;
    gap: 10px;
    text-align: center;
  }

  nav {
    flex-direction: column;
  }

  nav button {
    width: 100%;
  }

  .chat-envio {
    flex-direction: column;
  }

}

</style>

<script>

let token = "";
let usuario = {};
let dados = {};


// ============================================================
// LOADING
// ============================================================

function loading(valor) {

  document
    .getElementById("loading")
    .classList.toggle(
      "ativo",
      valor
    );

}


// ============================================================
// LOGIN / CADASTRO
// ============================================================

function mostrarCadastro() {

  document
    .getElementById("loginTela")
    .classList.add("escondido");

  document
    .getElementById("cadastroTela")
    .classList.remove("escondido");

}


function mostrarLogin() {

  document
    .getElementById("cadastroTela")
    .classList.add("escondido");

  document
    .getElementById("loginTela")
    .classList.remove("escondido");

}


function mostrarCamposCadastro() {

  const tipo =
    document
      .getElementById("cadTipo")
      .value;

  document
    .getElementById("camposAluno")
    .classList.add("escondido");

  document
    .getElementById("camposProfessor")
    .classList.add("escondido");


  if (tipo === "ALUNO") {

    document
      .getElementById("camposAluno")
      .classList.remove("escondido");

  }

  if (tipo === "PROFESSOR") {

    document
      .getElementById("camposProfessor")
      .classList.remove("escondido");

  }

}


// ============================================================
// CADASTRAR
// ============================================================

function cadastrar() {

  const dadosCadastro = {

    nome:
      document
        .getElementById("cadNome")
        .value
        .trim(),

    email:
      document
        .getElementById("cadEmail")
        .value
        .trim(),

    senha:
      document
        .getElementById("cadSenha")
        .value,

    tipo:
      document
        .getElementById("cadTipo")
        .value,

    ra:
      document
        .getElementById("cadRA")
        .value,

    turma:
      document
        .getElementById("cadTurma")
        .value,

    nascimento:
      document
        .getElementById("cadNascimento")
        .value,

    materia:
      document
        .getElementById("cadMateria")
        .value

  };


  if (
    !dadosCadastro.nome ||
    !dadosCadastro.email ||
    !dadosCadastro.senha ||
    !dadosCadastro.tipo
  ) {

    alert(
      "Preencha nome, e-mail, senha e tipo."
    );

    return;

  }


  loading(true);


  google.script.run

    .withSuccessHandler(function(resposta) {

      loading(false);

      alert(
        resposta.mensagem
      );

      mostrarLogin();

    })

    .withFailureHandler(function(erro) {

      loading(false);

      alert(
        erro.message
      );

    })

    .cadastrarUsuario(
      dadosCadastro
    );

}


// ============================================================
// LOGIN
// ============================================================

function fazerLogin() {

  const email =
    document
      .getElementById("loginEmail")
      .value
      .trim();

  const senha =
    document
      .getElementById("loginSenha")
      .value;


  if (!email || !senha) {

    alert(
      "Digite seu e-mail e senha."
    );

    return;

  }


  loading(true);


  google.script.run

    .withSuccessHandler(function(resposta) {

      loading(false);

      token =
        resposta.token;

      usuario =
        resposta.usuario;


      abrirSistema();

    })

    .withFailureHandler(function(erro) {

      loading(false);

      alert(
        erro.message
      );

    })

    .login(
      email,
      senha
    );

}


// ============================================================
// ABRIR SISTEMA
// ============================================================

function abrirSistema() {

  document
    .getElementById("loginTela")
    .classList.add("escondido");

  document
    .getElementById("cadastroTela")
    .classList.add("escondido");

  document
    .getElementById("app")
    .classList.remove("escondido");


  document
    .getElementById("usuarioNome")
    .textContent =
      usuario.nome +
      " (" +
      usuario.tipo +
      ")";


  document
    .getElementById("nomeInicio")
    .textContent =
      usuario.nome;


  if (usuario.tipo === "ALUNO") {

    document
      .querySelectorAll(".professor-only")
      .forEach(function(elemento) {

        elemento.classList.add(
          "escondido"
        );

      });

    document
      .getElementById("menuResultados")
      .style.display = "block";

  }


  carregarDados();

}


// ============================================================
// DADOS
// ============================================================

function carregarDados() {

  loading(true);


  const funcao =
    usuario.tipo === "PROFESSOR"
      ? "dadosProfessor"
      : "dadosAluno";


  google.script.run

    .withSuccessHandler(function(resposta) {

      loading(false);

      dados = resposta;

      atualizarTudo();

    })

    .withFailureHandler(function(erro) {

      loading(false);

      alert(
        erro.message
      );

    })[funcao](token);

}


// ============================================================
// ATUALIZAR
// ============================================================

function atualizarTudo() {

  mostrarInicio();

  montarMaterias();

  montarProvas();

  montarTrabalhos();

  montarAgenda();

  montarResultados();

  montarContatos();

}


// ============================================================
// NAVEGAÇÃO
// ============================================================

function mostrarPagina(nome) {

  const paginas = [
    "inicio",
    "materias",
    "provas",
    "trabalhos",
    "agenda",
    "resultados",
    "chat"
  ];


  paginas.forEach(function(pagina) {

    const elemento =
      document.getElementById(
        "pagina-" + pagina
      );

    if (elemento) {

      elemento.classList.add(
        "escondido"
      );

    }

  });


  const pagina =
    document.getElementById(
      "pagina-" + nome
    );

  if (pagina) {

    pagina.classList.remove(
      "escondido"
    );

  }

}


// ============================================================
// INÍCIO
// ============================================================

function mostrarInicio() {

  const cards =
    document.getElementById("cards");

  cards.innerHTML = "";


  const lista = [

    ["📚 Matérias", dados.materias?.length || 0],

    ["📝 Provas", dados.provas?.length || 0],

    ["📄 Trabalhos", dados.trabalhos?.length || 0],

    ["📅 Agenda", dados.agenda?.length || 0]

  ];


  lista.forEach(function(item) {

    cards.innerHTML +=

      "<div class='card'>" +

      "<h3>" +
      item[0] +
      "</h3>" +

      "<strong>" +
      item[1] +
      "</strong>" +

      "</div>";

  });

}


// ============================================================
// MATÉRIAS
// ============================================================

function montarMaterias() {

  const elemento =
    document.getElementById(
      "materiasLista"
    );

  elemento.innerHTML = "";


  (dados.materias || [])
    .forEach(function(item) {

      elemento.innerHTML +=

        "<div class='item'>" +

        "<h3>" +
        escapar(item.Nome) +
        "</h3>" +

        "<small>" +
        "Professor: " +
        escapar(item.Professor) +
        "</small>" +

        "<p>" +
        escapar(item.Descricao) +
        "</p>" +

        "</div>";

    });

}


// ============================================================
// PROVAS
// ============================================================

function montarProvas() {

  const elemento =
    document.getElementById(
      "provasLista"
    );

  elemento.innerHTML = "";


  (dados.provas || [])
    .forEach(function(item) {

      elemento.innerHTML +=

        "<div class='item'>" +

        "<h3>" +
        escapar(item.Titulo) +
        "</h3>" +

        "<p>📅 " +
        escapar(item.Data) +
        "</p>" +

        "<p>📚 " +
        escapar(item.Materia) +
        "</p>" +

        "<p>👨‍🏫 " +
        escapar(item.Professor) +
        "</p>" +

        "<p>" +
        escapar(item.Descricao) +
        "</p>" +

        "</div>";

    });

}


// ============================================================
// TRABALHOS
// ============================================================

function montarTrabalhos() {

  const elemento =
    document.getElementById(
      "trabalhosLista"
    );

  elemento.innerHTML = "";


  const select =
    document.getElementById(
      "entregaTrabalho"
    );

  select.innerHTML =
    "<option value=''>Selecione o trabalho</option>";


  (dados.trabalhos || [])
    .forEach(function(item) {

      elemento.innerHTML +=

        "<div class='item'>" +

        "<h3>" +
        escapar(item.Titulo) +
        "</h3>" +

        "<p>📅 Entrega: " +
        escapar(item.DataEntrega) +
        "</p>" +

        "<p>📚 " +
        escapar(item.Materia) +
        "</p>" +

        "<p>" +
        escapar(item.Descricao) +
        "</p>" +

        "</div>";


      if (usuario.tipo === "ALUNO") {

        select.innerHTML +=

          "<option value='" +
          escapar(item.ID) +
          "'>" +

          escapar(item.Titulo) +

          "</option>";

      }

    });

}


// ============================================================
// AGENDA
// ============================================================

function montarAgenda() {

  const elemento =
    document.getElementById(
      "agendaLista"
    );

  elemento.innerHTML = "";


  (dados.agenda || [])
    .forEach(function(item) {

      elemento.innerHTML +=

        "<div class='item'>" +

        "<h3>" +
        escapar(item.Titulo) +
        "</h3>" +

        "<p>📅 " +
        escapar(item.Data) +
        "</p>" +

        "<p>⏰ " +
        escapar(item.Hora) +
        "</p>" +

        "<p>📌 " +
        escapar(item.Tipo) +
        "</p>" +

        "<p>" +
        escapar(item.Descricao) +
        "</p>" +

        "</div>";

    });

}


// ============================================================
// RESULTADOS
// ============================================================

function montarResultados() {

  const elemento =
    document.getElementById(
      "resultadosLista"
    );

  elemento.innerHTML = "";


  (dados.resultados || [])
    .forEach(function(item) {

      elemento.innerHTML +=

        "<div class='item'>" +

        "<h3>" +
        escapar(item.Prova) +
        "</h3>" +

        "<p>Aluno: " +
        escapar(item.Aluno) +
        "</p>" +

        "<p>Nota: <strong>" +
        escapar(item.Nota) +
        "</strong></p>" +

        "<p>" +
        escapar(item.Observacao) +
        "</p>" +

        "</div>";

    });

}


// ============================================================
// CONTATOS
// ============================================================

function montarContatos() {

  google.script.run

    .withSuccessHandler(function(contatos) {

      const select =
        document.getElementById(
          "contato"
        );

      select.innerHTML =
        "<option value=''>Selecione uma pessoa</option>";


      contatos.forEach(function(contato) {

        select.innerHTML +=

          "<option value='" +
          escapar(contato.ID) +
          "'>" +

          escapar(contato.Nome) +
          " - " +
          escapar(contato.Tipo) +

          "</option>";

      });

    })

    .listarContatos(token);

}


// ============================================================
// CHAT
// ============================================================

function carregarChat() {

  const outroID =
    document
      .getElementById("contato")
      .value;

  if (!outroID) return;


  google.script.run

    .withSuccessHandler(function(mensagens) {

      const area =
        document.getElementById(
          "chatMensagens"
        );

      area.innerHTML = "";


      mensagens.forEach(function(msg) {

        const minha =
          String(msg.RemetenteID) ===
          String(usuario.id);


        area.innerHTML +=

          "<div class='mensagem " +
          (minha ? "eu" : "") +
          "'>" +

          "<strong>" +
          escapar(msg.RemetenteNome) +
          "</strong><br>" +

          escapar(msg.Mensagem) +

          "</div>";

      });


      area.scrollTop =
        area.scrollHeight;

    })

    .listarChat(
      token,
      outroID
    );

}


function enviarChat() {

  const outroID =
    document
      .getElementById("contato")
      .value;

  const texto =
    document
      .getElementById("chatTexto")
      .value
      .trim();


  if (!outroID) {

    alert(
      "Selecione uma pessoa."
    );

    return;

  }


  if (!texto) {

    alert(
      "Digite uma mensagem."
    );

    return;

  }


  google.script.run

    .withSuccessHandler(function() {

      document
        .getElementById("chatTexto")
        .value = "";

      carregarChat();

    })

    .withFailureHandler(function(erro) {

      alert(
        erro.message
      );

    })

    .enviarMensagem(
      token,
      outroID,
      texto
    );

}


// ============================================================
// PROFESSOR - MATÉRIA
// ============================================================

function salvarMateria() {

  google.script.run

    .withSuccessHandler(function(msg) {

      alert(msg);

      carregarDados();

    })

    .withFailureHandler(function(erro) {

      alert(erro.message);

    })

    .cadastrarMateria(

      token,

      {

        nome:
          document
            .getElementById("materiaNome")
            .value,

        turma:
          document
            .getElementById("materiaTurma")
            .value,

        descricao:
          document
            .getElementById("materiaDescricao")
            .value

      }

    );

}


// ============================================================
// PROFESSOR - PROVA
// ============================================================

function salvarProva() {

  google.script.run

    .withSuccessHandler(function(msg) {

      alert(msg);

      carregarDados();

    })

    .withFailureHandler(function(erro) {

      alert(erro.message);

    })

    .cadastrarProva(

      token,

      {

        titulo:
          document
            .getElementById("provaTitulo")
            .value,

        data:
          document
            .getElementById("provaData")
            .value,

        turma:
          document
            .getElementById("provaTurma")
            .value,

        materia:
          document
            .getElementById("provaMateria")
            .value,

        valor:
          document
            .getElementById("provaValor")
            .value,

        descricao:
          document
            .getElementById("provaDescricao")
            .value

      }

    );

}


// ============================================================
// PROFESSOR - TRABALHO
// ============================================================

function salvarTrabalho() {

  google.script.run

    .withSuccessHandler(function(msg) {

      alert(msg);

      carregarDados();

    })

    .withFailureHandler(function(erro) {

      alert(erro.message);

    })

    .cadastrarTrabalho(

      token,

      {

        titulo:
          document
            .getElementById("trabalhoTitulo")
            .value,

        data:
          document
            .getElementById("trabalhoData")
            .value,

        turma:
          document
            .getElementById("trabalhoTurma")
            .value,

        materia:
          document
            .getElementById("trabalhoMateria")
            .value,

        descricao:
          document
            .getElementById("trabalhoDescricao")
            .value

      }

    );

}


// ============================================================
// PROFESSOR - AGENDA
// ============================================================

function salvarAgenda() {

  google.script.run

    .withSuccessHandler(function(msg) {

      alert(msg);

      carregarDados();

    })

    .withFailureHandler(function(erro) {

      alert(erro.message);

    })

    .cadastrarAgenda(

      token,

      {

        titulo:
          document
            .getElementById("agendaTitulo")
            .value,

        data:
          document
            .getElementById("agendaData")
            .value,

        hora:
          document
            .getElementById("agendaHora")
            .value,

        tipo:
          document
            .getElementById("agendaTipo")
            .value,

        turma:
          document
            .getElementById("agendaTurma")
            .value,

        materia:
          document
            .getElementById("agendaMateria")
            .value,

        descricao:
          document
            .getElementById("agendaDescricao")
            .value

      }

    );

}


// ============================================================
// PROFESSOR - RESULTADO
// ============================================================

function salvarResultado() {

  google.script.run

    .withSuccessHandler(function(msg) {

      alert(msg);

      carregarDados();

    })

    .withFailureHandler(function(erro) {

      alert(erro.message);

    })

    .cadastrarResultado(

      token,

      {

        prova:
          document
            .getElementById("resultadoProva")
            .value,

        aluno:
          document
            .getElementById("resultadoAluno")
            .value,

        turma:
          document
            .getElementById("resultadoTurma")
            .value,

        nota:
          document
            .getElementById("resultadoNota")
            .value,

        observacao:
          document
            .getElementById("resultadoObs")
            .value

      }

    );

}


// ============================================================
// ENTREGA
// ============================================================

function enviarEntrega() {

  const trabalhoID =
    document
      .getElementById("entregaTrabalho")
      .value;

  const mensagem =
    document
      .getElementById("entregaMensagem")
      .value;


  if (!trabalhoID) {

    alert(
      "Selecione um trabalho."
    );

    return;

  }


  google.script.run

    .withSuccessHandler(function(msg) {

      alert(msg);

      document
        .getElementById("entregaMensagem")
        .value = "";

    })

    .withFailureHandler(function(erro) {

      alert(erro.message);

    })

    .enviarTrabalho(

      token,

      {

        trabalhoID:
          trabalhoID,

        mensagem:
          mensagem

      }

    );

}


// ============================================================
// SAIR
// ============================================================

function sair() {

  google.script.run
    .logout(token);

  token = "";
  usuario = {};
  dados = {};

  document
    .getElementById("app")
    .classList.add("escondido");

  document
    .getElementById("loginTela")
    .classList.remove("escondido");

}


// ============================================================
// SEGURANÇA HTML
// ============================================================

function escapar(valor) {

  if (
    valor === null ||
    valor === undefined
  ) {
    return "";
  }

  return String(valor)

    .replace(/&/g, "&amp;")

    .replace(/</g, "&lt;")

    .replace(/>/g, "&gt;")

    .replace(/"/g, "&quot;")

    .replace(/'/g, "&#039;");

}

</script>