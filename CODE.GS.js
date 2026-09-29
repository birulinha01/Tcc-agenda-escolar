CODE.GS

// ============================================================
// GESTÃO ESCOLAR
// CODE.GS
// ============================================================

// ============================================================
// CONFIGURAÇÃO PRINCIPAL
// ============================================================

const ID_PLANILHA = "1-ZgtXS5Kz48NCMc-gcm7fcpnT2rsExNBi15eY0yvO-s";


// ============================================================
// ABAS
// ============================================================

const ABAS = {
  ALUNO: "ALUNO",
  PROFESSOR: "PROFESSOR",
  MATERIA: "MATERIA",
  PROVA: "PROVA",
  TRABALHO: "TRABALHO",
  MEDIA: "MEDIA",
  DUVIDA: "DUVIDA",
  CALENDARIO: "CALENDARIO"
};


// ============================================================
// SEGURANÇA
// ============================================================

const HASH_ITERACOES = 10000;


// ============================================================
// PLANILHA
// ============================================================

function getPlanilha() {

  return SpreadsheetApp.openById(
    ID_PLANILHA
  );

}


function getAba(nome) {

  const ss =
    getPlanilha();

  let aba =
    ss.getSheetByName(nome);

  if (!aba) {

    aba =
      ss.insertSheet(nome);

  }

  return aba;

}


// ============================================================
// CONFIGURAR PLANILHA
// ============================================================

function configurarPlanilha() {

  criarAba(
    ABAS.ALUNO,
    [
      "RA",
      "Nome",
      "Usuario",
      "SenhaHash",
      "Salt",
      "SerieAno",
      "Turma",
      "DataCadastro"
    ]
  );


  criarAba(
    ABAS.PROFESSOR,
    [
      "ID",
      "RA",
      "Nome",
      "Usuario",
      "SenhaHash",
      "Salt",
      "SerieAno",
      "Materia",
      "DataCadastro"
    ]
  );


  criarAba(
    ABAS.MATERIA,
    [
      "ID",
      "Nome",
      "Conteudo",
      "Links",
      "ProfessorRA",
      "Professor",
      "SerieAno",
      "DataCadastro"
    ]
  );


  criarAba(
    ABAS.PROVA,
    [
      "ID",
      "Data",
      "Titulo",
      "AlunoRA",
      "Aluno",
      "SerieAno",
      "Materia",
      "Nota",
      "Descricao",
      "ProfessorRA",
      "Professor",
      "DataCadastro"
    ]
  );


  criarAba(
    ABAS.TRABALHO,
    [
      "ID",
      "Data",
      "Titulo",
      "AlunoRA",
      "Aluno",
      "SerieAno",
      "Materia",
      "Nota",
      "Descricao",
      "ProfessorRA",
      "Professor",
      "DataCadastro"
    ]
  );


  // ==========================================================
  // NOVA ABA DE MÉDIAS
  // ==========================================================

  criarAba(
    ABAS.MEDIA,
    [
      "ID",
      "AlunoRA",
      "Aluno",
      "SerieAno",
      "Materia",
      "Media",
      "ProfessorRA",
      "Professor",
      "DataAtualizacao"
    ]
  );


  criarAba(
    ABAS.DUVIDA,
    [
      "ID",
      "Data",
      "AlunoRA",
      "Aluno",
      "ProfessorRA",
      "Professor",
      "Assunto",
      "Mensagem",
      "Resposta",
      "StatusAluno",
      "StatusProfessor",
      "DataResposta"
    ]
  );


  criarAba(
    ABAS.CALENDARIO,
    [
      "ID","Data","Tipo","Titulo","Descricao","SerieAno","Turma",
      "Materia","ProfessorRA","Professor","DataCadastro"
    ]
  );

  migrarSenhasAntigas();


  // ==========================================================
  // CORRIGIR FORMATAÇÃO DAS NOTAS
  // ==========================================================

  const abaProva =
    getAba(ABAS.PROVA);

  if (
    abaProva.getMaxRows() > 1
  ) {

    abaProva
      .getRange(
        2,
        8,
        abaProva.getMaxRows() - 1,
        1
      )
      .setNumberFormat("0.0");

  }


  const abaTrabalho =
    getAba(ABAS.TRABALHO);

  if (
    abaTrabalho.getMaxRows() > 1
  ) {

    abaTrabalho
      .getRange(
        2,
        8,
        abaTrabalho.getMaxRows() - 1,
        1
      )
      .setNumberFormat("0.0");

  }


  // ==========================================================
  // CORRIGIR FORMATAÇÃO DAS MÉDIAS
  // ==========================================================

  const abaMedia =
    getAba(ABAS.MEDIA);

  if (
    abaMedia.getMaxRows() > 1
  ) {

    abaMedia
      .getRange(
        2,
        6,
        abaMedia.getMaxRows() - 1,
        1
      )
      .setNumberFormat("0.0");

  }


  return "Planilha configurada com sucesso.";

}


// ============================================================
// CRIAR ABA
// ============================================================

function criarAba(
  nome,
  cabecalhos
) {

  const aba =
    getAba(nome);


  if (
    aba.getLastRow() === 0
  ) {

    aba
      .getRange(
        1,
        1,
        1,
        cabecalhos.length
      )
      .setValues([
        cabecalhos
      ]);


    aba
      .getRange(
        1,
        1,
        1,
        cabecalhos.length
      )
      .setFontWeight("bold");


    aba.setFrozenRows(1);

  }

}


// ============================================================
// MIGRAÇÃO DE SENHAS
// ============================================================

function migrarSenhasAntigas() {

  migrarSenhaAba(
    ABAS.ALUNO,
    4
  );

  migrarSenhaAba(
    ABAS.PROFESSOR,
    5
  );

}


function migrarSenhaAba(
  nomeAba,
  colunaSenha
) {

  const aba =
    getAba(nomeAba);


  if (
    aba.getLastRow() < 1
  ) {

    return;

  }


  const ultimaColuna =
    aba.getLastColumn();


  const cabecalhos =
    aba
      .getRange(
        1,
        1,
        1,
        ultimaColuna
      )
      .getValues()[0];


  const indiceSenha =
    cabecalhos.indexOf("Senha");


  const indiceHash =
    cabecalhos.indexOf("SenhaHash");


  const indiceSalt =
    cabecalhos.indexOf("Salt");


  if (
    indiceHash !== -1 &&
    indiceSalt !== -1
  ) {

    return;

  }


  if (
    indiceSenha !== -1
  ) {

    const colunaSenhaReal =
      indiceSenha + 1;


    aba
      .getRange(
        1,
        colunaSenhaReal
      )
      .setValue("SenhaHash");


    aba.insertColumnAfter(
      colunaSenhaReal
    );


    aba
      .getRange(
        1,
        colunaSenhaReal + 1
      )
      .setValue("Salt");


    const ultimaLinha =
      aba.getLastRow();


    if (
      ultimaLinha < 2
    ) {

      return;

    }


    const quantidade =
      ultimaLinha - 1;


    const senhas =
      aba
        .getRange(
          2,
          colunaSenhaReal,
          quantidade,
          1
        )
        .getValues();


    const hashes = [];
    const salts = [];


    for (
      let i = 0;
      i < senhas.length;
      i++
    ) {

      const senha =
        String(
          senhas[i][0] || ""
        );


      if (!senha) {

        hashes.push([""]);
        salts.push([""]);

        continue;

      }


      const salt =
        gerarSalt();


      const hash =
        gerarHashSenha(
          senha,
          salt
        );


      hashes.push([
        hash
      ]);


      salts.push([
        salt
      ]);

    }


    aba
      .getRange(
        2,
        colunaSenhaReal,
        quantidade,
        1
      )
      .setValues(
        hashes
      );


    aba
      .getRange(
        2,
        colunaSenhaReal + 1,
        quantidade,
        1
      )
      .setValues(
        salts
      );

  }

}


// ============================================================
// SHA-256
// ============================================================

function sha256(texto) {

  const digest =
    Utilities.computeDigest(
      Utilities.DigestAlgorithm.SHA_256,
      String(texto),
      Utilities.Charset.UTF_8
    );


  return digest
    .map(function(byte) {

      return (
        "0" +
        (byte & 0xFF)
          .toString(16)
      )
        .slice(-2);

    })
    .join("");

}


// ============================================================
// SALT
// ============================================================

function gerarSalt() {

  const base =
    Utilities.getUuid() +
    "_" +
    Utilities.getUuid() +
    "_" +
    new Date().getTime();


  return sha256(base);

}


// ============================================================
// HASH
// ============================================================

function gerarHashSenha(
  senha,
  salt
) {

  senha =
    String(
      senha || ""
    );


  salt =
    String(
      salt || ""
    );


  if (!senha) {

    throw new Error(
      "Senha não informada."
    );

  }


  let hash =
    sha256(
      salt +
      senha
    );


  for (
    let i = 1;
    i < HASH_ITERACOES;
    i++
  ) {

    hash =
      sha256(
        salt +
        hash
      );

  }


  return hash;

}


// ============================================================
// VERIFICAR SENHA
// ============================================================

function verificarSenha(
  senhaDigitada,
  hashSalvo,
  salt
) {

  if (
    !senhaDigitada ||
    !hashSalvo ||
    !salt
  ) {

    return false;

  }


  const hashDigitado =
    gerarHashSenha(
      senhaDigitada,
      salt
    );


  return (
    hashDigitado ===
    String(hashSalvo)
  );

}


// ============================================================
// WEB APP
// ============================================================

function doGet() {

  configurarPlanilha();


  return HtmlService
    .createTemplateFromFile("index")
    .evaluate()
    .setTitle("Gestão Escolar")
    .setXFrameOptionsMode(
      HtmlService.XFrameOptionsMode.ALLOWALL
    );

}


// ============================================================
// UTILIDADES
// ============================================================

function gerarID(prefixo) {

  return (
    prefixo +
    "_" +
    new Date().getTime() +
    "_" +
    Math.floor(
      Math.random() * 99999
    )
  );

}


function normalizar(valor) {

  return String(valor || "")
    .trim()
    .toLowerCase();

}


function formatarData(valor) {

  if (!valor) {

    return "";

  }


  if (
    Object.prototype.toString.call(valor) ===
    "[object Date]"
  ) {

    return Utilities.formatDate(
      valor,
      Session.getScriptTimeZone(),
      "dd/MM/yyyy"
    );

  }


  return String(valor);

}


function obterDadosAba(nome) {

  const aba =
    getAba(nome);


  const ultimaLinha =
    aba.getLastRow();


  const ultimaColuna =
    aba.getLastColumn();


  if (
    ultimaLinha < 2 ||
    ultimaColuna === 0
  ) {

    return [];

  }


  return aba
    .getRange(
      2,
      1,
      ultimaLinha - 1,
      ultimaColuna
    )
    .getValues();

}


// ============================================================
// SESSÕES
// ============================================================

function criarToken() {

  return Utilities
    .getUuid();

}


function salvarSessao(
  token,
  usuario
) {

  CacheService
    .getScriptCache()
    .put(
      "SESSAO_" + token,
      JSON.stringify(usuario),
      21600
    );

}


function obterSessao(token) {

  if (!token) {

    throw new Error(
      "Sessão inválida."
    );

  }


  const dados =
    CacheService
      .getScriptCache()
      .get(
        "SESSAO_" + token
      );


  if (!dados) {

    throw new Error(
      "Sua sessão expirou. Faça login novamente."
    );

  }


  return JSON.parse(
    dados
  );

}


function logout(token) {

  if (token) {

    CacheService
      .getScriptCache()
      .remove(
        "SESSAO_" + token
      );

  }


  return true;

}


// ============================================================
// CADASTRO
// ============================================================

function cadastrarUsuario(dados) {

  if (!dados) {

    return {
      sucesso: false,
      mensagem: "Dados não informados."
    };

  }


  const tipo =
    String(
      dados.tipo || ""
    )
      .toUpperCase();


  const ra =
    String(
      dados.ra || ""
    )
      .trim();


  const nome =
    String(
      dados.nome || ""
    )
      .trim();


  const usuario =
    String(
      dados.usuario || ""
    )
      .trim();


  const senha =
    String(
      dados.senha || ""
    );


  const serieAno =
    String(
      dados.serieAno || ""
    )
      .trim();


  const turma =
    String(
      dados.turma || ""
    )
      .trim();


  const materia =
    String(
      dados.materia || ""
    )
      .trim();


  if (!ra) {

    return {
      sucesso: false,
      mensagem: "Informe o RA/ID."
    };

  }


  if (!nome) {

    return {
      sucesso: false,
      mensagem: "Informe o nome."
    };

  }


  if (!usuario) {

    return {
      sucesso: false,
      mensagem: "Informe o usuário."
    };

  }


  if (!senha) {

    return {
      sucesso: false,
      mensagem: "Informe a senha."
    };

  }


  if (
    senha.length < 6
  ) {

    return {
      sucesso: false,
      mensagem:
        "A senha precisa ter pelo menos 6 caracteres."
    };

  }


  if (!serieAno) {

    return {
      sucesso: false,
      mensagem:
        "Informe a Série/Ano."
    };

  }


  if (
    tipo === "ALUNO" &&
    !turma
  ) {

    return {
      sucesso: false,
      mensagem:
        "Informe a Turma."
    };

  }


  if (
    tipo === "PROFESSOR" &&
    !materia
  ) {

    return {
      sucesso: false,
      mensagem:
        "Informe a matéria."
    };

  }


  // ==========================================================
  // ALUNO
  // ==========================================================

  if (
    tipo === "ALUNO"
  ) {

    const aba =
      getAba(
        ABAS.ALUNO
      );


    const dadosAba =
      obterDadosAba(
        ABAS.ALUNO
      );


    for (
      let i = 0;
      i < dadosAba.length;
      i++
    ) {

      const linha =
        dadosAba[i];


      if (
        normalizar(linha[0]) ===
        normalizar(ra)
      ) {

        return {
          sucesso: false,
          mensagem:
            "Este RA já está cadastrado."
        };

      }


      if (
        normalizar(linha[2]) ===
        normalizar(usuario)
      ) {

        return {
          sucesso: false,
          mensagem:
            "Este usuário já está cadastrado."
        };

      }

    }


    const salt =
      gerarSalt();


    const senhaHash =
      gerarHashSenha(
        senha,
        salt
      );


    aba.appendRow([
      ra,
      nome,
      usuario,
      senhaHash,
      salt,
      serieAno,
      turma,
      new Date()
    ]);


    return {
      sucesso: true,
      mensagem:
        "Aluno cadastrado com sucesso!"
    };

  }


  // ==========================================================
  // PROFESSOR
  // ==========================================================

  if (
    tipo === "PROFESSOR"
  ) {

    const aba =
      getAba(
        ABAS.PROFESSOR
      );


    const dadosAba =
      obterDadosAba(
        ABAS.PROFESSOR
      );


    for (
      let i = 0;
      i < dadosAba.length;
      i++
    ) {

      const linha =
        dadosAba[i];


      if (
        normalizar(linha[1]) ===
        normalizar(ra)
      ) {

        return {
          sucesso: false,
          mensagem:
            "Este ID já está cadastrado."
        };

      }


      if (
        normalizar(linha[3]) ===
        normalizar(usuario)
      ) {

        return {
          sucesso: false,
          mensagem:
            "Este usuário já está cadastrado."
        };

      }

    }


    const salt =
      gerarSalt();


    const senhaHash =
      gerarHashSenha(
        senha,
        salt
      );


    aba.appendRow([
      gerarID("PROF"),
      ra,
      nome,
      usuario,
      senhaHash,
      salt,
      serieAno,
      materia,
      new Date()
    ]);


    return {
      sucesso: true,
      mensagem:
        "Professor cadastrado com sucesso!"
    };

  }


  return {
    sucesso: false,
    mensagem:
      "Tipo de usuário inválido."
  };

}


// ============================================================
// LOGIN
// ============================================================

function login(dados) {

  if (!dados) {

    return {
      sucesso: false,
      mensagem:
        "Informe os dados de login."
    };

  }


  const usuarioLogin =
    normalizar(
      dados.usuario
    );


  const senha =
    String(
      dados.senha || ""
    );


  const tipo =
    String(
      dados.tipo || ""
    )
      .toUpperCase();


  if (
    !usuarioLogin ||
    !senha
  ) {

    return {
      sucesso: false,
      mensagem:
        "Informe usuário e senha."
    };

  }


  // ==========================================================
  // ALUNO
  // ==========================================================

  if (
    tipo === "ALUNO"
  ) {

    const alunos =
      obterDadosAba(
        ABAS.ALUNO
      );


    for (
      let i = 0;
      i < alunos.length;
      i++
    ) {

      const linha =
        alunos[i];


      const usuarioPlanilha =
        normalizar(
          linha[2]
        );


      const senhaHash =
        String(
          linha[3] || ""
        );


      const salt =
        String(
          linha[4] || ""
        );


      if (
        usuarioPlanilha ===
        usuarioLogin
      ) {

        if (
          !verificarSenha(
            senha,
            senhaHash,
            salt
          )
        ) {

          return {
            sucesso: false,
            mensagem:
              "Usuário ou senha incorretos."
          };

        }


        const usuario = {

          tipo: "ALUNO",

          ra:
            String(linha[0]),

          nome:
            String(linha[1]),

          usuario:
            String(linha[2]),

          serieAno:
            String(linha[5]),

          turma:
            String(
              linha[6] || ""
            )

        };


        const token =
          criarToken();


        salvarSessao(
          token,
          usuario
        );


        return {

          sucesso: true,

          mensagem:
            "Login realizado.",

          token:
            token,

          usuario:
            usuario

        };

      }

    }

  }


  // ==========================================================
  // PROFESSOR
  // ==========================================================

  if (
    tipo === "PROFESSOR"
  ) {

    const professores =
      obterDadosAba(
        ABAS.PROFESSOR
      );


    for (
      let i = 0;
      i < professores.length;
      i++
    ) {

      const linha =
        professores[i];


      const usuarioPlanilha =
        normalizar(
          linha[3]
        );


      const senhaHash =
        String(
          linha[4] || ""
        );


      const salt =
        String(
          linha[5] || ""
        );


      if (
        usuarioPlanilha ===
        usuarioLogin
      ) {

        if (
          !verificarSenha(
            senha,
            senhaHash,
            salt
          )
        ) {

          return {
            sucesso: false,
            mensagem:
              "Usuário ou senha incorretos."
          };

        }


        const usuario = {

          tipo: "PROFESSOR",

          id:
            String(linha[0]),

          ra:
            String(linha[1]),

          nome:
            String(linha[2]),

          usuario:
            String(linha[3]),

          serieAno:
            String(linha[6]),

          materia:
            String(linha[7])

        };


        const token =
          criarToken();


        salvarSessao(
          token,
          usuario
        );


        return {

          sucesso: true,

          mensagem:
            "Login realizado.",

          token:
            token,

          usuario:
            usuario

        };

      }

    }

  }


  return {
    sucesso: false,
    mensagem:
      "Usuário, senha ou tipo de conta incorreto."
  };

}


// ============================================================
// USUÁRIO LOGADO
// ============================================================

function obterUsuarioLogado(
  token
) {

  return obterSessao(
    token
  );

}


// ============================================================
// MATÉRIAS
// ============================================================

function cadastrarMateria(
  token,
  dados
) {

  const professor =
    obterSessao(token);


  if (
    professor.tipo !==
    "PROFESSOR"
  ) {

    throw new Error(
      "Apenas professores podem cadastrar matérias."
    );

  }


  if (
    !dados.nome
  ) {

    throw new Error(
      "Informe o nome da matéria."
    );

  }


  getAba(
    ABAS.MATERIA
  )
    .appendRow([

      gerarID("MAT"),

      dados.nome,

      dados.conteudo || "",

      dados.links || "",

      professor.ra,

      professor.nome,

      professor.serieAno,

      new Date()

    ]);


  return {

    sucesso: true,

    mensagem:
      "Matéria publicada com sucesso!"

  };

}


function listarMaterias(
  token
) {

  const usuario =
    obterSessao(token);


  const dados =
    obterDadosAba(
      ABAS.MATERIA
    );


  return dados

    .map(function(linha, index) {

      return {

        linha:
          index + 2,

        id:
          String(linha[0]),

        nome:
          String(linha[1]),

        conteudo:
          String(
            linha[2] || ""
          ),

        links:
          String(
            linha[3] || ""
          ),

        professorRA:
          String(linha[4]),

        professor:
          String(linha[5]),

        serieAno:
          String(linha[6]),

        data:
          formatarData(
            linha[7]
          )

      };

    })

    .filter(function(materia) {

      return (
        normalizar(
          materia.serieAno
        ) ===
        normalizar(
          usuario.serieAno
        )
      );

    });

}


// ============================================================
// ALUNOS
// ============================================================

function listarAlunos(
  token
) {

  const professor =
    obterSessao(token);


  if (
    professor.tipo !==
    "PROFESSOR"
  ) {

    throw new Error(
      "Apenas professores podem acessar esta lista."
    );

  }


  const alunos =
    obterDadosAba(
      ABAS.ALUNO
    );


  return alunos

    .map(function(linha) {

      return {

        ra:
          String(linha[0]),

        nome:
          String(linha[1]),

        usuario:
          String(linha[2]),

        serieAno:
          String(linha[5]),

        turma:
          String(
            linha[6] || ""
          )

      };

    })

    .filter(function(aluno) {

      return (
        normalizar(
          aluno.serieAno
        ) ===
        normalizar(
          professor.serieAno
        )
      );

    });

}


// ============================================================
// PROVAS
// ============================================================
// O professor cadastra a prova.
// A prova é criada para todos os alunos da Série/Ano.
// A nota é lançada posteriormente na área "Lançar Notas".
// ============================================================

function cadastrarProva(
  token,
  dados
) {

  const professor =
    obterSessao(token);


  if (
    professor.tipo !==
    "PROFESSOR"
  ) {

    throw new Error(
      "Apenas professores podem cadastrar provas."
    );

  }


  if (!dados) {

    throw new Error(
      "Dados da prova não informados."
    );

  }


  const data =
    String(
      dados.data || ""
    )
      .trim();


  const titulo =
    String(
      dados.titulo || ""
    )
      .trim();


  const materia =
    String(
      dados.materia || ""
    )
      .trim();


  const descricao =
    String(
      dados.descricao || ""
    )
      .trim();


  if (!data) {

    throw new Error(
      "Informe a data da prova."
    );

  }


  if (!titulo) {

    throw new Error(
      "Informe o título da prova."
    );

  }


  if (!materia) {

    throw new Error(
      "Informe a matéria."
    );

  }


  const alunos =
    obterDadosAba(
      ABAS.ALUNO
    );


  const alunosDestino =
    alunos.filter(function(linha) {

      return (
        normalizar(linha[5]) ===
        normalizar(professor.serieAno)
      );

    });


  if (
    !alunosDestino.length
  ) {

    throw new Error(
      "Nenhum aluno encontrado nesta Série/Ano."
    );

  }


  const idProva =
    gerarID("PROVA");


  const aba =
    getAba(
      ABAS.PROVA
    );


  const linhas = [];


  alunosDestino.forEach(function(aluno) {

    linhas.push([

      idProva,

      data,

      titulo,

      String(aluno[0]),

      String(aluno[1]),

      String(aluno[5]),

      materia,

      "",

      descricao,

      professor.ra,

      professor.nome,

      new Date()

    ]);

  });


  const primeiraLinha =
    aba.getLastRow() + 1;


  aba
    .getRange(
      primeiraLinha,
      1,
      linhas.length,
      12
    )
    .setValues(
      linhas
    );


  aba
    .getRange(
      primeiraLinha,
      8,
      linhas.length,
      1
    )
    .setNumberFormat(
      "0.0"
    );

  criarEventosCalendarioParaPublicacao(
    professor,
    "PROVA",
    idProva,
    data,
    titulo,
    descricao,
    materia,
    alunosDestino
  );


  return {

    sucesso: true,

    id:
      idProva,

    quantidade:
      alunosDestino.length,

    mensagem:
      "Prova cadastrada para " +
      alunosDestino.length +
      " aluno(s) da Série/Ano."

  };

}


// ============================================================
// LISTAR PROVAS
// ============================================================

function listarProvas(
  token
) {

  const usuario =
    obterSessao(token);


  const dados =
    obterDadosAba(
      ABAS.PROVA
    );


  return dados

    .map(function(linha, index) {

      let nota =
        linha[7];


      if (
        nota === null ||
        nota === undefined ||
        nota === ""
      ) {

        nota = "";

      }

      else if (
        Object.prototype.toString.call(nota) ===
        "[object Date]"
      ) {

        nota = "";

      }

      else {

        const texto =
          String(nota)
            .trim()
            .replace(",", ".");


        const numero =
          Number(texto);


        if (
          !isNaN(numero)
        ) {

          nota =
            numero
              .toFixed(1)
              .replace(".", ",");

        } else {

          nota =
            String(nota);

        }

      }


      return {

        linha:
          index + 2,

        id:
          String(
            linha[0] || ""
          ),

        data:
          formatarData(
            linha[1]
          ),

        titulo:
          String(
            linha[2] || ""
          ),

        alunoRA:
          String(
            linha[3] || ""
          ),

        aluno:
          String(
            linha[4] || ""
          ),

        serieAno:
          String(
            linha[5] || ""
          ),

        materia:
          String(
            linha[6] || ""
          ),

        nota:
          nota,

        descricao:
          String(
            linha[8] || ""
          ),

        professorRA:
          String(
            linha[9] || ""
          ),

        professor:
          String(
            linha[10] || ""
          )

      };

    })

    .filter(function(prova) {

      if (
        usuario.tipo ===
        "ALUNO"
      ) {

        return (
          String(
            prova.alunoRA
          ) ===
          String(
            usuario.ra
          )
        );

      }


      return (
        String(
          prova.professorRA
        ) ===
        String(
          usuario.ra
        )
      );

    });

}


// ============================================================
// PROVAS - LISTAR PARA LANÇAMENTO DE NOTAS
// ============================================================

function listarProvasParaNotas(
  token
) {

  const professor =
    obterSessao(token);


  if (
    professor.tipo !==
    "PROFESSOR"
  ) {

    throw new Error(
      "Apenas professores podem acessar as notas."
    );

  }


  const dados =
    obterDadosAba(
      ABAS.PROVA
    );


  const provas = [];
  const idsEncontrados = {};


  dados.forEach(function(linha) {

    const id =
      String(
        linha[0] || ""
      )
        .trim();


    const professorRA =
      String(
        linha[9] || ""
      )
        .trim();


    const serieAno =
      String(
        linha[5] || ""
      )
        .trim();


    if (!id) {

      return;

    }


    if (
      professorRA !==
      String(professor.ra)
    ) {

      return;

    }


    if (
      normalizar(serieAno) !==
      normalizar(professor.serieAno)
    ) {

      return;

    }


    if (
      idsEncontrados[id]
    ) {

      return;

    }


    idsEncontrados[id] = true;


    provas.push({

      id:
        id,

      data:
        formatarData(
          linha[1]
        ),

      titulo:
        String(
          linha[2] || ""
        ),

      serieAno:
        serieAno,

      materia:
        String(
          linha[6] || ""
        ),

      descricao:
        String(
          linha[8] || ""
        ),

      professorRA:
        professorRA,

      professor:
        String(
          linha[10] || ""
        )

    });

  });


  return provas.reverse();

}


// ============================================================
// PROVAS - LISTAR TODOS OS ALUNOS DE UMA PROVA
// ============================================================

function listarAlunosDaProva(
  token,
  idProva
) {

  const professor =
    obterSessao(token);


  if (
    professor.tipo !==
    "PROFESSOR"
  ) {

    throw new Error(
      "Apenas professores podem acessar esta função."
    );

  }


  idProva =
    String(
      idProva || ""
    )
      .trim();


  if (!idProva) {

    throw new Error(
      "Informe a prova."
    );

  }


  const dados =
    obterDadosAba(
      ABAS.PROVA
    );


  const alunos = [];


  for (
    let i = 0;
    i < dados.length;
    i++
  ) {

    const linha =
      dados[i];


    const id =
      String(
        linha[0] || ""
      )
        .trim();


    if (
      id !== idProva
    ) {

      continue;

    }


    if (
      String(
        linha[9] || ""
      ) !==
      String(
        professor.ra
      )
    ) {

      throw new Error(
        "Você não pode acessar esta prova."
      );

    }


    let nota =
      linha[7];


    if (
      nota === null ||
      nota === undefined ||
      nota === ""
    ) {

      nota = "";

    }

    else if (
      Object.prototype.toString.call(nota) ===
      "[object Date]"
    ) {

      nota = "";

    }

    else {

      const numero =
        Number(
          String(nota)
            .trim()
            .replace(",", ".")
        );


      if (
        !isNaN(numero)
      ) {

        nota =
          numero
            .toFixed(1)
            .replace(".", ",");

      } else {

        nota =
          String(nota);

      }

    }


    alunos.push({

      linha:
        i + 2,

      ra:
        String(
          linha[3] || ""
        ),

      nome:
        String(
          linha[4] || ""
        ),

      serieAno:
        String(
          linha[5] || ""
        ),

      materia:
        String(
          linha[6] || ""
        ),

      nota:
        nota

    });

  }


  if (
    alunos.length === 0
  ) {

    throw new Error(
      "Nenhum aluno encontrado nesta prova."
    );

  }


  return alunos;

}


// ============================================================
// PROVAS - LANÇAR NOTA INDIVIDUAL
// ============================================================

function lancarNotaProva(
  token,
  idProva,
  alunoRA,
  nota
) {

  const professor =
    obterSessao(token);


  if (
    professor.tipo !==
    "PROFESSOR"
  ) {

    throw new Error(
      "Apenas professores podem lançar notas."
    );

  }


  idProva =
    String(
      idProva || ""
    )
      .trim();


  alunoRA =
    String(
      alunoRA || ""
    )
      .trim();


  if (!idProva) {

    throw new Error(
      "Informe a prova."
    );

  }


  if (!alunoRA) {

    throw new Error(
      "Informe o aluno."
    );

  }


  let notaFinal = "";


  if (
    nota !== undefined &&
    nota !== null
  ) {

    const texto =
      String(
        nota
      )
        .trim()
        .replace(",", ".");


    if (
      texto !== ""
    ) {

      const numero =
        Number(texto);


      if (
        isNaN(numero)
      ) {

        throw new Error(
          "A nota deve ser um número válido."
        );

      }


      if (
        numero < 0 ||
        numero > 10
      ) {

        throw new Error(
          "A nota deve estar entre 0 e 10."
        );

      }


      notaFinal =
        Math.round(
          numero * 10
        ) / 10;

    }

  }


  const aba =
    getAba(
      ABAS.PROVA
    );


  const dados =
    obterDadosAba(
      ABAS.PROVA
    );


  for (
    let i = 0;
    i < dados.length;
    i++
  ) {

    const linha =
      dados[i];


    const idLinha =
      String(
        linha[0] || ""
      )
        .trim();


    const alunoLinha =
      String(
        linha[3] || ""
      )
        .trim();


    const professorLinha =
      String(
        linha[9] || ""
      )
        .trim();


    if (
      idLinha === idProva &&
      alunoLinha === alunoRA
    ) {

      if (
        professorLinha !==
        String(
          professor.ra
        )
      ) {

        throw new Error(
          "Você não pode alterar esta nota."
        );

      }


      const numeroLinha =
        i + 2;


      const celulaNota =
        aba
          .getRange(
            numeroLinha,
            8
          );


      celulaNota
        .setNumberFormat(
          "0.0"
        );


      if (
        notaFinal === ""
      ) {

        celulaNota
          .clearContent();

      } else {

        celulaNota
          .setValue(
            notaFinal
          );

      }


      return {

        sucesso: true,

        mensagem:
          "Nota lançada com sucesso.",

        alunoRA:
          alunoRA,

        nota:
          notaFinal

      };

    }

  }


  throw new Error(
    "Aluno não encontrado nesta prova."
  );

}


// ============================================================
// PROVAS - SALVAR TODAS AS NOTAS
// ============================================================

function salvarNotasProva(
  token,
  idProva,
  notas
) {

  const professor =
    obterSessao(token);


  if (
    professor.tipo !==
    "PROFESSOR"
  ) {

    throw new Error(
      "Apenas professores podem lançar notas."
    );

  }


  idProva =
    String(
      idProva || ""
    )
      .trim();


  if (!idProva) {

    throw new Error(
      "Informe a prova."
    );

  }


  if (
    !notas ||
    !Array.isArray(notas)
  ) {

    throw new Error(
      "Nenhuma nota foi informada."
    );

  }


  const dadosProvas =
    obterDadosAba(
      ABAS.PROVA
    );


  let provaEncontrada =
    false;


  for (
    let i = 0;
    i < dadosProvas.length;
    i++
  ) {

    if (
      String(
        dadosProvas[i][0] || ""
      ) === idProva
    ) {

      if (
        String(
          dadosProvas[i][9] || ""
        ) !==
        String(
          professor.ra
        )
      ) {

        throw new Error(
          "Você não pode alterar esta prova."
        );

      }


      provaEncontrada = true;

      break;

    }

  }


  if (
    !provaEncontrada
  ) {

    throw new Error(
      "Prova não encontrada."
    );

  }


  let quantidade =
    0;


  notas.forEach(function(item) {

    if (!item) {

      return;

    }


    const alunoRA =
      String(
        item.alunoRA || ""
      )
        .trim();


    const nota =
      item.nota;


    if (!alunoRA) {

      return;

    }


    if (
      nota === undefined ||
      nota === null ||
      String(nota).trim() === ""
    ) {

      lancarNotaProva(
        token,
        idProva,
        alunoRA,
        ""
      );


      quantidade++;

      return;

    }


    const numero =
      Number(
        String(nota)
          .trim()
          .replace(",", ".")
      );


    if (
      isNaN(numero)
    ) {

      throw new Error(
        "A nota do aluno " +
        alunoRA +
        " é inválida."
      );

    }


    if (
      numero < 0 ||
      numero > 10
    ) {

      throw new Error(
        "A nota do aluno " +
        alunoRA +
        " deve estar entre 0 e 10."
      );

    }


    lancarNotaProva(
      token,
      idProva,
      alunoRA,
      numero
    );


    quantidade++;

  });


  return {

    sucesso: true,

    quantidade:
      quantidade,

    mensagem:
      quantidade +
      " nota(s) salva(s) com sucesso."

  };

}


// ============================================================
// MÉDIAS
// ============================================================
// SOMENTE PROFESSOR PODE LANÇAR/ALTERAR.
// ALUNO SOMENTE VISUALIZA SUA PRÓPRIA MÉDIA.
// ============================================================


// ============================================================
// LISTAR MATÉRIAS PARA LANÇAMENTO DE MÉDIAS
// ============================================================

function listarMateriasParaMedias(
  token
) {

  const professor =
    obterSessao(token);


  if (
    professor.tipo !==
    "PROFESSOR"
  ) {

    throw new Error(
      "Apenas professores podem acessar as médias."
    );

  }


  const alunos =
    obterDadosAba(
      ABAS.ALUNO
    );


  const materias = [];


  // ----------------------------------------------------------
  // PRIMEIRO: matéria do cadastro do professor
  // ----------------------------------------------------------

  if (
    professor.materia
  ) {

    materias.push(
      String(
        professor.materia
      ).trim()
    );

  }


  // ----------------------------------------------------------
  // SEGUNDO: matérias cadastradas pelo professor
  // ----------------------------------------------------------

  const dadosMaterias =
    obterDadosAba(
      ABAS.MATERIA
    );


  dadosMaterias.forEach(function(linha) {

    const professorRA =
      String(
        linha[4] || ""
      ).trim();


    const serieAno =
      String(
        linha[6] || ""
      ).trim();


    const nomeMateria =
      String(
        linha[1] || ""
      ).trim();


    if (
      professorRA ===
      String(professor.ra)
      &&
      normalizar(serieAno) ===
      normalizar(professor.serieAno)
      &&
      nomeMateria
    ) {

      if (
        !materias.some(function(item) {

          return (
            normalizar(item) ===
            normalizar(nomeMateria)
          );

        })
      ) {

        materias.push(
          nomeMateria
        );

      }

    }

  });


  // ----------------------------------------------------------
  // TERCEIRO: matérias existentes nas provas do professor
  // ----------------------------------------------------------

  const provas =
    obterDadosAba(
      ABAS.PROVA
    );


  provas.forEach(function(linha) {

    const professorRA =
      String(
        linha[9] || ""
      ).trim();


    const serieAno =
      String(
        linha[5] || ""
      ).trim();


    const materia =
      String(
        linha[6] || ""
      ).trim();


    if (
      professorRA ===
      String(professor.ra)
      &&
      normalizar(serieAno) ===
      normalizar(professor.serieAno)
      &&
      materia
    ) {

      if (
        !materias.some(function(item) {

          return (
            normalizar(item) ===
            normalizar(materia)
          );

        })
      ) {

        materias.push(
          materia
        );

      }

    }

  });


  // Evita variável não utilizada em alguns ambientes
  alunos.length;


  return materias;

}


// ============================================================
// LISTAR ALUNOS PARA LANÇAMENTO DE MÉDIAS
// ============================================================

function listarAlunosParaMedias(
  token,
  materia
) {

  const professor =
    obterSessao(token);


  if (
    professor.tipo !==
    "PROFESSOR"
  ) {

    throw new Error(
      "Apenas professores podem acessar as médias."
    );

  }


  materia =
    String(
      materia || ""
    )
      .trim();


  if (!materia) {

    throw new Error(
      "Informe a matéria."
    );

  }


  const alunos =
    obterDadosAba(
      ABAS.ALUNO
    );


  const medias =
    obterDadosAba(
      ABAS.MEDIA
    );


  return alunos

    .filter(function(linha) {

      return (
        normalizar(linha[5]) ===
        normalizar(professor.serieAno)
      );

    })

    .map(function(linha) {

      const alunoRA =
        String(
          linha[0] || ""
        ).trim();


      let media = "";


      // --------------------------------------------------------
      // Procurar média existente
      // --------------------------------------------------------

      for (
        let i = 0;
        i < medias.length;
        i++
      ) {

        const mediaRA =
          String(
            medias[i][1] || ""
          ).trim();


        const mediaMateria =
          String(
            medias[i][4] || ""
          ).trim();


        const mediaProfessor =
          String(
            medias[i][6] || ""
          ).trim();


        if (
          mediaRA === alunoRA
          &&
          normalizar(mediaMateria) ===
          normalizar(materia)
          &&
          mediaProfessor ===
          String(professor.ra)
        ) {

          const valor =
            medias[i][5];


          if (
            valor !== null &&
            valor !== undefined &&
            valor !== ""
          ) {

            if (
              Object.prototype.toString.call(valor) !==
              "[object Date]"
            ) {

              const numero =
                Number(
                  String(valor)
                    .trim()
                    .replace(",", ".")
                );


              if (
                !isNaN(numero)
              ) {

                media =
                  numero
                    .toFixed(1)
                    .replace(".", ",");

              } else {

                media =
                  String(valor);

              }

            }

          }

          break;

        }

      }


      return {

        ra:
          alunoRA,

        nome:
          String(
            linha[1] || ""
          ),

        serieAno:
          String(
            linha[5] || ""
          ),

        turma:
          String(
            linha[6] || ""
          ),

        materia:
          materia,

        media:
          media

      };

    });

}


// ============================================================
// SALVAR MÉDIA DE UM ALUNO
// ============================================================

function salvarMedia(
  token,
  alunoRA,
  materia,
  media
) {

  const professor =
    obterSessao(token);


  if (
    professor.tipo !==
    "PROFESSOR"
  ) {

    throw new Error(
      "Somente professores podem lançar médias."
    );

  }


  alunoRA =
    String(
      alunoRA || ""
    )
      .trim();


  materia =
    String(
      materia || ""
    )
      .trim();


  if (!alunoRA) {

    throw new Error(
      "Informe o aluno."
    );

  }


  if (!materia) {

    throw new Error(
      "Informe a matéria."
    );

  }


  // ==========================================================
  // VALIDAR ALUNO
  // ==========================================================

  const alunos =
    obterDadosAba(
      ABAS.ALUNO
    );


  let alunoEncontrado =
    null;


  for (
    let i = 0;
    i < alunos.length;
    i++
  ) {

    if (
      String(
        alunos[i][0] || ""
      ).trim() === alunoRA
    ) {

      alunoEncontrado =
        alunos[i];

      break;

    }

  }


  if (!alunoEncontrado) {

    throw new Error(
      "Aluno não encontrado."
    );

  }


  // ==========================================================
  // GARANTIR QUE O ALUNO É DA SÉRIE DO PROFESSOR
  // ==========================================================

  if (
    normalizar(alunoEncontrado[5]) !==
    normalizar(professor.serieAno)
  ) {

    throw new Error(
      "Você não pode lançar média para este aluno."
    );

  }


  // ==========================================================
  // TRATAMENTO DA MÉDIA
  // ==========================================================

  let mediaFinal = "";


  if (
    media !== undefined &&
    media !== null
  ) {

    const texto =
      String(
        media
      )
        .trim()
        .replace(",", ".");


    if (
      texto !== ""
    ) {

      const numero =
        Number(texto);


      if (
        isNaN(numero)
      ) {

        throw new Error(
          "A média deve ser um número válido."
        );

      }


      if (
        numero < 0 ||
        numero > 10
      ) {

        throw new Error(
          "A média deve estar entre 0 e 10."
        );

      }


      mediaFinal =
        Math.round(
          numero * 10
        ) / 10;

    }

  }


  const aba =
    getAba(
      ABAS.MEDIA
    );


  const dados =
    obterDadosAba(
      ABAS.MEDIA
    );


  // ==========================================================
  // PROCURAR MÉDIA JÁ EXISTENTE
  // ==========================================================

  for (
    let i = 0;
    i < dados.length;
    i++
  ) {

    const linha =
      dados[i];


    const idLinha =
      String(
        linha[0] || ""
      ).trim();


    const alunoLinha =
      String(
        linha[1] || ""
      ).trim();


    const materiaLinha =
      String(
        linha[4] || ""
      ).trim();


    const professorLinha =
      String(
        linha[6] || ""
      ).trim();


    if (
      alunoLinha === alunoRA
      &&
      normalizar(materiaLinha) ===
      normalizar(materia)
      &&
      professorLinha ===
      String(professor.ra)
    ) {

      const numeroLinha =
        i + 2;


      const celulaMedia =
        aba.getRange(
          numeroLinha,
          6
        );


      celulaMedia
        .setNumberFormat(
          "0.0"
        );


      if (
        mediaFinal === ""
      ) {

        celulaMedia
          .clearContent();

      } else {

        celulaMedia
          .setValue(
            mediaFinal
          );

      }


      aba
        .getRange(
          numeroLinha,
          9
        )
        .setValue(
          new Date()
        );


      return {

        sucesso: true,

        mensagem:
          "Média atualizada com sucesso.",

        alunoRA:
          alunoRA,

        media:
          mediaFinal,

        id:
          idLinha

      };

    }

  }


  // ==========================================================
  // CRIAR NOVA MÉDIA
  // ==========================================================

  const novaLinha = [

    gerarID("MED"),

    alunoRA,

    String(
      alunoEncontrado[1] || ""
    ),

    String(
      alunoEncontrado[5] || ""
    ),

    materia,

    mediaFinal,

    professor.ra,

    professor.nome,

    new Date()

  ];


  const primeiraLinha =
    aba.getLastRow() + 1;


  aba
    .getRange(
      primeiraLinha,
      1,
      1,
      9
    )
    .setValues([
      novaLinha
    ]);


  aba
    .getRange(
      primeiraLinha,
      6
    )
    .setNumberFormat(
      "0.0"
    );


  return {

    sucesso: true,

    mensagem:
      "Média lançada com sucesso.",

    alunoRA:
      alunoRA,

    media:
      mediaFinal

  };

}


// ============================================================
// SALVAR TODAS AS MÉDIAS
// ============================================================

function salvarMedias(
  token,
  materia,
  medias
) {

  const professor =
    obterSessao(token);


  if (
    professor.tipo !==
    "PROFESSOR"
  ) {

    throw new Error(
      "Somente professores podem lançar médias."
    );

  }


  materia =
    String(
      materia || ""
    )
      .trim();


  if (!materia) {

    throw new Error(
      "Informe a matéria."
    );

  }


  if (
    !Array.isArray(medias)
  ) {

    throw new Error(
      "Nenhuma média foi informada."
    );

  }


  let quantidade =
    0;


  medias.forEach(function(item) {

    if (!item) {

      return;

    }


    const alunoRA =
      String(
        item.alunoRA || ""
      )
        .trim();


    if (!alunoRA) {

      return;

    }


    salvarMedia(
      token,
      alunoRA,
      materia,
      item.media
    );


    quantidade++;

  });


  return {

    sucesso: true,

    quantidade:
      quantidade,

    mensagem:
      quantidade +
      " média(s) salva(s) com sucesso."

  };

}


// ============================================================
// LISTAR MÉDIAS DO ALUNO
// SOMENTE AS MÉDIAS DO ALUNO LOGADO
// ============================================================

function listarMediasAluno(
  token
) {

  const aluno =
    obterSessao(token);


  if (
    aluno.tipo !==
    "ALUNO"
  ) {

    throw new Error(
      "Apenas alunos podem acessar suas médias."
    );

  }


  const dados =
    obterDadosAba(
      ABAS.MEDIA
    );


  return dados

    .filter(function(linha) {

      return (
        String(
          linha[1] || ""
        ).trim() ===
        String(
          aluno.ra
        ).trim()
      );

    })

    .map(function(linha, index) {

      let media =
        linha[5];


      if (
        media === null ||
        media === undefined ||
        media === ""
      ) {

        media = "";

      }

      else if (
        Object.prototype.toString.call(media) ===
        "[object Date]"
      ) {

        media = "";

      }

      else {

        const numero =
          Number(
            String(media)
              .trim()
              .replace(",", ".")
          );


        if (
          !isNaN(numero)
        ) {

          media =
            numero
              .toFixed(1)
              .replace(".", ",");

        } else {

          media =
            String(media);

        }

      }


      return {

        linha:
          index + 2,

        id:
          String(
            linha[0] || ""
          ),

        alunoRA:
          String(
            linha[1] || ""
          ),

        aluno:
          String(
            linha[2] || ""
          ),

        serieAno:
          String(
            linha[3] || ""
          ),

        materia:
          String(
            linha[4] || ""
          ),

        media:
          media,

        professorRA:
          String(
            linha[6] || ""
          ),

        professor:
          String(
            linha[7] || ""
          ),

        dataAtualizacao:
          formatarData(
            linha[8]
          )

      };

    });

}


// ============================================================
// LISTAR TODAS AS MÉDIAS DO PROFESSOR
// ============================================================

function listarMediasProfessor(
  token
) {

  const professor =
    obterSessao(token);


  if (
    professor.tipo !==
    "PROFESSOR"
  ) {

    throw new Error(
      "Apenas professores podem acessar as médias."
    );

  }


  const dados =
    obterDadosAba(
      ABAS.MEDIA
    );


  return dados

    .map(function(linha, index) {

      let media =
        linha[5];


      if (
        media === null ||
        media === undefined ||
        media === ""
      ) {

        media = "";

      }

      else if (
        Object.prototype.toString.call(media) ===
        "[object Date]"
      ) {

        media = "";

      }

      else {

        const numero =
          Number(
            String(media)
              .trim()
              .replace(",", ".")
          );


        if (
          !isNaN(numero)
        ) {

          media =
            numero
              .toFixed(1)
              .replace(".", ",");

        } else {

          media =
            String(media);

        }

      }


      return {

        linha:
          index + 2,

        id:
          String(
            linha[0] || ""
          ),

        alunoRA:
          String(
            linha[1] || ""
          ),

        aluno:
          String(
            linha[2] || ""
          ),

        serieAno:
          String(
            linha[3] || ""
          ),

        materia:
          String(
            linha[4] || ""
          ),

        media:
          media,

        professorRA:
          String(
            linha[6] || ""
          ),

        professor:
          String(
            linha[7] || ""
          ),

        dataAtualizacao:
          formatarData(
            linha[8]
          )

      };

    })

    .filter(function(item) {

      return (
        String(
          item.professorRA
        ) ===
        String(
          professor.ra
        )
      );

    });

}


// ============================================================
// TRABALHOS
// NÃO ALTERADO
// ============================================================

function cadastrarTrabalho(
  token,
  dados
) {

  const professor =
    obterSessao(token);


  if (
    professor.tipo !==
    "PROFESSOR"
  ) {

    throw new Error(
      "Apenas professores podem cadastrar trabalhos."
    );

  }


  if (!dados.data)
    throw new Error(
      "Informe a data."
    );


  if (!dados.titulo)
    throw new Error(
      "Informe o título."
    );


  if (!dados.materia)
    throw new Error(
      "Informe a matéria."
    );


  const alunos =
    obterDadosAba(
      ABAS.ALUNO
    );


  let alunosDestino = [];


  if (
    dados.todosSerie === true ||
    String(
      dados.todosSerie
    ) === "true"
  ) {

    alunosDestino =
      alunos.filter(function(linha) {

        return (
          normalizar(linha[5]) ===
          normalizar(professor.serieAno)
        );

      });

  } else {

    const aluno =
      alunos.find(function(linha) {

        return (
          String(linha[0]) ===
          String(dados.alunoRA)
        );

      });


    if (!aluno) {

      throw new Error(
        "Aluno não encontrado pelo RA."
      );

    }


    if (
      normalizar(aluno[5]) !==
      normalizar(professor.serieAno)
    ) {

      throw new Error(
        "O aluno não pertence à sua Série/Ano."
      );

    }


    alunosDestino = [
      aluno
    ];

  }


  if (
    !alunosDestino.length
  ) {

    throw new Error(
      "Nenhum aluno encontrado nesta Série/Ano."
    );

  }


  const aba =
    getAba(
      ABAS.TRABALHO
    );


  const linhas = [];
  const idTrabalho = gerarID("TRAB");


  alunosDestino.forEach(function(aluno) {

    let nota = "";


    if (
      dados.nota !== undefined &&
      dados.nota !== null
    ) {

      nota =
        String(
          dados.nota
        )
          .trim()
          .replace(",", ".");


      if (
        nota !== "" &&
        !isNaN(Number(nota))
      ) {

        nota =
          Number(nota);

      }

    }


    linhas.push([

      idTrabalho,

      dados.data,

      dados.titulo,

      String(aluno[0]),

      String(aluno[1]),

      String(aluno[5]),

      dados.materia,

      nota,

      dados.descricao || "",

      professor.ra,

      professor.nome,

      new Date()

    ]);

  });


  const primeiraLinha =
    aba.getLastRow() + 1;


  aba
    .getRange(
      primeiraLinha,
      1,
      linhas.length,
      12
    )
    .setValues(
      linhas
    );

  criarEventosCalendarioParaPublicacao(
    professor,
    "TRABALHO",
    idTrabalho,
    dados.data,
    dados.titulo,
    dados.descricao || "",
    dados.materia,
    alunosDestino
  );


  return {

    sucesso: true,

    quantidade:
      alunosDestino.length,

    mensagem:
      dados.todosSerie
        ? "Trabalho publicado para " +
          alunosDestino.length +
          " aluno(s) da Série/Ano."
        : "Trabalho cadastrado para o aluno."

  };

}


// ============================================================
// LISTAR TRABALHOS
// NÃO ALTERADO
// ============================================================

function listarTrabalhos(
  token
) {

  const usuario =
    obterSessao(token);


  const dados =
    obterDadosAba(
      ABAS.TRABALHO
    );


  return dados

    .map(function(linha, index) {

      let nota =
        linha[7];


      if (
        nota === null ||
        nota === undefined ||
        nota === ""
      ) {

        nota = "";

      } else if (
        Object.prototype.toString.call(nota) ===
        "[object Date]"
      ) {

        nota = "";

      } else {

        const numero =
          Number(
            String(nota)
              .trim()
              .replace(",", ".")
          );


        if (
          !isNaN(numero)
        ) {

          nota =
            numero
              .toFixed(1)
              .replace(".", ",");

        } else {

          nota =
            String(nota);

        }

      }


      return {

        linha:
          index + 2,

        id:
          String(linha[0] || ""),

        data:
          formatarData(
            linha[1]
          ),

        titulo:
          String(
            linha[2] || ""
          ),

        alunoRA:
          String(
            linha[3] || ""
          ),

        aluno:
          String(
            linha[4] || ""
          ),

        serieAno:
          String(
            linha[5] || ""
          ),

        materia:
          String(
            linha[6] || ""
          ),

        nota:
          nota,

        descricao:
          String(
            linha[8] || ""
          ),

        professorRA:
          String(
            linha[9] || ""
          ),

        professor:
          String(
            linha[10] || ""
          )

      };

    })

    .filter(function(trabalho) {

      if (
        usuario.tipo ===
        "ALUNO"
      ) {

        return (
          String(
            trabalho.alunoRA
          ) ===
          String(
            usuario.ra
          )
        );

      }


      return (
        String(
          trabalho.professorRA
        ) ===
        String(
          usuario.ra
        )
      );

    });

}


// ============================================================
// EXCLUIR PROVA
// ============================================================

function excluirProva(
  token,
  linha
) {

  const professor =
    obterSessao(token);


  if (
    professor.tipo !==
    "PROFESSOR"
  ) {

    throw new Error(
      "Acesso negado."
    );

  }


  const aba =
    getAba(
      ABAS.PROVA
    );


  const numeroLinha =
    Number(linha);


  if (
    isNaN(numeroLinha) ||
    numeroLinha < 2 ||
    numeroLinha > aba.getLastRow()
  ) {

    throw new Error(
      "Linha da prova inválida."
    );

  }


  const dadosLinha =
    aba
      .getRange(
        numeroLinha,
        1,
        1,
        12
      )
      .getValues()[0];


  const idProva =
    String(
      dadosLinha[0] || ""
    );


  const professorRA =
    String(
      dadosLinha[9] || ""
    );


  if (
    professorRA !==
    String(
      professor.ra
    )
  ) {

    throw new Error(
      "Você não pode excluir esta prova."
    );

  }


  if (!idProva) {

    throw new Error(
      "ID da prova não encontrado."
    );

  }


  const todasLinhas =
    aba
      .getRange(
        2,
        1,
        aba.getLastRow() - 1,
        12
      )
      .getValues();


  const linhasExcluir = [];


  for (
    let i = 0;
    i < todasLinhas.length;
    i++
  ) {

    if (
      String(
        todasLinhas[i][0] || ""
      ) === idProva
    ) {

      linhasExcluir.push(
        i + 2
      );

    }

  }


  for (
    let i = linhasExcluir.length - 1;
    i >= 0;
    i--
  ) {

    aba.deleteRow(
      linhasExcluir[i]
    );

  }


  return {

    sucesso: true,

    mensagem:
      "Prova excluída com sucesso."

  };

}


// ============================================================
// EXCLUIR TRABALHO
// ============================================================

function excluirTrabalho(
  token,
  linha
) {

  const professor =
    obterSessao(token);


  if (
    professor.tipo !==
    "PROFESSOR"
  ) {

    throw new Error(
      "Acesso negado."
    );

  }


  const aba =
    getAba(
      ABAS.TRABALHO
    );


  const dados =
    aba
      .getRange(
        Number(linha),
        1,
        1,
        12
      )
      .getValues()[0];


  if (
    String(dados[9]) !==
    String(professor.ra)
  ) {

    throw new Error(
      "Você não pode excluir este trabalho."
    );

  }


  aba.deleteRow(
    Number(linha)
  );


  return {

    sucesso: true,

    mensagem:
      "Trabalho excluído."

  };

}


// ============================================================
// DÚVIDAS
// ============================================================

function cadastrarDuvida(
  token,
  dados
) {

  const aluno =
    obterSessao(token);


  if (
    aluno.tipo !==
    "ALUNO"
  ) {

    throw new Error(
      "Somente alunos podem enviar dúvidas."
    );

  }


  const assunto =
    String(
      dados.assunto || ""
    )
      .trim();


  const mensagem =
    String(
      dados.mensagem || ""
    )
      .trim();


  const professorRA =
    String(
      dados.professorRA || ""
    )
      .trim();


  if (!assunto) {

    throw new Error(
      "Informe o assunto."
    );

  }


  if (!mensagem) {

    throw new Error(
      "Digite sua dúvida."
    );

  }


  let professor =
    null;


  const professores =
    obterDadosAba(
      ABAS.PROFESSOR
    );


  if (
    professorRA
  ) {

    professor =
      professores.find(function(linha) {

        return (
          String(linha[1]) ===
          professorRA
        );

      });

  }


  if (!professor) {

    professor =
      professores.find(function(linha) {

        return (
          normalizar(linha[6]) ===
          normalizar(aluno.serieAno)
        );

      });

  }


  if (!professor) {

    throw new Error(
      "Nenhum professor encontrado para sua Série/Ano."
    );

  }


  getAba(
    ABAS.DUVIDA
  )
    .appendRow([

      gerarID("DUV"),

      new Date(),

      aluno.ra,

      aluno.nome,

      String(professor[1]),

      String(professor[2]),

      assunto,

      mensagem,

      "",

      "LIDA",

      "NAO_LIDA",

      ""

    ]);


  return {

    sucesso: true,

    mensagem:
      "Dúvida enviada ao professor."

  };

}


// ============================================================
// PROFESSORES
// ============================================================

function listarProfessores(
  token
) {

  const usuario =
    obterSessao(token);


  const professores =
    obterDadosAba(
      ABAS.PROFESSOR
    );


  if (
    usuario.tipo ===
    "ALUNO"
  ) {

    return professores

      .map(function(linha) {

        return {

          ra:
            String(linha[1]),

          nome:
            String(linha[2]),

          materia:
            String(linha[7]),

          serieAno:
            String(linha[6])

        };

      })

      .filter(function(professor) {

        return (
          normalizar(
            professor.serieAno
          ) ===
          normalizar(
            usuario.serieAno
          )
        );

      });

  }


  return [];

}


// ============================================================
// DÚVIDAS - LISTAR
// ============================================================

function listarDuvidas(
  token,
  filtro
) {

  const usuario =
    obterSessao(token);


  const dados =
    obterDadosAba(
      ABAS.DUVIDA
    );


  let lista =
    dados.map(function(linha, index) {

      return {

        linha:
          index + 2,

        id:
          String(linha[0]),

        data:
          formatarData(
            linha[1]
          ),

        alunoRA:
          String(linha[2]),

        aluno:
          String(linha[3]),

        professorRA:
          String(linha[4]),

        professor:
          String(linha[5]),

        assunto:
          String(linha[6]),

        mensagem:
          String(linha[7]),

        resposta:
          String(linha[8] || ""),

        statusAluno:
          String(
            linha[9] ||
            "LIDA"
          ),

        statusProfessor:
          String(
            linha[10] ||
            "NAO_LIDA"
          ),

        dataResposta:
          formatarData(
            linha[11]
          )

      };

    });


  if (
    usuario.tipo ===
    "ALUNO"
  ) {

    lista =
      lista.filter(function(item) {

        return (
          String(item.alunoRA) ===
          String(usuario.ra)
        );

      });

  } else {

    lista =
      lista.filter(function(item) {

        return (
          String(item.professorRA) ===
          String(usuario.ra)
        );

      });

  }


  filtro =
    String(
      filtro || "TODAS"
    )
      .toUpperCase();


  if (
    filtro ===
    "NAO_LIDAS"
  ) {

    lista =
      lista.filter(function(item) {

        if (
          usuario.tipo ===
          "PROFESSOR"
        ) {

          return (
            item.statusProfessor ===
            "NAO_LIDA"
          );

        }


        return (
          item.statusAluno ===
          "NAO_LIDA"
        );

      });

  }


  if (
    filtro ===
    "LIDAS"
  ) {

    lista =
      lista.filter(function(item) {

        if (
          usuario.tipo ===
          "PROFESSOR"
        ) {

          return (
            item.statusProfessor ===
            "LIDA"
          );

        }


        return (
          item.statusAluno ===
          "LIDA"
        );

      });

  }


  return lista.reverse();

}


// ============================================================
// MARCAR DÚVIDA COMO LIDA
// ============================================================

function marcarDuvidaLida(
  token,
  id
) {

  const usuario =
    obterSessao(token);


  const aba =
    getAba(
      ABAS.DUVIDA
    );


  const dados =
    obterDadosAba(
      ABAS.DUVIDA
    );


  for (
    let i = 0;
    i < dados.length;
    i++
  ) {

    if (
      String(dados[i][0]) ===
      String(id)
    ) {

      const linha =
        i + 2;


      if (
        usuario.tipo ===
        "PROFESSOR"
      ) {

        if (
          String(dados[i][4]) !==
          String(usuario.ra)
        ) {

          throw new Error(
            "Acesso negado."
          );

        }


        aba
          .getRange(
            linha,
            11
          )
          .setValue(
            "LIDA"
          );

      } else {

        if (
          String(dados[i][2]) !==
          String(usuario.ra)
        ) {

          throw new Error(
            "Acesso negado."
          );

        }


        aba
          .getRange(
            linha,
            10
          )
          .setValue(
            "LIDA"
          );

      }


      return {

        sucesso: true,

        mensagem:
          "Dúvida marcada como lida."

      };

    }

  }


  throw new Error(
    "Dúvida não encontrada."
  );

}


// ============================================================
// RESPONDER DÚVIDA
// ============================================================

function responderDuvida(
  token,
  id,
  resposta
) {

  const professor =
    obterSessao(token);


  if (
    professor.tipo !==
    "PROFESSOR"
  ) {

    throw new Error(
      "Somente professores podem responder dúvidas."
    );

  }


  resposta =
    String(
      resposta || ""
    )
      .trim();


  if (!resposta) {

    throw new Error(
      "Digite uma resposta."
    );

  }


  const aba =
    getAba(
      ABAS.DUVIDA
    );


  const dados =
    obterDadosAba(
      ABAS.DUVIDA
    );


  for (
    let i = 0;
    i < dados.length;
    i++
  ) {

    if (
      String(dados[i][0]) ===
      String(id)
    ) {

      if (
        String(dados[i][4]) !==
        String(professor.ra)
      ) {

        throw new Error(
          "Você não pode responder esta dúvida."
        );

      }


      const linha =
        i + 2;


      aba
        .getRange(
          linha,
          9
        )
        .setValue(
          resposta
        );


      aba
        .getRange(
          linha,
          10
        )
        .setValue(
          "NAO_LIDA"
        );


      aba
        .getRange(
          linha,
          11
        )
        .setValue(
          "LIDA"
        );


      aba
        .getRange(
          linha,
          12
        )
        .setValue(
          new Date()
        );


      return {

        sucesso: true,

        mensagem:
          "Resposta enviada ao aluno."

      };

    }

  }


  throw new Error(
    "Dúvida não encontrada."
  );

}


// ============================================================
// ESTATÍSTICAS
// ============================================================

function obterEstatisticas(
  token
) {

  const usuario =
    obterSessao(token);


  const materias =
    listarMaterias(
      token
    );


  const provas =
    listarProvas(
      token
    );


  const trabalhos =
    listarTrabalhos(
      token
    );


  let alunos = [];


  if (
    usuario.tipo ===
    "PROFESSOR"
  ) {

    alunos =
      listarAlunos(
        token
      );

  }


  return {

    materias:
      materias.length,

    provas:
      provas.length,

    trabalhos:
      trabalhos.length,

    alunos:
      alunos.length

  };

}


// ============================================================
// RECUPERAÇÃO DE SENHA
// MANTIDA
// ============================================================

function recuperarSenha(
  dados
) {

  try {

    if (!dados) {

      return {
        sucesso: false,
        mensagem:
          "Dados não informados."
      };

    }


    const tipo =
      String(
        dados.tipo || ""
      )
        .trim()
        .toUpperCase();


    const ra =
      String(
        dados.ra || ""
      )
        .trim();


    const usuario =
      String(
        dados.usuario || ""
      )
        .trim();


    const novaSenha =
      String(
        dados.novaSenha || ""
      );


    if (
      tipo !== "ALUNO" &&
      tipo !== "PROFESSOR"
    ) {

      return {
        sucesso: false,
        mensagem:
          "Tipo de conta inválido."
      };

    }


    if (!ra) {

      return {
        sucesso: false,
        mensagem:
          "Informe o RA/ID."
      };

    }


    if (!usuario) {

      return {
        sucesso: false,
        mensagem:
          "Informe o usuário."
      };

    }


    if (!novaSenha) {

      return {
        sucesso: false,
        mensagem:
          "Informe a nova senha."
      };

    }


    if (
      novaSenha.length < 6
    ) {

      return {
        sucesso: false,
        mensagem:
          "A nova senha precisa ter pelo menos 6 caracteres."
      };

    }


    // ========================================================
    // ALUNO
    // ========================================================

    if (
      tipo === "ALUNO"
    ) {

      const aba =
        getAba(
          ABAS.ALUNO
        );


      const alunos =
        obterDadosAba(
          ABAS.ALUNO
        );


      for (
        let i = 0;
        i < alunos.length;
        i++
      ) {

        const linha =
          alunos[i];


        const raPlanilha =
          String(
            linha[0] || ""
          )
            .trim();


        const usuarioPlanilha =
          String(
            linha[2] || ""
          )
            .trim();


        if (
          normalizar(raPlanilha) ===
          normalizar(ra)
          &&
          normalizar(usuarioPlanilha) ===
          normalizar(usuario)
        ) {

          const numeroLinha =
            i + 2;


          const novoSalt =
            gerarSalt();


          const novoHash =
            gerarHashSenha(
              novaSenha,
              novoSalt
            );


          aba
            .getRange(
              numeroLinha,
              4
            )
            .setValue(
              novoHash
            );


          aba
            .getRange(
              numeroLinha,
              5
            )
            .setValue(
              novoSalt
            );


          return {

            sucesso: true,

            mensagem:
              "Senha do aluno alterada com sucesso!"

          };

        }

      }


      return {

        sucesso: false,

        mensagem:
          "RA ou usuário do aluno não encontrado."

      };

    }


    // ========================================================
    // PROFESSOR
    // ========================================================

    if (
      tipo === "PROFESSOR"
    ) {

      const aba =
        getAba(
          ABAS.PROFESSOR
        );


      const professores =
        obterDadosAba(
          ABAS.PROFESSOR
        );


      for (
        let i = 0;
        i < professores.length;
        i++
      ) {

        const linha =
          professores[i];


        const raPlanilha =
          String(
            linha[1] || ""
          )
            .trim();


        const usuarioPlanilha =
          String(
            linha[3] || ""
          )
            .trim();


        if (
          normalizar(raPlanilha) ===
          normalizar(ra)
          &&
          normalizar(usuarioPlanilha) ===
          normalizar(usuario)
        ) {

          const numeroLinha =
            i + 2;


          const novoSalt =
            gerarSalt();


          const novoHash =
            gerarHashSenha(
              novaSenha,
              novoSalt
            );


          aba
            .getRange(
              numeroLinha,
              5
            )
            .setValue(
              novoHash
            );


          aba
            .getRange(
              numeroLinha,
              6
            )
            .setValue(
              novoSalt
            );


          return {

            sucesso: true,

            mensagem:
              "Senha do professor alterada com sucesso!"

          };

        }

      }


      return {

        sucesso: false,

        mensagem:
          "RA ou usuário do professor não encontrado."

      };

    }


    return {

      sucesso: false,

      mensagem:
        "Tipo de conta inválido."

    };

  }

  catch (erro) {

    console.error(
      erro
    );


    return {

      sucesso: false,

      mensagem:
        "Ocorreu um erro ao redefinir a senha."

    };

  }

}



// ============================================================
// CALENDÁRIO ESCOLAR
// ============================================================

function obterTurmasDaSerie(alunos, serieAno) {
  const turmas = {};

  alunos.forEach(function(aluno) {
    if (normalizar(aluno[5]) !== normalizar(serieAno)) return;

    const turma = String(aluno[6] || "").trim();
    if (turma) turmas[normalizar(turma)] = turma;
  });

  return Object.keys(turmas).map(function(chave) {
    return turmas[chave];
  });
}

function sincronizarEventoComRegistro(professor, dadosEvento, origemId, turmaAnterior) {
  const aba = getAba(ABAS.CALENDARIO);

  const tipo = String(dadosEvento.tipo || "").toUpperCase();
  const prefixo = tipo === "PROVA" ? "PROVA" : tipo === "TRABALHO" ? "TRABALHO" : "";

  if (!prefixo || !origemId) return;

  const idPrefixo = prefixo + "|" + String(origemId) + "|";
  const alunos = obterDadosAba(ABAS.ALUNO);
  const turmaAtual = String(dadosEvento.turma || "").trim();
  const turmaAnteriorNormalizada = normalizar(turmaAnterior || "");
  const turmaAtualNormalizada = normalizar(turmaAtual);
  const turmas = [turmaAtual];

  // Se a turma mudou, os alunos da nova turma passam a receber
  // a data/título deste evento. A turma anterior mantém seu registro.
  // Isso permite que cada turma tenha sua própria data no calendário.

  const mapaTurma = {};
  alunos.forEach(function(aluno) {
    mapaTurma[String(aluno[0])] = String(aluno[6] || "");
  });

  const abaRegistro = getAba(prefixo === "PROVA" ? ABAS.PROVA : ABAS.TRABALHO);
  if (abaRegistro.getLastRow() < 2) return;

  const registros = abaRegistro.getRange(2, 1, abaRegistro.getLastRow() - 1, 12).getValues();

  registros.forEach(function(linha, indice) {
    if (String(linha[0] || "") !== String(origemId)) return;
    if (String(linha[9] || "") !== String(professor.ra)) return;

    const turmaAluno = mapaTurma[String(linha[3] || "")] || "";
    const pertence = turmas.some(function(turma) {
      return normalizar(turmaAluno) === normalizar(turma);
    });

    if (!pertence) return;

    abaRegistro.getRange(indice + 2, 2, 1, 2).setValues([[
      dadosEvento.data,
      dadosEvento.titulo
    ]]);

    abaRegistro.getRange(indice + 2, 7, 1, 1).setValue(dadosEvento.materia);
    abaRegistro.getRange(indice + 2, 9, 1, 1).setValue(dadosEvento.descricao || "");
  });
}

function criarEventosCalendarioParaPublicacao(professor, tipo, origemId, data, titulo, descricao, materia, alunos) {
  const aba = getAba(ABAS.CALENDARIO);
  const turmas = obterTurmasDaSerie(alunos, professor.serieAno);

  turmas.forEach(function(turma) {
    const idEvento = tipo + "|" + String(origemId) + "|" + turma;

    const existentes = aba.getLastRow() >= 2
      ? aba.getRange(2, 1, aba.getLastRow() - 1, 11).getValues()
      : [];

    const jaExiste = existentes.some(function(linha) {
      return String(linha[0] || "") === idEvento;
    });

    if (!jaExiste) {
      aba.appendRow([
        idEvento,
        data,
        tipo,
        titulo,
        descricao || "",
        professor.serieAno,
        turma,
        materia,
        professor.ra,
        professor.nome,
        new Date()
      ]);
    }
  });
}

function atualizarRegistroAoEditarCalendario(professor, abaCalendario, linhaEvento, dadosEvento, turmaAnterior) {
  const atual = abaCalendario.getRange(linhaEvento, 1, 1, 11).getValues()[0];
  const idEvento = String(atual[0] || "");
  const partes = idEvento.split("|");

  if (partes.length < 3) return;

  const tipoOrigem = partes[0];
  const origemId = partes[1];

  if (tipoOrigem !== "PROVA" && tipoOrigem !== "TRABALHO") return;

  sincronizarEventoComRegistro(
    professor,
    dadosEvento,
    origemId,
    turmaAnterior
  );
}

function cadastrarEventoCalendario(token, dados) {

  const professor = obterSessao(token);

  if (professor.tipo !== "PROFESSOR") {
    throw new Error("Apenas professores podem cadastrar eventos.");
  }

  if (!dados) {
    throw new Error("Dados do evento não informados.");
  }

  const data = String(dados.data || "").trim();
  const tipo = String(dados.tipo || "").trim().toUpperCase();
  const titulo = String(dados.titulo || "").trim();
  const descricao = String(dados.descricao || "").trim();
  const turma = String(dados.turma || "").trim();
  const materia = String(dados.materia || professor.materia || "").trim();

  const tiposValidos = ["AULA", "PROVA", "TRABALHO"];

  if (!data) throw new Error("Informe a data.");
  if (!tiposValidos.includes(tipo)) throw new Error("Tipo de evento inválido.");
  if (!titulo) throw new Error("Informe o título.");
  if (!turma) throw new Error("Informe a turma.");
  if (!materia) throw new Error("Informe a matéria.");

  const idOrigem = String(dados.origemId || "").trim();
  const id = idOrigem
    ? tipo + "|" + idOrigem + "|" + turma
    : gerarID("CAL");

  getAba(ABAS.CALENDARIO).appendRow([
    id,
    data,
    tipo,
    titulo,
    descricao,
    professor.serieAno,
    turma,
    materia,
    professor.ra,
    professor.nome,
    new Date()
  ]);

  return {
    sucesso: true,
    id: id,
    mensagem: "Evento adicionado ao calendário."
  };
}


function normalizarDataCalendario(valor) {
  if (!valor) return "";

  // Datas salvas como texto no formato usado pelo input type=date.
  if (typeof valor === "string") {
    const texto = valor.trim();
    if (/^\d{4}-\d{2}-\d{2}$/.test(texto)) return texto;
    if (/^\d{2}\/\d{2}\/\d{4}$/.test(texto)) {
      const p = texto.split("/");
      return p[2] + "-" + p[1] + "-" + p[0];
    }
  }

  // Caso o Google Sheets tenha convertido a célula para Date.
  if (Object.prototype.toString.call(valor) === "[object Date]" && !isNaN(valor.getTime())) {
    return Utilities.formatDate(valor, Session.getScriptTimeZone() || "America/Sao_Paulo", "yyyy-MM-dd");
  }

  const texto = String(valor).trim();
  const convertido = new Date(texto);
  if (!isNaN(convertido.getTime())) {
    return Utilities.formatDate(convertido, Session.getScriptTimeZone() || "America/Sao_Paulo", "yyyy-MM-dd");
  }

  return texto;
}

function listarEventosCalendario(token) {

  const usuario = obterSessao(token);
  const dados = obterDadosAba(ABAS.CALENDARIO);

  return dados
    .map(function(linha, index) {
      return {
        linha: index + 2,
        id: String(linha[0] || ""),
        data: normalizarDataCalendario(linha[1]),
        tipo: String(linha[2] || "").trim().toUpperCase(),
        titulo: String(linha[3] || ""),
        descricao: String(linha[4] || ""),
        serieAno: String(linha[5] || ""),
        turma: String(linha[6] || "").trim(),
        materia: String(linha[7] || ""),
        professorRA: String(linha[8] || ""),
        professor: String(linha[9] || "")
      };
    })
    .filter(function(evento) {

      if (usuario.tipo === "ALUNO") {
        return (
          normalizar(evento.serieAno) === normalizar(usuario.serieAno) &&
          normalizar(evento.turma) === normalizar(usuario.turma)
        );
      }

      return String(evento.professorRA) === String(usuario.ra);
    })
    .sort(function(a, b) {
      return String(a.data).localeCompare(String(b.data));
    });
}


function editarEventoCalendario(token, dados) {

  const professor = obterSessao(token);

  if (professor.tipo !== "PROFESSOR") {
    throw new Error("Apenas professores podem editar eventos.");
  }

  if (!dados || !dados.linha) {
    throw new Error("Evento inválido.");
  }

  const linha = Number(dados.linha);
  const aba = getAba(ABAS.CALENDARIO);

  if (
    isNaN(linha) ||
    linha < 2 ||
    linha > aba.getLastRow()
  ) {
    throw new Error("Linha do evento inválida.");
  }

  const atual = aba.getRange(linha, 1, 1, 11).getValues()[0];
  const idEventoAnterior = String(atual[0] || "");
  const turmaAnterior = String(atual[6] || "").trim();

  if (String(atual[8] || "") !== String(professor.ra)) {
    throw new Error("Você não pode editar este evento.");
  }

  const data = String(dados.data || "").trim();
  const tipo = String(dados.tipo || "").trim().toUpperCase();
  const titulo = String(dados.titulo || "").trim();
  const descricao = String(dados.descricao || "").trim();
  const turma = String(dados.turma || "").trim();
  const materia = String(dados.materia || professor.materia || "").trim();

  if (!data || !titulo || !turma || !materia) {
    throw new Error("Preencha data, título, turma e matéria.");
  }

  if (!["AULA", "PROVA", "TRABALHO"].includes(tipo)) {
    throw new Error("Tipo de evento inválido.");
  }

  const partesOrigem = idEventoAnterior.split("|");
  const origemVinculada = partesOrigem.length >= 3 &&
    (partesOrigem[0] === "PROVA" || partesOrigem[0] === "TRABALHO");

  aba.getRange(linha, 1, 1, 8).setValues([[
    origemVinculada
      ? partesOrigem[0] + "|" + partesOrigem[1] + "|" + turma
      : idEventoAnterior,
    data,
    tipo,
    titulo,
    descricao,
    professor.serieAno,
    turma,
    materia
  ]]);

  if (origemVinculada && partesOrigem[0] === tipo) {
    const novoIdEvento = partesOrigem[0] + "|" + partesOrigem[1] + "|" + turma;

    // Evita dois eventos para a mesma prova/trabalho e a mesma turma.
    if (aba.getLastRow() >= 2) {
      const eventos = aba.getRange(2, 1, aba.getLastRow() - 1, 11).getValues();
      for (let i = eventos.length - 1; i >= 0; i--) {
        const linhaReal = i + 2;
        if (linhaReal !== linha && String(eventos[i][0] || "") === novoIdEvento) {
          aba.deleteRow(linhaReal);
        }
      }
    }

    atualizarRegistroAoEditarCalendario(
      professor,
      aba,
      linha,
      {
        data: data,
        titulo: titulo,
        descricao: descricao,
        materia: materia,
        tipo: tipo,
        turma: turma
      },
      turmaAnterior
    );
  }

  return {
    sucesso: true,
    mensagem: "Evento atualizado com sucesso."
  };
}


function excluirEventoCalendario(token, linha) {

  const professor = obterSessao(token);

  if (professor.tipo !== "PROFESSOR") {
    throw new Error("Apenas professores podem excluir eventos.");
  }

  const numeroLinha = Number(linha);
  const aba = getAba(ABAS.CALENDARIO);

  if (
    isNaN(numeroLinha) ||
    numeroLinha < 2 ||
    numeroLinha > aba.getLastRow()
  ) {
    throw new Error("Linha do evento inválida.");
  }

  const dados = aba.getRange(numeroLinha, 1, 1, 11).getValues()[0];

  if (String(dados[8] || "") !== String(professor.ra)) {
    throw new Error("Você não pode excluir este evento.");
  }

  aba.deleteRow(numeroLinha);

  return {
    sucesso: true,
    mensagem: "Evento excluído do calendário."
  };
}
