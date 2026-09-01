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
  DUVIDA: "DUVIDA"
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

      gerarID("TRAB"),

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
<!DOCTYPE html>
<html lang="pt-BR">

<head>

<meta charset="UTF-8">

<meta name="viewport"
content="width=device-width, initial-scale=1">

<base target="_top">

<title>Gestão Escolar</title>

<style>

* {
  box-sizing: border-box;
  margin: 0;
  padding: 0;
}

body {
  font-family: Arial, Helvetica, sans-serif;
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

button {
  cursor: pointer;
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
  background: linear-gradient(
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
  box-shadow: 0 30px 80px rgba(0,0,0,.30);
}

.logo {
  width: 72px;
  height: 72px;
  margin: 0 auto 20px;
  display: flex;
  justify-content: center;
  align-items: center;
  border-radius: 20px;
  background: linear-gradient(
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

.type-buttons {
  display: grid;
  grid-template-columns: 1fr 1fr;
  gap: 8px;
  margin-bottom: 20px;
}

.type-buttons button {
  padding: 12px;
  border: 1px solid #cbd5e1;
  background: white;
  border-radius: 10px;
  font-weight: bold;
}

.type-buttons button.active {
  background: #eff6ff;
  color: #2563eb;
  border-color: #2563eb;
}

label {
  display: block;
  margin: 13px 0 6px;
  font-size: 14px;
  font-weight: bold;
}

input,
select,
textarea {
  width: 100%;
  padding: 12px;
  border: 1px solid #cbd5e1;
  border-radius: 10px;
  outline: none;
  font-size: 15px;
}

textarea {
  min-height: 110px;
  resize: vertical;
}

input:focus,
select:focus,
textarea:focus {
  border-color: #2563eb;
  box-shadow: 0 0 0 3px rgba(37,99,235,.12);
}

.btn {
  border: none;
  padding: 12px 16px;
  border-radius: 10px;
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

.btn-warning {
  background: #f59e0b;
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
  font-weight: bold;
}

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
  overflow-y: auto;
}

.brand {
  display: flex;
  align-items: center;
  gap: 10px;
  padding-bottom: 22px;
  border-bottom: 1px solid rgba(255,255,255,.1);
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
  margin: 22px 8px 8px;
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
  font-weight: bold;
}

.nav:hover,
.nav.active {
  background: #1d4ed8;
  color: white;
}

.main {
  margin-left: 255px;
  width: calc(100% - 255px);
}

.top {
  height: 72px;
  padding: 0 25px;
  display: flex;
  align-items: center;
  justify-content: space-between;
  background: white;
  border-bottom: 1px solid #e2e8f0;
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

.hero {
  padding: 30px;
  border-radius: 18px;
  color: white;
  margin-bottom: 20px;
  background: linear-gradient(
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

.stats {
  display: grid;
  grid-template-columns:
    repeat(auto-fit,minmax(180px,1fr));
  gap: 15px;
  margin-bottom: 20px;
}

.stat {
  padding: 20px;
  border-radius: 15px;
  background: white;
  box-shadow: 0 5px 18px rgba(15,23,42,.06);
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

.card {
  background: white;
  padding: 22px;
  border-radius: 15px;
  margin-bottom: 20px;
  box-shadow: 0 5px 18px rgba(15,23,42,.06);
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

.grid {
  display: grid;
  grid-template-columns:
    repeat(auto-fit,minmax(280px,1fr));
  gap: 15px;
}

.item {
  border: 1px solid #e2e8f0;
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

.subject {
  border: 1px solid #e2e8f0;
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
  border-top: 1px solid #e2e8f0;
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

.note {
  display: inline-block;
  margin-top: 8px;
  padding: 6px 10px;
  border-radius: 15px;
  background: #dcfce7;
  color: #166534;
  font-weight: bold;
  font-size: 13px;
}

.empty {
  grid-column: 1 / -1;
  text-align: center;
  padding: 40px;
  color: #64748b;
}


/* ==========================================================
MÉDIAS / NOTAS
========================================================== */

.media-toolbar {
  display: grid;
  grid-template-columns: minmax(220px, 320px) 1fr;
  gap: 15px;
  align-items: end;
}

.media-table-wrap {
  width: 100%;
  overflow-x: auto;
  border: 1px solid #e2e8f0;
  border-radius: 12px;
}

.media-table {
  width: 100%;
  min-width: 700px;
  border-collapse: collapse;
  background: white;
}

.media-table th {
  background: #eff6ff;
  color: #1e3a8a;
  text-align: left;
  padding: 13px;
  font-size: 13px;
  border-bottom: 1px solid #dbeafe;
}

.media-table td {
  padding: 11px 13px;
  border-bottom: 1px solid #e2e8f0;
  font-size: 14px;
}

.media-table tr:last-child td {
  border-bottom: none;
}

.media-table input {
  max-width: 130px;
  padding: 9px 10px;
}

.media-table .media-name {
  font-weight: bold;
}

.media-readonly {
  font-weight: bold;
  font-size: 17px;
  color: #166534;
}

.media-status {
  padding: 6px 10px;
  border-radius: 20px;
  background: #f1f5f9;
  color: #64748b;
  font-size: 12px;
  display: inline-block;
}

.media-status.lancada {
  background: #dcfce7;
  color: #166534;
}

.media-status.pendente {
  background: #fef3c7;
  color: #92400e;
}

.media-info {
  background: #f8fafc;
  border: 1px solid #e2e8f0;
  padding: 14px;
  border-radius: 10px;
  color: #475569;
  font-size: 13px;
  margin-bottom: 15px;
}

.media-acoes {
  display: flex;
  gap: 10px;
  flex-wrap: wrap;
  margin-top: 15px;
}

.media-loading {
  text-align: center;
  padding: 30px;
  color: #64748b;
}

.media-aluno-card {
  border: 1px solid #e2e8f0;
  border-radius: 14px;
  overflow: hidden;
}

.media-aluno-card-head {
  padding: 17px;
  background: #eff6ff;
}

.media-aluno-card-head h3 {
  color: #1e3a8a;
  margin-bottom: 5px;
}

.media-aluno-card-body {
  padding: 16px;
}

.media-valor-grande {
  font-size: 30px;
  font-weight: bold;
  color: #166534;
  margin-top: 8px;
}


/* ==========================================================
DÚVIDAS
========================================================== */

.duvida {
  border: 1px solid #e2e8f0;
  border-radius: 14px;
  padding: 18px;
  background: white;
}

.duvida-nao-lida {
  border-left: 5px solid #2563eb;
  background: #eff6ff;
}

.duvida-lida {
  border-left: 5px solid #16a34a;
}

.status {
  display: inline-block;
  padding: 5px 9px;
  border-radius: 20px;
  font-size: 12px;
  font-weight: bold;
  margin-top: 8px;
}

.status-nao {
  background: #dbeafe;
  color: #1d4ed8;
}

.status-sim {
  background: #dcfce7;
  color: #166534;
}

.duvida-mensagem {
  background: #f8fafc;
  padding: 14px;
  border-radius: 10px;
  margin-top: 12px;
  white-space: pre-line;
}

.resposta {
  background: #ecfdf5;
  border: 1px solid #bbf7d0;
  padding: 14px;
  border-radius: 10px;
  margin-top: 12px;
  white-space: pre-line;
}

.filtro-duvidas {
  display: flex;
  gap: 8px;
  flex-wrap: wrap;
}

.filtro-duvidas button {
  border: 1px solid #cbd5e1;
  background: white;
  color: #334155;
  padding: 9px 12px;
  border-radius: 8px;
  font-weight: bold;
}

.filtro-duvidas button.active {
  background: #2563eb;
  color: white;
  border-color: #2563eb;
}


/* ==========================================================
MOBILE
========================================================== */

#menuMobile {
  display: none;
  border: none;
  background: #f1f5f9;
  padding: 8px 12px;
  border-radius: 8px;
}

@media(max-width:850px) {

  .sidebar {
    transform: translateX(-100%);
    transition: .2s;
  }

  .sidebar.open {
    transform: translateX(0);
  }

  .main {
    margin-left: 0;
    width: 100%;
  }

  #menuMobile {
    display: block;
  }

  .media-toolbar {
    grid-template-columns: 1fr;
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

  .login-card {
    padding: 25px;
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

<div class="logo">🏫</div>

<h1>Gestão Escolar</h1>

<p class="subtitle">
Acesse sua conta
</p>

<div class="type-buttons">

<button id="loginAluno" class="active">
👨‍🎓 Aluno
</button>

<button id="loginProfessor">
👨‍🏫 Professor
</button>

</div>

<button
id="esqueciSenhaBtn"
class="link-button">
🔑 Esqueci minha senha
</button>

<label>Usuário</label>

<input
id="loginUsuario"
placeholder="Digite seu usuário">

<label>Senha</label>

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

<div id="loginMsg" class="message"></div>

</div>

</div>


<!-- ==========================================================
CADASTRO
========================================================== -->

<div id="registerPage"
class="login-page hidden">

<div class="login-card">

<div class="logo">📝</div>

<h1>Criar conta</h1>

<p class="subtitle">
Escolha o tipo de usuário
</p>

<div class="type-buttons">

<button id="cadAluno" class="active">
👨‍🎓 Aluno
</button>

<button id="cadProfessor">
👨‍🏫 Professor
</button>

</div>

<label>RA / ID</label>

<input
id="cadRA"
placeholder="RA do aluno ou ID do professor">

<label>Nome completo</label>

<input
id="cadNome"
placeholder="Digite seu nome">

<label>Usuário</label>

<input
id="cadUsuario"
placeholder="Escolha seu usuário">

<label>Senha</label>

<input
id="cadSenha"
type="password"
placeholder="Mínimo 6 caracteres">

<label>Série / Ano</label>

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

<div id="cadTurmaBox">

<label>Turma</label>

<input
id="cadTurma"
placeholder="Ex: A">

</div>

<div id="cadMateriaBox"
class="hidden">

<label>Matéria</label>

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

<div id="cadMsg" class="message"></div>

</div>

</div>


<!-- ==========================================================
RECUPERAÇÃO DE SENHA
========================================================== -->

<div
id="forgotPasswordPage"
class="login-page hidden">

<div class="login-card">

<div class="logo">🔑</div>

<h1>Recuperar senha</h1>

<p class="subtitle">
Crie uma nova senha para sua conta
</p>

<div class="type-buttons">

<button
id="forgotAluno"
class="active">
👨‍🎓 Aluno
</button>

<button
id="forgotProfessor">
👨‍🏫 Professor
</button>

</div>

<label>RA / ID</label>

<input
id="forgotRA"
placeholder="Digite seu RA ou ID">

<label>Usuário</label>

<input
id="forgotUsuario"
placeholder="Digite seu usuário">

<label>Nova senha</label>

<input
id="forgotNovaSenha"
type="password"
placeholder="Mínimo 6 caracteres">

<label>Confirmar nova senha</label>

<input
id="forgotConfirmarSenha"
type="password"
placeholder="Digite a senha novamente">

<button
id="alterarSenhaBtn"
class="btn btn-primary full">
🔐 Alterar senha
</button>

<button
id="voltarLoginSenha"
class="btn btn-gray full">
← Voltar para o login
</button>

<div
id="forgotMsg"
class="message">
</div>

</div>

</div>


<!-- ==========================================================
APP
========================================================== -->

<div id="app">

<div id="mobileOverlay"></div>

<div class="layout">

<aside id="sidebar" class="sidebar">

<div class="brand">

<div class="brand-icon">🏫</div>

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

<!-- NOVO -->
<button
id="lancarNotasNav"
class="nav"
data-screen="lancarNotas"
style="display:none;">
📊 Lançar Notas
</button>

<!-- NOVO -->
<button
id="minhasNotasNav"
class="nav"
data-screen="minhasNotas"
style="display:none;">
📊 Minhas Notas
</button>

<button
id="duvidasNav"
class="nav"
data-screen="duvidas">
❓ Dúvidas
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

<button id="logoutBtn" class="nav">
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

<div id="avatar" class="avatar">
U
</div>

<div>

<div id="userName"
class="user-name">
Usuário
</div>

<span id="userType"
class="user-type">
Conta
</span>

</div>

</div>

</div>


<div class="content">


<!-- ==========================================================
DASHBOARD
========================================================== -->

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


<!-- ==========================================================
MATÉRIAS
========================================================== -->

<section
id="screenMaterias"
class="screen">

<div
id="materiaForm"
class="card">

<h2>📚 Nova matéria</h2>

<label>Nome da matéria</label>

<input
id="materiaNome"
placeholder="Ex: Matemática">

<label>Conteúdo</label>

<textarea
id="materiaConteudo"
placeholder="Conteúdo da aula">
</textarea>

<label>Links de estudo</label>

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

<h2>📚 Matérias</h2>

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


<!-- ==========================================================
PROVAS
========================================================== -->

<section
id="screenProvas"
class="screen">

<div
id="provaForm"
class="card">

<h2>📝 Nova prova</h2>

<label>Data</label>

<input
id="provaData"
type="date">

<label>Título</label>

<input
id="provaTitulo"
placeholder="Ex: Prova de Matemática">

<label>Destino da prova</label>

<select id="provaDestino">

<option value="ALUNO">
Apenas um aluno
</option>

<option value="TODOS">
Todos os alunos da minha Série/Ano
</option>

</select>

<div id="provaAlunoBox">

<label>Aluno</label>

<select id="provaAluno">

<option value="">
Selecione o aluno
</option>

</select>

</div>

<label>Matéria</label>

<input
id="provaMateria"
placeholder="Ex: Matemática">

<label>Nota</label>

<input
id="provaNota"
type="number"
min="0"
max="10"
step="0.1"
placeholder="0 a 10">

<label>Descrição</label>

<textarea
id="provaDescricao"
placeholder="Descrição da prova">
</textarea>

<button
id="provaSalvar"
class="btn btn-primary">
💾 Publicar prova
</button>

</div>

<div class="card">

<div class="card-header">

<h2>📝 Provas e Notas</h2>

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


<!-- ==========================================================
TRABALHOS
========================================================== -->

<section
id="screenTrabalhos"
class="screen">

<div
id="trabalhoForm"
class="card">

<h2>📂 Novo trabalho</h2>

<label>Data</label>

<input
id="trabalhoData"
type="date">

<label>Título</label>

<input
id="trabalhoTitulo"
placeholder="Ex: Trabalho de História">

<label>Destino do trabalho</label>

<select id="trabalhoDestino">

<option value="ALUNO">
Apenas um aluno
</option>

<option value="TODOS">
Todos os alunos da minha Série/Ano
</option>

</select>

<div id="trabalhoAlunoBox">

<label>Aluno</label>

<select id="trabalhoAluno">

<option value="">
Selecione o aluno
</option>

</select>

</div>

<label>Matéria</label>

<input
id="trabalhoMateria"
placeholder="Ex: História">

<label>Nota</label>

<input
id="trabalhoNota"
type="number"
min="0"
max="10"
step="0.1">

<label>Descrição</label>

<textarea
id="trabalhoDescricao"
placeholder="Descrição do trabalho">
</textarea>

<button
id="trabalhoSalvar"
class="btn btn-primary">
💾 Publicar trabalho
</button>

</div>

<div class="card">

<div class="card-header">

<h2>📂 Trabalhos</h2>

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


<!-- ==========================================================
LANÇAR NOTAS - PROFESSOR
========================================================== -->

<section
id="screenLancarNotas"
class="screen">

<div class="card">

<div class="card-header">

<div>

<h2>📊 Lançar Notas</h2>

<p style="color:#64748b;margin-top:5px;">
Lance ou altere a média dos alunos da sua Série/Ano.
</p>

</div>

<button
id="mediaProfessorAtualizar"
class="btn btn-gray">
🔄
</button>

</div>

<div class="media-info">
🔐 <strong>Área exclusiva do professor.</strong>
Apenas professores podem editar as médias.
Os alunos somente poderão visualizar suas próprias notas.
</div>

<div class="media-toolbar">

<div>

<label>Matéria</label>

<select id="mediaMateriaSelect">

<option value="">
Carregando matérias...
</option>

</select>

</div>

<div>

<button
id="carregarAlunosMediasBtn"
class="btn btn-primary">
👨‍🎓 Carregar alunos
</button>

</div>

</div>

</div>


<div class="card">

<div class="card-header">

<div>

<h2>👨‍🎓 Alunos e Médias</h2>

<p style="color:#64748b;margin-top:5px;">
Preencha as médias e clique em salvar.
</p>

</div>

</div>

<div
id="mediasProfessorGrid">

<div class="media-loading">
Selecione uma matéria para começar.
</div>

</div>

</div>

</section>


<!-- ==========================================================
MINHAS NOTAS - ALUNO
========================================================== -->

<section
id="screenMinhasNotas"
class="screen">

<div class="card">

<div class="card-header">

<div>

<h2>📊 Minhas Notas</h2>

<p style="color:#64748b;margin-top:5px;">
Consulte suas médias por matéria.
</p>

</div>

<button
id="minhasNotasAtualizar"
class="btn btn-gray">
🔄
</button>

</div>

<div class="media-info">
👨‍🎓 <strong>Área de consulta.</strong>
Você pode visualizar suas médias, mas não pode alterá-las.
</div>

<div
id="mediasAlunoGrid"
class="grid">

<div class="media-loading">
Carregando suas notas...
</div>

</div>

</div>

</section>


<!-- ==========================================================
DÚVIDAS
========================================================== -->

<section
id="screenDuvidas"
class="screen">

<div
id="duvidaAlunoForm"
class="card">

<h2>
❓ Enviar uma dúvida
</h2>

<label>
Professor
</label>

<select id="duvidaProfessor">

<option value="">
Selecione o professor
</option>

</select>

<label>
Assunto
</label>

<input
id="duvidaAssunto"
placeholder="Ex: Dúvida sobre a prova">

<label>
Dúvida
</label>

<textarea
id="duvidaMensagem"
placeholder="Digite sua dúvida para o professor...">
</textarea>

<button
id="duvidaEnviar"
class="btn btn-primary">
📨 Enviar dúvida
</button>

</div>

<div class="card">

<div class="card-header">

<div>

<h2>
❓ Dúvidas
</h2>

<p>
Perguntas enviadas pelos alunos.
</p>

</div>

<button
id="duvidaAtualizar"
class="btn btn-gray">
🔄
</button>

</div>

<div class="filtro-duvidas">

<button
id="filtroTodas"
class="active"
onclick="mudarFiltroDuvidas('TODAS')">
Todas
</button>

<button
id="filtroNaoLidas"
onclick="mudarFiltroDuvidas('NAO_LIDAS')">
🔵 Não lidas
</button>

<button
id="filtroLidas"
onclick="mudarFiltroDuvidas('LIDAS')">
🟢 Lidas
</button>

</div>

<br>

<div
id="duvidasGrid"
class="grid">
</div>

</div>

</section>


<!-- ==========================================================
ALUNOS
========================================================== -->

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
let forgotTipo = "ALUNO";

let token = null;
let usuario = null;

let filtroDuvidas = "TODAS";


// ==========================================================
// UTILIDADES
// ==========================================================

function el(id) {
  return document.getElementById(id);
}


function mensagem(id, texto, sucesso) {

  const box = el(id);

  if (!box)
    return;

  box.textContent = texto;

  box.className =
    "message show " +
    (
      sucesso
        ? "success"
        : "error"
    );
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
    .classList
    .add("active");

  el("loginProfessor")
    .classList
    .remove("active");
};


el("loginProfessor").onclick =
function() {

  loginTipo = "PROFESSOR";

  el("loginProfessor")
    .classList
    .add("active");

  el("loginAluno")
    .classList
    .remove("active");
};


// ==========================================================
// ABRIR RECUPERAÇÃO
// ==========================================================

el("esqueciSenhaBtn").onclick =
function() {

  el("loginPage")
    .classList
    .add("hidden");

  el("registerPage")
    .classList
    .add("hidden");

  el("forgotPasswordPage")
    .classList
    .remove("hidden");

  el("forgotRA").value = "";
  el("forgotUsuario").value = "";
  el("forgotNovaSenha").value = "";
  el("forgotConfirmarSenha").value = "";

  el("forgotMsg").className =
    "message";

  el("forgotMsg").textContent = "";
};


// ==========================================================
// RECUPERAÇÃO - ALUNO
// ==========================================================

el("forgotAluno").onclick =
function() {

  forgotTipo = "ALUNO";

  el("forgotAluno")
    .classList
    .add("active");

  el("forgotProfessor")
    .classList
    .remove("active");
};


// ==========================================================
// RECUPERAÇÃO - PROFESSOR
// ==========================================================

el("forgotProfessor").onclick =
function() {

  forgotTipo = "PROFESSOR";

  el("forgotProfessor")
    .classList
    .add("active");

  el("forgotAluno")
    .classList
    .remove("active");
};


// ==========================================================
// VOLTAR PARA LOGIN
// ==========================================================

el("voltarLoginSenha").onclick =
function() {

  el("forgotPasswordPage")
    .classList
    .add("hidden");

  el("loginPage")
    .classList
    .remove("hidden");

  el("forgotMsg").className =
    "message";

  el("forgotMsg").textContent = "";
};


// ==========================================================
// ALTERAR SENHA
// ==========================================================

el("alterarSenhaBtn").onclick =
function() {

  const ra =
    el("forgotRA")
      .value
      .trim();

  const usuarioDigitado =
    el("forgotUsuario")
      .value
      .trim();

  const novaSenha =
    el("forgotNovaSenha")
      .value;

  const confirmarSenha =
    el("forgotConfirmarSenha")
      .value;


  if (!ra) {

    mensagem(
      "forgotMsg",
      "Informe o RA/ID.",
      false
    );

    return;
  }


  if (!usuarioDigitado) {

    mensagem(
      "forgotMsg",
      "Informe o usuário.",
      false
    );

    return;
  }


  if (novaSenha.length < 6) {

    mensagem(
      "forgotMsg",
      "A nova senha precisa ter pelo menos 6 caracteres.",
      false
    );

    return;
  }


  if (novaSenha !== confirmarSenha) {

    mensagem(
      "forgotMsg",
      "As senhas não coincidem.",
      false
    );

    return;
  }


  mensagem(
    "forgotMsg",
    "Alterando senha...",
    true
  );


  google.script.run

    .withSuccessHandler(
      function(resultado) {

        if (!resultado) {

          mensagem(
            "forgotMsg",
            "Não foi possível alterar a senha.",
            false
          );

          return;
        }


        mensagem(
          "forgotMsg",
          resultado.mensagem,
          resultado.sucesso
        );


        if (resultado.sucesso) {

          el("forgotRA").value = "";
          el("forgotUsuario").value = "";
          el("forgotNovaSenha").value = "";
          el("forgotConfirmarSenha").value = "";


          setTimeout(
            function() {

              el("forgotPasswordPage")
                .classList
                .add("hidden");

              el("loginPage")
                .classList
                .remove("hidden");

              el("loginMsg")
                .className =
                "message";

              el("loginMsg")
                .textContent = "";

            },
            1500
          );
        }
      }
    )

    .withFailureHandler(
      function(error) {

        mensagem(
          "forgotMsg",
          error &&
          error.message
            ? error.message
            : "Erro ao alterar a senha.",
          false
        );
      }
    )

    .recuperarSenha({

      tipo:
        forgotTipo,

      ra:
        ra,

      usuario:
        usuarioDigitado,

      novaSenha:
        novaSenha
    });
};


// ==========================================================
// CADASTRO
// ==========================================================

el("cadastroBtn").onclick =
function() {

  el("loginPage")
    .classList
    .add("hidden");

  el("registerPage")
    .classList
    .remove("hidden");
};


el("voltarLogin").onclick =
function() {

  el("registerPage")
    .classList
    .add("hidden");

  el("loginPage")
    .classList
    .remove("hidden");
};


el("cadAluno").onclick =
function() {

  cadastroTipo = "ALUNO";

  el("cadAluno")
    .classList
    .add("active");

  el("cadProfessor")
    .classList
    .remove("active");

  el("cadMateriaBox")
    .classList
    .add("hidden");

  el("cadTurmaBox")
    .classList
    .remove("hidden");
};


el("cadProfessor").onclick =
function() {

  cadastroTipo = "PROFESSOR";

  el("cadProfessor")
    .classList
    .add("active");

  el("cadAluno")
    .classList
    .remove("active");

  el("cadMateriaBox")
    .classList
    .remove("hidden");

  el("cadTurmaBox")
    .classList
    .add("hidden");
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

    turma:
      el("cadTurma").value.trim(),

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
    cadastroTipo === "ALUNO" &&
    !dados.turma
  ) {

    mensagem(
      "cadMsg",
      "Informe a Turma.",
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

          setTimeout(
            function() {

              el("registerPage")
                .classList
                .add("hidden");

              el("loginPage")
                .classList
                .remove("hidden");

            },
            1000
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

    .cadastrarUsuario(
      dados
    );
};


// ==========================================================
// LOGIN
// ==========================================================

el("entrarBtn").onclick =
fazerLogin;


el("loginSenha").onkeydown =
function(event) {

  if (
    event.key === "Enter"
  ) {

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

        if (
          !resultado.sucesso
        ) {

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
// APP
// ==========================================================

function abrirApp() {

  el("loginPage")
    .classList
    .add("hidden");

  el("registerPage")
    .classList
    .add("hidden");

  el("forgotPasswordPage")
    .classList
    .add("hidden");

  el("app")
    .style.display =
    "block";


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


  const professor =
    usuario.tipo ===
    "PROFESSOR";


  if (professor) {

    el("heroText")
      .textContent =
      "Gerencie matérias, provas, trabalhos, notas e responda dúvidas.";

  } else {

    el("heroText")
      .textContent =
      "Veja suas matérias, provas, trabalhos, notas e tire suas dúvidas.";
  }


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


  el("duvidaAlunoForm")
    .style.display =
    professor
      ? "none"
      : "block";


  // NOVO
  el("lancarNotasNav")
    .style.display =
    professor
      ? "block"
      : "none";


  // NOVO
  el("minhasNotasNav")
    .style.display =
    professor
      ? "none"
      : "block";


  carregarMaterias();
  carregarProvas();
  carregarTrabalhos();
  carregarDuvidas();

  if (professor) {

    carregarAlunos();

    carregarMateriasParaMedias();

  } else {

    carregarProfessores();

    carregarMinhasNotas();

  }

  carregarEstatisticas();
}


// ==========================================================
// ESTATÍSTICAS
// ==========================================================

function carregarEstatisticas() {

  google.script.run

    .withSuccessHandler(
      function(stats) {

        el("materiaCount")
          .textContent =
          stats.materias;

        el("provaCount")
          .textContent =
          stats.provas;

        el("trabalhoCount")
          .textContent =
          stats.trabalhos;

        el("alunoCount")
          .textContent =
          stats.alunos;
      }
    )

    .withFailureHandler(
      mostrarErro
    )

    .obterEstatisticas(
      token
    );
}


// ==========================================================
// NAVEGAÇÃO
// ==========================================================

document
.querySelectorAll(
  ".nav[data-screen]"
)
.forEach(
  function(button) {

    button.onclick =
    function() {

      document
      .querySelectorAll(
        ".screen"
      )
      .forEach(
        function(screen) {

          screen
            .classList
            .remove("active");
        }
      );


      const nome =
        this.dataset.screen;


      const id =
        "screen" +
        nome
          .charAt(0)
          .toUpperCase() +
        nome.slice(1);


      const tela =
        el(id);


      if (!tela)
        return;


      tela
        .classList
        .add("active");


      document
      .querySelectorAll(
        ".nav"
      )
      .forEach(
        function(nav) {

          nav.classList
            .remove("active");
        }
      );


      this.classList
        .add("active");


      const titulos = {

        dashboard:
          "Dashboard",

        materias:
          "Matérias",

        provas:
          "Provas e Notas",

        trabalhos:
          "Trabalhos",

        lancarNotas:
          "Lançar Notas",

        minhasNotas:
          "Minhas Notas",

        duvidas:
          "Dúvidas",

        alunos:
          "Alunos"
      };


      el("pageTitle")
        .textContent =
        titulos[nome] ||
        "Gestão Escolar";


      if (
        nome === "materias"
      )
        carregarMaterias();


      if (
        nome === "provas"
      ) {

        carregarProvas();

        if (
          usuario.tipo ===
          "PROFESSOR"
        ) {

          carregarAlunos();
        }
      }


      if (
        nome === "trabalhos"
      ) {

        carregarTrabalhos();

        if (
          usuario.tipo ===
          "PROFESSOR"
        ) {

          carregarAlunos();
        }
      }


      if (
        nome === "lancarNotas"
      ) {

        if (
          usuario.tipo ===
          "PROFESSOR"
        ) {

          carregarMateriasParaMedias();

          setTimeout(
            function() {

              carregarAlunosParaMedias();

            },
            100
          );
        }
      }


      if (
        nome === "minhasNotas"
      ) {

        if (
          usuario.tipo ===
          "ALUNO"
        ) {

          carregarMinhasNotas();
        }
      }


      if (
        nome === "duvidas"
      ) {

        carregarDuvidas();

        if (
          usuario.tipo ===
          "ALUNO"
        ) {

          carregarProfessores();
        }
      }


      if (
        nome === "alunos"
      ) {

        carregarAlunos();
      }


      // Fecha o menu mobile após navegar
      el("sidebar")
        .classList
        .remove("open");

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

        // NOVO
        if (
          usuario &&
          usuario.tipo ===
          "PROFESSOR"
        ) {

          carregarMateriasParaMedias();
        }
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

        const grid =
          el("materiasGrid");

        grid.innerHTML = "";


        if (!materias.length) {

          grid.innerHTML =
            '<div class="empty">' +
            '📚<br><br>Nenhuma matéria encontrada.' +
            '</div>';

          return;
        }


        materias.forEach(
          function(materia) {

            let linksHTML = "";

            const links =
              String(
                materia.links || ""
              )
              .split(/\n|,/)
              .map(
                x => x.trim()
              )
              .filter(Boolean);


            if (links.length) {

              linksHTML =
                '<div class="links">' +
                '<strong>🔗 Materiais</strong>' +

                links.map(
                  function(link) {

                    const url =
                      /^https?:\/\//i
                        .test(link)
                        ? link
                        : "https://" + link;

                    return (
                      '<a class="study-link" href="' +
                      escapeHTML(url) +
                      '" target="_blank">' +
                      '🌐 ' +
                      escapeHTML(link) +
                      '</a>'
                    );
                  }
                ).join("") +

                '</div>';
            }


            let botoes = "";


            if (
              usuario.tipo ===
              "PROFESSOR" &&

              String(
                materia.professorRA
              ) ===
              String(
                usuario.ra
              )
            ) {

              botoes =
                '<div style="margin-top:15px;display:flex;gap:8px">' +

                '<button class="btn btn-primary" ' +
                'onclick="editarMateria(\'' +
                escapeHTML(materia.linha) +
                '\')">' +
                '✏️ Editar' +
                '</button>' +

                '<button class="btn btn-red" ' +
                'onclick="excluirMateria(\'' +
                escapeHTML(materia.linha) +
                '\')">' +
                '🗑 Excluir' +
                '</button>' +

                '</div>';
            }


            const card =
              document.createElement(
                "div"
              );

            card.className =
              "subject";


            card.innerHTML =

              '<div class="subject-head">' +

              '<h3>📚 ' +
              escapeHTML(materia.nome) +
              '</h3>' +

              '<div class="subject-meta">' +
              'Professor: ' +
              escapeHTML(materia.professor) +
              '<br>' +
              'Série/Ano: ' +
              escapeHTML(materia.serieAno) +
              '</div>' +

              '</div>' +

              '<div class="subject-body">' +

              '<div class="subject-content">' +
              '<strong>📖 Conteúdo</strong>' +
              '<br><br>' +
              escapeHTML(
                materia.conteudo ||
                "Nenhum conteúdo cadastrado."
              ) +
              '</div>' +

              linksHTML +

              botoes +

              '</div>';


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
      "Conteúdo:"
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

        if (
          usuario &&
          usuario.tipo ===
          "PROFESSOR"
        ) {

          carregarMateriasParaMedias();
        }
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
  )
    return;


  google.script.run

    .withSuccessHandler(
      function(resultado) {

        alert(
          resultado.mensagem
        );

        carregarMaterias();

        if (
          usuario &&
          usuario.tipo ===
          "PROFESSOR"
        ) {

          carregarMateriasParaMedias();
        }
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

  if (
    usuario.tipo !==
    "PROFESSOR"
  )
    return;


  google.script.run

    .withSuccessHandler(
      function(alunos) {

        el("alunoCount")
          .textContent =
          alunos.length;


        preencherSelectAlunos(
          "provaAluno",
          alunos
        );

        preencherSelectAlunos(
          "trabalhoAluno",
          alunos
        );


        const grid =
          el("alunosGrid");

        grid.innerHTML = "";


        if (!alunos.length) {

          grid.innerHTML =
            '<div class="empty">' +
            'Nenhum aluno cadastrado nesta Série/Ano.' +
            '</div>';

          return;
        }


        alunos.forEach(
          function(aluno) {

            const card =
              document.createElement(
                "div"
              );

            card.className =
              "item";


            card.innerHTML =

              '<h3>👨‍🎓 ' +
              escapeHTML(aluno.nome) +
              '</h3>' +

              '<p><strong>RA:</strong> ' +
              escapeHTML(aluno.ra) +
              '</p>' +

              '<p><strong>Série/Ano:</strong> ' +
              escapeHTML(aluno.serieAno) +
              '</p>' +

              '<p><strong>Turma:</strong> ' +
              escapeHTML(aluno.turma || "-") +
              '</p>';


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


function preencherSelectAlunos(
  id,
  alunos
) {

  const select =
    el(id);

  if (!select)
    return;


  select.innerHTML =
    '<option value="">Selecione o aluno</option>';


  alunos.forEach(
    function(aluno) {

      const option =
        document.createElement(
          "option"
        );

      option.value =
        aluno.ra;

      option.textContent =
        aluno.nome +
        " — RA: " +
        aluno.ra;

      select.appendChild(
        option
      );
    }
  );
}


el("alunosAtualizar").onclick =
carregarAlunos;


// ==========================================================
// PROVAS
// ==========================================================

el("provaDestino").onchange =
function() {

  el("provaAlunoBox")
    .style.display =
    this.value === "TODOS"
      ? "none"
      : "block";
};


el("provaSalvar").onclick =
function() {

  const todos =
    el("provaDestino")
      .value === "TODOS";


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

    todosSerie:
      todos,

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

    alert(
      "Informe a data."
    );

    return;
  }


  if (!dados.titulo) {

    alert(
      "Informe o título."
    );

    return;
  }


  if (
    !todos &&
    !dados.alunoRA
  ) {

    alert(
      "Selecione o aluno."
    );

    return;
  }


  if (!dados.materia) {

    alert(
      "Informe a matéria."
    );

    return;
  }


  google.script.run

    .withSuccessHandler(
      function(resultado) {

        alert(
          resultado.mensagem
        );

        limparFormularioProva();

        carregarProvas();

        carregarEstatisticas();
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


function limparFormularioProva() {

  el("provaData").value = "";
  el("provaTitulo").value = "";
  el("provaAluno").value = "";
  el("provaMateria").value = "";
  el("provaNota").value = "";
  el("provaDescricao").value = "";
}


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
            '📝<br><br>Nenhuma prova encontrada.' +
            '</div>';

          return;
        }


        provas.forEach(
          function(prova) {

            const card =
              document.createElement(
                "div"
              );

            card.className =
              "item";


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
                '🗑 Excluir' +
                '</button>';
            }


            const nota =
              prova.nota !== ""
                ? '<span class="note">⭐ Nota: ' +
                  escapeHTML(prova.nota) +
                  '</span>'
                : '<span class="note">⏳ Nota não lançada</span>';


            card.innerHTML =

              '<h3>📝 ' +
              escapeHTML(prova.titulo) +
              '</h3>' +

              '<p><strong>Data:</strong> ' +
              escapeHTML(prova.data) +
              '</p>' +

              '<p><strong>Aluno:</strong> ' +
              escapeHTML(prova.aluno) +
              '</p>' +

              '<p><strong>RA:</strong> ' +
              escapeHTML(prova.alunoRA) +
              '</p>' +

              '<p><strong>Matéria:</strong> ' +
              escapeHTML(prova.materia) +
              '</p>' +

              '<p><strong>Professor:</strong> ' +
              escapeHTML(prova.professor) +
              '</p>' +

              nota +

              '<p><strong>Descrição:</strong> ' +
              escapeHTML(
                prova.descricao || "-"
              ) +
              '</p>' +

              botao;


            grid.appendChild(card);
          }
        );
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
  )
    return;


  google.script.run

    .withSuccessHandler(
      function(resultado) {

        alert(
          resultado.mensagem
        );

        carregarProvas();

        carregarEstatisticas();
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

el("trabalhoDestino").onchange =
function() {

  el("trabalhoAlunoBox")
    .style.display =
    this.value === "TODOS"
      ? "none"
      : "block";
};


el("trabalhoSalvar").onclick =
function() {

  const todos =
    el("trabalhoDestino")
      .value === "TODOS";


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

    todosSerie:
      todos,

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

    alert(
      "Informe a data."
    );

    return;
  }


  if (!dados.titulo) {

    alert(
      "Informe o título."
    );

    return;
  }


  if (
    !todos &&
    !dados.alunoRA
  ) {

    alert(
      "Selecione o aluno."
    );

    return;
  }


  if (!dados.materia) {

    alert(
      "Informe a matéria."
    );

    return;
  }


  google.script.run

    .withSuccessHandler(
      function(resultado) {

        alert(
          resultado.mensagem
        );

        limparFormularioTrabalho();

        carregarTrabalhos();

        carregarEstatisticas();
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


function limparFormularioTrabalho() {

  el("trabalhoData").value = "";
  el("trabalhoTitulo").value = "";
  el("trabalhoAluno").value = "";
  el("trabalhoMateria").value = "";
  el("trabalhoNota").value = "";
  el("trabalhoDescricao").value = "";
}


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
            '📂<br><br>Nenhum trabalho encontrado.' +
            '</div>';

          return;
        }


        trabalhos.forEach(
          function(trabalho) {

            const card =
              document.createElement(
                "div"
              );

            card.className =
              "item";


            const nota =
              trabalho.nota !== ""
                ? '<span class="note">⭐ Nota: ' +
                  escapeHTML(trabalho.nota) +
                  '</span>'
                : '<span class="note">⏳ Nota não lançada</span>';


            let botao = "";


            if (
              usuario.tipo ===
              "PROFESSOR"
            ) {

              botao =
                '<br><button class="btn btn-red" ' +
                'onclick="excluirTrabalho(\'' +
                escapeHTML(trabalho.linha) +
                '\')">' +
                '🗑 Excluir' +
                '</button>';
            }


            card.innerHTML =

              '<h3>📂 ' +
              escapeHTML(trabalho.titulo) +
              '</h3>' +

              '<p><strong>Data:</strong> ' +
              escapeHTML(trabalho.data) +
              '</p>' +

              '<p><strong>Aluno:</strong> ' +
              escapeHTML(trabalho.aluno) +
              '</p>' +

              '<p><strong>Matéria:</strong> ' +
              escapeHTML(trabalho.materia) +
              '</p>' +

              '<p><strong>Professor:</strong> ' +
              escapeHTML(trabalho.professor) +
              '</p>' +

              nota +

              '<p><strong>Descrição:</strong> ' +
              escapeHTML(
                trabalho.descricao || "-"
              ) +
              '</p>' +

              botao;


            grid.appendChild(card);
          }
        );
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
  )
    return;


  google.script.run

    .withSuccessHandler(
      function(resultado) {

        alert(
          resultado.mensagem
        );

        carregarTrabalhos();

        carregarEstatisticas();
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
// MÉDIAS / NOTAS - PROFESSOR
// ==========================================================

function carregarMateriasParaMedias() {

  if (
    !usuario ||
    usuario.tipo !==
    "PROFESSOR"
  )
    return;


  const select =
    el("mediaMateriaSelect");


  if (!select)
    return;


  select.innerHTML =
    '<option value="">Carregando matérias...</option>';


  google.script.run

    .withSuccessHandler(
      function(materias) {

        select.innerHTML =
          '<option value="">Selecione a matéria</option>';


        if (
          !materias ||
          !materias.length
        ) {

          select.innerHTML =
            '<option value="">Nenhuma matéria encontrada</option>';


          el("mediasProfessorGrid")
            .innerHTML =
            '<div class="empty">' +
            '📚<br><br>Nenhuma matéria disponível para lançamento de notas.' +
            '</div>';

          return;
        }


        materias.forEach(
          function(materia) {

            const option =
              document.createElement(
                "option"
              );

            option.value =
              materia;

            option.textContent =
              materia;

            select.appendChild(
              option
            );
          }
        );


        // Seleciona a primeira matéria automaticamente
        if (!select.value) {

          select.selectedIndex = 1;
        }


        carregarAlunosParaMedias();
      }
    )

    .withFailureHandler(
      mostrarErro
    )

    .listarMateriasParaMedias(
      token
    );
}


el("mediaMateriaSelect").onchange =
function() {

  if (
    !this.value
  ) {

    el("mediasProfessorGrid")
      .innerHTML =
      '<div class="empty">' +
      '📚<br><br>Selecione uma matéria.' +
      '</div>';

    return;
  }


  carregarAlunosParaMedias();
};


el("carregarAlunosMediasBtn").onclick =
carregarAlunosParaMedias;


el("mediaProfessorAtualizar").onclick =
function() {

  carregarMateriasParaMedias();

};


function carregarAlunosParaMedias() {

  if (
    !usuario ||
    usuario.tipo !==
    "PROFESSOR"
  )
    return;


  const materia =
    el("mediaMateriaSelect")
      .value;


  if (!materia) {

    el("mediasProfessorGrid")
      .innerHTML =
      '<div class="empty">' +
      '📚<br><br>Selecione uma matéria para carregar os alunos.' +
      '</div>';

    return;
  }


  el("mediasProfessorGrid")
    .innerHTML =
    '<div class="media-loading">' +
    '⏳ Carregando alunos...' +
    '</div>';


  google.script.run

    .withSuccessHandler(
      function(alunos) {

        renderizarAlunosParaMedias(
          alunos
        );
      }
    )

    .withFailureHandler(
      mostrarErro
    )

    .listarAlunosParaMedias(
      token,
      materia
    );
}


function renderizarAlunosParaMedias(
  alunos
) {

  const container =
    el("mediasProfessorGrid");


  container.innerHTML = "";


  if (
    !alunos ||
    !alunos.length
  ) {

    container.innerHTML =
      '<div class="empty">' +
      '👨‍🎓<br><br>Nenhum aluno encontrado nesta Série/Ano.' +
      '</div>';

    return;
  }


  const tabela =
    document.createElement(
      "div"
    );

  tabela.className =
    "media-table-wrap";


  let html =

    '<table class="media-table">' +

    '<thead>' +

    '<tr>' +

    '<th>Aluno</th>' +

    '<th>RA</th>' +

    '<th>Turma</th>' +

    '<th>Média</th>' +

    '<th>Status</th>' +

    '</tr>' +

    '</thead>' +

    '<tbody>';


  alunos.forEach(
    function(aluno, index) {

      const valor =
        aluno.media || "";


      const status =
        valor !== ""
          ? '<span class="media-status lancada">✅ Lançada</span>'
          : '<span class="media-status pendente">⏳ Pendente</span>';


      html +=

        '<tr>' +

        '<td class="media-name">' +
        escapeHTML(aluno.nome) +
        '</td>' +

        '<td>' +
        escapeHTML(aluno.ra) +
        '</td>' +

        '<td>' +
        escapeHTML(aluno.turma || "-") +
        '</td>' +

        '<td>' +

        '<input ' +
        'type="text" ' +
        'inputmode="decimal" ' +
        'class="media-input" ' +
        'data-ra="' +
        escapeHTML(aluno.ra) +
        '" ' +
        'value="' +
        escapeHTML(valor) +
        '" ' +
        'placeholder="0 a 10" ' +
        'maxlength="4">' +

        '</td>' +

        '<td>' +
        status +
        '</td>' +

        '</tr>';
    }
  );


  html +=
    '</tbody>' +
    '</table>';


  tabela.innerHTML =
    html;


  container.appendChild(
    tabela
  );


  const acoes =
    document.createElement(
      "div"
    );

  acoes.className =
    "media-acoes";


  const salvarBtn =
    document.createElement(
      "button"
    );

  salvarBtn.className =
    "btn btn-green";

  salvarBtn.textContent =
    "💾 Salvar todas as médias";


  salvarBtn.onclick =
    function() {

      salvarTodasAsMedias();
    };


  const limparBtn =
    document.createElement(
      "button"
    );

  limparBtn.className =
    "btn btn-gray";

  limparBtn.textContent =
    "🧹 Limpar campos";


  limparBtn.onclick =
    function() {

      container
        .querySelectorAll(
          ".media-input"
        )
        .forEach(
          function(input) {

            input.value = "";
          }
        );
    };


  acoes.appendChild(
    salvarBtn
  );

  acoes.appendChild(
    limparBtn
  );


  container.appendChild(
    acoes
  );
}


function salvarTodasAsMedias() {

  if (
    !usuario ||
    usuario.tipo !==
    "PROFESSOR"
  ) {

    mostrarErro({
      message:
        "Apenas professores podem lançar médias."
    });

    return;
  }


  const materia =
    el("mediaMateriaSelect")
      .value;


  if (!materia) {

    alert(
      "Selecione a matéria."
    );

    return;
  }


  const inputs =
    document
      .querySelectorAll(
        "#mediasProfessorGrid .media-input"
      );


  if (!inputs.length) {

    alert(
      "Nenhum aluno foi carregado."
    );

    return;
  }


  const medias = [];


  let erro =
    false;


  inputs.forEach(
    function(input) {

      if (erro)
        return;


      const valor =
        input.value
          .trim()
          .replace(",", ".");


      if (valor !== "") {

        const numero =
          Number(valor);


        if (
          isNaN(numero) ||
          numero < 0 ||
          numero > 10
        ) {

          erro = true;

          input.focus();

          alert(
            "A média deve ser um número entre 0 e 10."
          );

          return;
        }
      }


      medias.push({

        alunoRA:
          input.dataset.ra,

        media:
          valor
      });
    }
  );


  if (erro)
    return;


  const confirmar =
    confirm(
      "Deseja salvar as médias desta matéria?"
    );


  if (!confirmar)
    return;


  const botao =
    document.querySelector(
      "#mediasProfessorGrid .btn-green"
    );


  if (botao) {

    botao.disabled =
      true;

    botao.textContent =
      "⏳ Salvando...";
  }


  google.script.run

    .withSuccessHandler(
      function(resultado) {

        if (botao) {

          botao.disabled =
            false;

          botao.textContent =
            "💾 Salvar todas as médias";
        }


        alert(
          resultado.mensagem
        );


        carregarAlunosParaMedias();
      }
    )

    .withFailureHandler(
      function(error) {

        if (botao) {

          botao.disabled =
            false;

          botao.textContent =
            "💾 Salvar todas as médias";
        }


        mostrarErro(error);
      }
    )

    .salvarMedias(
      token,
      materia,
      medias
    );
}


// ==========================================================
// MÉDIAS / NOTAS - ALUNO
// ==========================================================

el("minhasNotasAtualizar").onclick =
carregarMinhasNotas;


function carregarMinhasNotas() {

  if (
    !usuario ||
    usuario.tipo !==
    "ALUNO"
  )
    return;


  const grid =
    el("mediasAlunoGrid");


  if (!grid)
    return;


  grid.innerHTML =
    '<div class="media-loading">' +
    '⏳ Carregando suas notas...' +
    '</div>';


  google.script.run

    .withSuccessHandler(
      function(medias) {

        renderizarMediasAluno(
          medias
        );
      }
    )

    .withFailureHandler(
      mostrarErro
    )

    .listarMediasAluno(
      token
    );
}


function renderizarMediasAluno(
  medias
) {

  const grid =
    el("mediasAlunoGrid");


  grid.innerHTML = "";


  if (
    !medias ||
    !medias.length
  ) {

    grid.innerHTML =
      '<div class="empty">' +
      '📊<br><br>Nenhuma média foi lançada para você ainda.' +
      '</div>';

    return;
  }


  medias.forEach(
    function(item) {

      const card =
        document.createElement(
          "div"
        );

      card.className =
        "media-aluno-card";


      const valor =
        item.media !== ""
          ? escapeHTML(item.media)
          : "Não lançada";


      const classe =
        item.media !== ""
          ? "media-valor-grande"
          : "";


      card.innerHTML =

        '<div class="media-aluno-card-head">' +

        '<h3>📚 ' +
        escapeHTML(item.materia) +
        '</h3>' +

        '<p style="color:#64748b;font-size:12px">' +
        'Professor: ' +
        escapeHTML(item.professor || "-") +
        '</p>' +

        '</div>' +

        '<div class="media-aluno-card-body">' +

        '<p><strong>Série/Ano:</strong> ' +
        escapeHTML(item.serieAno || "-") +
        '</p>' +

        '<p><strong>RA:</strong> ' +
        escapeHTML(item.alunoRA || usuario.ra) +
        '</p>' +

        '<p><strong>Média:</strong></p>' +

        '<div class="' +
        classe +
        '">' +
        valor +
        '</div>' +

        '<p style="margin-top:12px;color:#64748b;font-size:12px">' +
        '<strong>Atualizada em:</strong> ' +
        escapeHTML(item.dataAtualizacao || "-") +
        '</p>' +

        '</div>';


      grid.appendChild(
        card
      );
    }
  );
}


// ==========================================================
// PROFESSORES
// ==========================================================

function carregarProfessores() {

  if (
    usuario.tipo !==
    "ALUNO"
  )
    return;


  google.script.run

    .withSuccessHandler(
      function(professores) {

        const select =
          el("duvidaProfessor");

        select.innerHTML =
          '<option value="">Selecione o professor</option>';


        professores.forEach(
          function(professor) {

            const option =
              document.createElement(
                "option"
              );

            option.value =
              professor.ra;

            option.textContent =
              professor.nome +
              " — " +
              professor.materia;

            select.appendChild(
              option
            );
          }
        );
      }
    )

    .withFailureHandler(
      mostrarErro
    )

    .listarProfessores(token);
}


// ==========================================================
// ENVIAR DÚVIDA
// ==========================================================

el("duvidaEnviar").onclick =
function() {

  const dados = {

    professorRA:
      el("duvidaProfessor")
        .value,

    assunto:
      el("duvidaAssunto")
        .value.trim(),

    mensagem:
      el("duvidaMensagem")
        .value.trim()
  };


  if (!dados.professorRA) {

    alert(
      "Selecione o professor."
    );

    return;
  }


  if (!dados.assunto) {

    alert(
      "Informe o assunto."
    );

    return;
  }


  if (!dados.mensagem) {

    alert(
      "Digite sua dúvida."
    );

    return;
  }


  google.script.run

    .withSuccessHandler(
      function(resultado) {

        alert(
          resultado.mensagem
        );

        el("duvidaAssunto")
          .value = "";

        el("duvidaMensagem")
          .value = "";

        carregarDuvidas();
      }
    )

    .withFailureHandler(
      mostrarErro
    )

    .cadastrarDuvida(
      token,
      dados
    );
};


// ==========================================================
// FILTRO DÚVIDAS
// ==========================================================

function mudarFiltroDuvidas(
  filtro
) {

  filtroDuvidas =
    filtro;


  document
    .querySelectorAll(
      ".filtro-duvidas button"
    )
    .forEach(
      function(btn) {

        btn.classList
          .remove("active");
      }
    );


  if (
    filtro === "TODAS"
  )
    el("filtroTodas")
      .classList
      .add("active");


  if (
    filtro === "NAO_LIDAS"
  )
    el("filtroNaoLidas")
      .classList
      .add("active");


  if (
    filtro === "LIDAS"
  )
    el("filtroLidas")
      .classList
      .add("active");


  carregarDuvidas();
}


el("duvidaAtualizar").onclick =
carregarDuvidas;


// ==========================================================
// CARREGAR DÚVIDAS
// ==========================================================

function carregarDuvidas() {

  google.script.run

    .withSuccessHandler(
      function(duvidas) {

        const grid =
          el("duvidasGrid");

        grid.innerHTML = "";


        if (!duvidas.length) {

          grid.innerHTML =
            '<div class="empty">' +
            '❓<br><br>Nenhuma dúvida encontrada.' +
            '</div>';

          return;
        }


        duvidas.forEach(
          function(duvida) {

            const professor =
              usuario.tipo ===
              "PROFESSOR";


            const status =
              professor
                ? duvida.statusProfessor
                : duvida.statusAluno;


            const naoLida =
              status ===
              "NAO_LIDA";


            const card =
              document.createElement(
                "div"
              );


            card.className =
              "duvida " +
              (
                naoLida
                  ? "duvida-nao-lida"
                  : "duvida-lida"
              );


            const statusHTML =
              naoLida

                ? '<span class="status status-nao">🔵 Não lida</span>'

                : '<span class="status status-sim">🟢 Lida</span>';


            let respostaHTML = "";


            if (
              duvida.resposta
            ) {

              respostaHTML =
                '<div class="resposta">' +

                '<strong>👨‍🏫 Resposta do professor:</strong>' +

                '<br><br>' +

                escapeHTML(
                  duvida.resposta
                ) +

                '</div>';
            }


            let botoes = "";


            if (naoLida) {

              botoes +=

                '<button class="btn btn-green" ' +

                'onclick="marcarDuvidaLida(\'' +

                escapeHTML(
                  duvida.id
                ) +

                '\')">' +

                '✓ Marcar como lida' +

                '</button>';
            }


            if (professor) {

              botoes +=

                '<div style="margin-top:12px;width:100%">' +

                '<textarea ' +

                'id="resposta_' +
                escapeHTML(
                  duvida.id
                ) +
                '" ' +

                'placeholder="Digite a resposta para o aluno...">' +

                '</textarea>' +

                '<button class="btn btn-primary" ' +

                'onclick="responderDuvida(\'' +

                escapeHTML(
                  duvida.id
                ) +

                '\')">' +

                '📨 Responder' +

                '</button>' +

                '</div>';
            }


            card.innerHTML =

              '<h3>❓ ' +
              escapeHTML(
                duvida.assunto
              ) +
              '</h3>' +

              '<p><strong>Aluno:</strong> ' +
              escapeHTML(
                duvida.aluno
              ) +
              '</p>' +

              '<p><strong>Data:</strong> ' +
              escapeHTML(
                duvida.data
              ) +
              '</p>' +

              statusHTML +

              '<div class="duvida-mensagem">' +

              '<strong>💬 Dúvida:</strong>' +

              '<br><br>' +

              escapeHTML(
                duvida.mensagem
              ) +

              '</div>' +

              respostaHTML +

              '<div style="margin-top:12px;display:flex;gap:8px;flex-wrap:wrap">' +

              botoes +

              '</div>';


            grid.appendChild(
              card
            );
          }
        );
      }
    )

    .withFailureHandler(
      mostrarErro
    )

    .listarDuvidas(
      token,
      filtroDuvidas
    );
}


// ==========================================================
// MARCAR DÚVIDA COMO LIDA
// ==========================================================

function marcarDuvidaLida(
  id
) {

  google.script.run

    .withSuccessHandler(
      function() {

        carregarDuvidas();
      }
    )

    .withFailureHandler(
      mostrarErro
    )

    .marcarDuvidaLida(
      token,
      id
    );
}


// ==========================================================
// RESPONDER
// ==========================================================

function responderDuvida(
  id
) {

  const campo =
    el(
      "resposta_" +
      id
    );


  if (!campo) {

    alert(
      "Campo de resposta não encontrado."
    );

    return;
  }


  const resposta =
    campo.value.trim();


  if (!resposta) {

    alert(
      "Digite uma resposta."
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

        carregarDuvidas();
      }
    )

    .withFailureHandler(
      mostrarErro
    )

    .responderDuvida(
      token,
      id,
      resposta
    );
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
  )
    return;


  google.script.run

    .withSuccessHandler(
      function() {

        localStorage
          .removeItem(
            "GE_TOKEN"
          );

        location.reload();
      }
    )

    .withFailureHandler(
      function() {

        localStorage
          .removeItem(
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

function mostrarErro(
  error
) {

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


    if (!salvo)
      return;


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

          localStorage
            .removeItem(
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

// ==========================================================
// MENU MOBILE
// ==========================================================

el("menuMobile").onclick = function() {

  el("sidebar")
    .classList
    .toggle("open");

  el("mobileOverlay")
    .classList
    .toggle("show");
};


// Fechar menu clicando no fundo

el("mobileOverlay").onclick = function() {

  el("sidebar")
    .classList
    .remove("open");

  el("mobileOverlay")
    .classList
    .remove("show");
};


// Fechar menu ao escolher uma página

document
.querySelectorAll(".nav[data-screen]")
.forEach(function(button) {

  button.addEventListener("click", function() {

    if (window.innerWidth <= 850) {

      el("sidebar")
        .classList
        .remove("open");

      el("mobileOverlay")
        .classList
        .remove("show");
    }

  });

});
/* ==========================================================
OVERLAY DO MENU MOBILE
========================================================== */

#mobileOverlay {
  display: none;

  position: fixed;

  inset: 0;

  background: rgba(0,0,0,.45);

  z-index: 25;
}


#mobileOverlay.show {
  display: block;
}


@media (min-width: 851px) {

  #mobileOverlay {
    display: none !important;
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

  const elemento =
    document.getElementById("loading");

  if (!elemento) return;

  elemento.classList.toggle(
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


  if (
    tipo === "ALUNO"
  ) {

    document
      .getElementById("camposAluno")
      .classList.remove("escondido");

  }


  if (
    tipo === "PROFESSOR"
  ) {

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
        .value
        .trim(),

    turma:
      document
        .getElementById("cadTurma")
        .value
        .trim(),

    nascimento:
      document
        .getElementById("cadNascimento")
        ?.value || "",

    materia:
      document
        .getElementById("cadMateria")
        .value
        .trim()

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


      if (
        resposta.sucesso !== false
      ) {

        mostrarLogin();

      }

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


  if (
    !email ||
    !senha
  ) {

    alert(
      "Digite seu e-mail e senha."
    );

    return;

  }


  loading(true);


  google.script.run

    .withSuccessHandler(function(resposta) {

      loading(false);


      if (
        !resposta ||
        resposta.sucesso === false
      ) {

        alert(
          resposta?.mensagem ||
          "Não foi possível realizar o login."
        );

        return;

      }


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

    .login({

      usuario:
        email,

      senha:
        senha,

      tipo:
        document
          .getElementById("loginTipo")
          ?.value || "ALUNO"

    });

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


  if (
    usuario.tipo ===
    "ALUNO"
  ) {

    document
      .querySelectorAll(".professor-only")
      .forEach(function(elemento) {

        elemento.classList.add(
          "escondido"
        );

      });


    const resultados =
      document.getElementById(
        "menuResultados"
      );


    if (resultados) {

      resultados.style.display =
        "block";

    }

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


  if (
    typeof google === "undefined" ||
    !google.script
  ) {

    loading(false);

    return;

  }


  google.script.run

    .withSuccessHandler(function(resposta) {

      loading(false);

      dados =
        resposta || {};


      atualizarTudo();

    })

    .withFailureHandler(function(erro) {

      loading(false);

      alert(
        erro.message
      );

    })

    [funcao](token);

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
    document.getElementById(
      "cards"
    );


  if (!cards) return;


  cards.innerHTML = "";


  const lista = [

    [
      "📚 Matérias",
      dados.materias?.length || 0
    ],

    [
      "📝 Provas",
      dados.provas?.length || 0
    ],

    [
      "📄 Trabalhos",
      dados.trabalhos?.length || 0
    ],

    [
      "📅 Agenda",
      dados.agenda?.length || 0
    ]

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


  if (!elemento) return;


  elemento.innerHTML = "";


  (dados.materias || [])
    .forEach(function(item) {

      elemento.innerHTML +=

        "<div class='item'>" +

        "<h3>" +
        escapar(
          item.nome ??
          item.Nome
        ) +
        "</h3>" +

        "<small>" +
        "Professor: " +
        escapar(
          item.professor ??
          item.Professor
        ) +
        "</small>" +

        "<p>" +
        escapar(
          item.conteudo ??
          item.Descricao ??
          ""
        ) +
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


  if (!elemento) return;


  elemento.innerHTML = "";


  (dados.provas || [])
    .forEach(function(item) {

      const titulo =
        item.titulo ??
        item.Titulo ??
        "";


      const data =
        item.data ??
        item.Data ??
        "";


      const materia =
        item.materia ??
        item.Materia ??
        "";


      const professor =
        item.professor ??
        item.Professor ??
        "";


      const aluno =
        item.aluno ??
        item.Aluno ??
        "";


      const nota =
        formatarNota(
          item.nota ??
          item.Nota ??
          ""
        );


      const descricao =
        item.descricao ??
        item.Descricao ??
        "";


      elemento.innerHTML +=

        "<div class='item'>" +

        "<h3>" +
        escapar(titulo) +
        "</h3>" +

        "<p>📅 " +
        escapar(data) +
        "</p>" +

        "<p>📚 " +
        escapar(materia) +
        "</p>" +

        "<p>👨‍🏫 " +
        escapar(professor) +
        "</p>" +

        (
          aluno
            ? "<p>👤 " +
              escapar(aluno) +
              "</p>"
            : ""
        ) +

        (
          nota
            ? "<p>⭐ Nota: <strong>" +
              escapar(nota) +
              "</strong></p>"
            : ""
        ) +

        "<p>" +
        escapar(descricao) +
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


  if (!elemento) return;


  elemento.innerHTML = "";


  const select =
    document.getElementById(
      "entregaTrabalho"
    );


  if (select) {

    select.innerHTML =
      "<option value=''>Selecione o trabalho</option>";

  }


  (dados.trabalhos || [])
    .forEach(function(item) {

      const titulo =
        item.titulo ??
        item.Titulo ??
        "";


      const data =
        item.data ??
        item.DataEntrega ??
        item.Data ??
        "";


      const materia =
        item.materia ??
        item.Materia ??
        "";


      const descricao =
        item.descricao ??
        item.Descricao ??
        "";


      const nota =
        formatarNota(
          item.nota ??
          item.Nota ??
          ""
        );


      elemento.innerHTML +=

        "<div class='item'>" +

        "<h3>" +
        escapar(titulo) +
        "</h3>" +

        "<p>📅 Entrega: " +
        escapar(data) +
        "</p>" +

        "<p>📚 " +
        escapar(materia) +
        "</p>" +

        (
          nota
            ? "<p>⭐ Nota: <strong>" +
              escapar(nota) +
              "</strong></p>"
            : ""
        ) +

        "<p>" +
        escapar(descricao) +
        "</p>" +

        "</div>";


      if (
        usuario.tipo === "ALUNO" &&
        select
      ) {

        const id =
          item.id ??
          item.ID ??
          "";


        select.innerHTML +=

          "<option value='" +
          escapar(id) +
          "'>" +

          escapar(titulo) +

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


  if (!elemento) return;


  elemento.innerHTML = "";


  (dados.agenda || [])
    .forEach(function(item) {

      elemento.innerHTML +=

        "<div class='item'>" +

        "<h3>" +
        escapar(
          item.titulo ??
          item.Titulo ??
          ""
        ) +
        "</h3>" +

        "<p>📅 " +
        escapar(
          item.data ??
          item.Data ??
          ""
        ) +
        "</p>" +

        "<p>⏰ " +
        escapar(
          item.hora ??
          item.Hora ??
          ""
        ) +
        "</p>" +

        "<p>📌 " +
        escapar(
          item.tipo ??
          item.Tipo ??
          ""
        ) +
        "</p>" +

        "<p>" +
        escapar(
          item.descricao ??
          item.Descricao ??
          ""
        ) +
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


  if (!elemento) return;


  elemento.innerHTML = "";


  /*
   * O sistema agora utiliza as provas como fonte
   * dos resultados.
   */

  const resultados =
    dados.resultados?.length
      ? dados.resultados
      : dados.provas || [];


  resultados
    .forEach(function(item) {

      const prova =
        item.prova ??
        item.Prova ??
        item.titulo ??
        item.Titulo ??
        "";


      const aluno =
        item.aluno ??
        item.Aluno ??
        "";


      const nota =
        formatarNota(
          item.nota ??
          item.Nota ??
          ""
        );


      const observacao =
        item.observacao ??
        item.Observacao ??
        item.descricao ??
        item.Descricao ??
        "";


      const data =
        item.data ??
        item.Data ??
        "";


      const materia =
        item.materia ??
        item.Materia ??
        "";


      const professor =
        item.professor ??
        item.Professor ??
        "";


      elemento.innerHTML +=

        "<div class='item'>" +

        "<h3>" +
        escapar(prova) +
        "</h3>" +

        (
          data
            ? "<p>📅 Data: " +
              escapar(data) +
              "</p>"
            : ""
        ) +

        (
          aluno
            ? "<p>👤 Aluno: " +
              escapar(aluno) +
              "</p>"
            : ""
        ) +

        (
          materia
            ? "<p>📚 Matéria: " +
              escapar(materia) +
              "</p>"
            : ""
        ) +

        (
          professor
            ? "<p>👨‍🏫 Professor: " +
              escapar(professor) +
              "</p>"
            : ""
        ) +

        (
          nota
            ? "<p>⭐ Nota: <strong>" +
              escapar(nota) +
              "</strong></p>"
            : "<p>⭐ Nota: <strong>Não informada</strong></p>"
        ) +

        (
          observacao
            ? "<p>📝 Descrição: " +
              escapar(observacao) +
              "</p>"
            : ""
        ) +

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


      if (!select) return;


      select.innerHTML =
        "<option value=''>Selecione uma pessoa</option>";


      (contatos || [])
        .forEach(function(contato) {

          select.innerHTML +=

            "<option value='" +
            escapar(contato.ID ?? contato.id) +
            "'>" +

            escapar(
              contato.Nome ??
              contato.nome ??
              ""
            ) +

            " - " +

            escapar(
              contato.Tipo ??
              contato.tipo ??
              ""
            ) +

            "</option>";

        });

    })

    .withFailureHandler(function(erro) {

      console.error(
        erro
      );

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
      ?.value;


  if (!outroID) return;


  google.script.run

    .withSuccessHandler(function(mensagens) {

      const area =
        document.getElementById(
          "chatMensagens"
        );


      if (!area) return;


      area.innerHTML = "";


      (mensagens || [])
        .forEach(function(msg) {

          const minha =
            String(msg.RemetenteID) ===
            String(usuario.id);


          area.innerHTML +=

            "<div class='mensagem " +
            (minha ? "eu" : "") +
            "'>" +

            "<strong>" +
            escapar(
              msg.RemetenteNome
            ) +
            "</strong><br>" +

            escapar(
              msg.Mensagem
            ) +

            "</div>";

        });


      area.scrollTop =
        area.scrollHeight;

    })

    .withFailureHandler(function(erro) {

      alert(
        erro.message
      );

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
      ?.value;


  const texto =
    document
      .getElementById("chatTexto")
      ?.value
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

      alert(
        msg.mensagem ??
        msg
      );


      carregarDados();

    })

    .withFailureHandler(function(erro) {

      alert(
        erro.message
      );

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
            ?.value || "",

        descricao:
          document
            .getElementById("materiaDescricao")
            ?.value || ""

      }

    );

}


// ============================================================
// PROFESSOR - PROVA
// ============================================================

function salvarProva() {

  const nota =
    document
      .getElementById("provaValor")
      ?.value
      .trim() || "";


  const dadosProva = {

    titulo:
      document
        .getElementById("provaTitulo")
        .value
        .trim(),

    data:
      document
        .getElementById("provaData")
        .value,

    turma:
      document
        .getElementById("provaTurma")
        ?.value || "",

    materia:
      document
        .getElementById("provaMateria")
        .value
        .trim(),

    // ========================================================
    // CORREÇÃO PRINCIPAL
    // Antes estava: valor
    // Agora está: nota
    // ========================================================

    nota:
      nota,

    descricao:
      document
        .getElementById("provaDescricao")
        .value
        .trim()

  };


  google.script.run

    .withSuccessHandler(function(msg) {

      alert(
        msg.mensagem ??
        msg
      );


      carregarDados();

    })

    .withFailureHandler(function(erro) {

      alert(
        erro.message
      );

    })

    .cadastrarProva(
      token,
      dadosProva
    );

}


// ============================================================
// PROFESSOR - TRABALHO
// ============================================================

function salvarTrabalho() {

  google.script.run

    .withSuccessHandler(function(msg) {

      alert(
        msg.mensagem ??
        msg
      );


      carregarDados();

    })

    .withFailureHandler(function(erro) {

      alert(
        erro.message
      );

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
            ?.value || "",

        materia:
          document
            .getElementById("trabalhoMateria")
            .value,

        descricao:
          document
            .getElementById("trabalhoDescricao")
            .value,

        nota:
          document
            .getElementById("trabalhoNota")
            ?.value || ""

      }

    );

}


// ============================================================
// PROFESSOR - AGENDA
// ============================================================

function salvarAgenda() {

  google.script.run

    .withSuccessHandler(function(msg) {

      alert(
        msg.mensagem ??
        msg
      );


      carregarDados();

    })

    .withFailureHandler(function(erro) {

      alert(
        erro.message
      );

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

      alert(
        msg.mensagem ??
        msg
      );


      carregarDados();

    })

    .withFailureHandler(function(erro) {

      alert(
        erro.message
      );

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

      alert(
        msg.mensagem ??
        msg
      );


      document
        .getElementById("entregaMensagem")
        .value = "";

    })

    .withFailureHandler(function(erro) {

      alert(
        erro.message
      );

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
// FORMATAÇÃO DE NOTA
// ============================================================

function formatarNota(valor) {

  if (
    valor === null ||
    valor === undefined ||
    valor === ""
  ) {

    return "";

  }


  // ----------------------------------------------------------
  // NÃO PERMITIR DATA COMO NOTA
  // ----------------------------------------------------------

  if (
    valor instanceof Date
  ) {

    return "";

  }


  const texto =
    String(valor)
      .trim();


  if (
    !texto
  ) {

    return "";

  }


  // Se vier como texto de Date
  if (
    texto.includes("GMT") &&
    !isNaN(Date.parse(texto))
  ) {

    return "";

  }


  const numero =
    Number(
      texto.replace(",", ".")
    );


  if (
    isNaN(numero)
  ) {

    return texto;

  }


  return numero
    .toFixed(1)
    .replace(".", ",");

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

    .replace(
      /&/g,
      "&amp;"
    )

    .replace(
      /</g,
      "&lt;"
    )

    .replace(
      />/g,
      "&gt;"
    )

    .replace(
      /"/g,
      "&quot;"
    )

    .replace(
      /'/g,
      "&#039;"
    );

}


// ============================================================
// RECUPERAÇÃO DE SENHA
// ============================================================

let recuperacaoTipo =
  "ALUNO";


// ============================================================
// ABRIR RECUPERAÇÃO
// ============================================================

function abrirRecuperacaoSenha() {

  const login =
    document.getElementById(
      "loginTela"
    );


  const cadastro =
    document.getElementById(
      "cadastroTela"
    );


  const recuperacao =
    document.getElementById(
      "forgotPasswordPage"
    );


  if (!recuperacao) {

    alert(
      "A tela de recuperação de senha não existe no HTML."
    );

    return;

  }


  if (login) {

    login.classList.add(
      "escondido"
    );

  }


  if (cadastro) {

    cadastro.classList.add(
      "escondido"
    );

  }


  recuperacao.classList.remove(
    "escondido"
  );


  const ra =
    document.getElementById(
      "forgotRA"
    );


  const usuario =
    document.getElementById(
      "forgotUsuario"
    );


  const novaSenha =
    document.getElementById(
      "forgotNovaSenha"
    );


  const confirmarSenha =
    document.getElementById(
      "forgotConfirmarSenha"
    );


  const mensagem =
    document.getElementById(
      "forgotMsg"
    );


  if (ra)
    ra.value = "";


  if (usuario)
    usuario.value = "";


  if (novaSenha)
    novaSenha.value = "";


  if (confirmarSenha)
    confirmarSenha.value = "";


  if (mensagem) {

    mensagem.textContent = "";

    mensagem.className =
      "message";

  }


  recuperarAluno();

}


// ============================================================
// ALUNO
// ============================================================

function recuperarAluno() {

  recuperacaoTipo =
    "ALUNO";


  const aluno =
    document.getElementById(
      "forgotAluno"
    );


  const professor =
    document.getElementById(
      "forgotProfessor"
    );


  if (aluno) {

    aluno.classList.add(
      "active"
    );

  }


  if (professor) {

    professor.classList.remove(
      "active"
    );

  }

}


// ============================================================
// PROFESSOR
// ============================================================

function recuperarProfessor() {

  recuperacaoTipo =
    "PROFESSOR";


  const aluno =
    document.getElementById(
      "forgotAluno"
    );


  const professor =
    document.getElementById(
      "forgotProfessor"
    );


  if (professor) {

    professor.classList.add(
      "active"
    );

  }


  if (aluno) {

    aluno.classList.remove(
      "active"
    );

  }

}


// ============================================================
// VOLTAR LOGIN
// ============================================================

function voltarLoginSenha() {

  const recuperacao =
    document.getElementById(
      "forgotPasswordPage"
    );


  const login =
    document.getElementById(
      "loginTela"
    );


  if (recuperacao) {

    recuperacao.classList.add(
      "escondido"
    );

  }


  if (login) {

    login.classList.remove(
      "escondido"
    );

  }


  const mensagem =
    document.getElementById(
      "forgotMsg"
    );


  if (mensagem) {

    mensagem.textContent = "";

    mensagem.className =
      "message";

  }

}


// ============================================================
// MENSAGEM
// ============================================================

function mensagemRecuperacao(
  texto,
  sucesso
) {

  const elemento =
    document.getElementById(
      "forgotMsg"
    );


  if (!elemento) {

    alert(texto);

    return;

  }


  elemento.textContent =
    texto;


  elemento.className =
    sucesso
      ? "message sucesso"
      : "message erro";

}


// ============================================================
// ALTERAR SENHA
// ============================================================

function executarAlteracaoSenha() {

  const raElemento =
    document.getElementById(
      "forgotRA"
    );


  const usuarioElemento =
    document.getElementById(
      "forgotUsuario"
    );


  const novaSenhaElemento =
    document.getElementById(
      "forgotNovaSenha"
    );


  const confirmarSenhaElemento =
    document.getElementById(
      "forgotConfirmarSenha"
    );


  if (
    !raElemento ||
    !usuarioElemento ||
    !novaSenhaElemento ||
    !confirmarSenhaElemento
  ) {

    mensagemRecuperacao(
      "Os campos da recuperação de senha não foram encontrados.",
      false
    );

    return;

  }


  const ra =
    raElemento.value.trim();


  const usuario =
    usuarioElemento.value.trim();


  const novaSenha =
    novaSenhaElemento.value;


  const confirmarSenha =
    confirmarSenhaElemento.value;


  if (!ra) {

    mensagemRecuperacao(
      "Informe o RA/ID.",
      false
    );

    raElemento.focus();

    return;

  }


  if (!usuario) {

    mensagemRecuperacao(
      "Informe o usuário.",
      false
    );

    usuarioElemento.focus();

    return;

  }


  if (!novaSenha) {

    mensagemRecuperacao(
      "Informe a nova senha.",
      false
    );

    novaSenhaElemento.focus();

    return;

  }


  if (
    novaSenha.length < 6
  ) {

    mensagemRecuperacao(
      "A nova senha precisa ter pelo menos 6 caracteres.",
      false
    );

    novaSenhaElemento.focus();

    return;

  }


  if (!confirmarSenha) {

    mensagemRecuperacao(
      "Confirme a nova senha.",
      false
    );

    confirmarSenhaElemento.focus();

    return;

  }


  if (
    novaSenha !==
    confirmarSenha
  ) {

    mensagemRecuperacao(
      "As senhas não são iguais.",
      false
    );

    confirmarSenhaElemento.focus();

    return;

  }


  const botao =
    document.getElementById(
      "btnRecuperarSenha"
    );


  if (botao) {

    botao.disabled = true;

    botao.dataset.textoOriginal =
      botao.textContent;

    botao.textContent =
      "Alterando senha...";

  }


  mensagemRecuperacao(
    "Verificando sua conta...",
    true
  );


  google.script.run

    .withSuccessHandler(function(resultado) {

      if (botao) {

        botao.disabled = false;

        botao.textContent =
          botao.dataset.textoOriginal ||
          "Alterar senha";

      }


      if (
        !resultado ||
        resultado.sucesso !== true
      ) {

        mensagemRecuperacao(

          resultado &&
          resultado.mensagem
            ? resultado.mensagem
            : "Não foi possível alterar a senha.",

          false

        );

        return;

      }


      mensagemRecuperacao(

        resultado.mensagem ||
        "Senha alterada com sucesso!",

        true

      );


      novaSenhaElemento.value = "";

      confirmarSenhaElemento.value = "";


      setTimeout(function() {

        voltarLoginSenha();


        const loginUsuario =
          document.getElementById(
            "loginEmail"
          );


        if (loginUsuario) {

          loginUsuario.value =
            usuario;

        }


        const loginSenha =
          document.getElementById(
            "loginSenha"
          );


        if (loginSenha) {

          loginSenha.value = "";

          loginSenha.focus();

        }

      }, 1500);

    })

    .withFailureHandler(function(erro) {

      if (botao) {

        botao.disabled = false;

        botao.textContent =
          botao.dataset.textoOriginal ||
          "Alterar senha";

      }


      mensagemRecuperacao(

        erro &&
        erro.message
          ? erro.message
          : "Erro ao alterar a senha.",

        false

      );

    })

    .recuperarSenha({

      ra:
        ra,

      usuario:
        usuario,

      novaSenha:
        novaSenha,

      tipo:
        recuperacaoTipo

    });

}

</script>
