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
  DUVIDA: "DUVIDA"
};


// ============================================================
// CONFIGURAÇÕES DE SEGURANÇA
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

  const ss = getPlanilha();

  let aba =
    ss.getSheetByName(nome);

  if (!aba) {

    aba =
      ss.insertSheet(nome);
  }

  return aba;
}


// ============================================================
// CONFIGURAR BANCO DE DADOS
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


  // ----------------------------------------------------------
  // MIGRAR SENHAS ANTIGAS
  // ----------------------------------------------------------

  migrarSenhasAntigas();


  return "Planilha configurada com sucesso.";
}


// ============================================================
// CRIAR ABA
// ============================================================

function criarAba(nome, cabecalhos) {

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
// MIGRAÇÃO AUTOMÁTICA DAS SENHAS
// ============================================================
//
// Caso a planilha antiga possua:
//
// Senha
//
// ela será convertida para:
//
// SenhaHash
// Salt
//
// A senha original será substituída pelo hash.
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


  // ----------------------------------------------------------
  // JÁ ESTÁ CONFIGURADO
  // ----------------------------------------------------------

  if (
    indiceHash !== -1 &&
    indiceSalt !== -1
  ) {

    return;
  }


  // ----------------------------------------------------------
  // PLANILHA ANTIGA
  // ----------------------------------------------------------

  if (
    indiceSenha !== -1
  ) {

    const colunaSenhaReal =
      indiceSenha + 1;


    // Renomeia Senha para SenhaHash
    aba
      .getRange(
        1,
        colunaSenhaReal
      )
      .setValue("SenhaHash");


    // Insere Salt logo depois
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
// GERAR SALT
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
// GERAR HASH DA SENHA
// ============================================================
//
// Salt + senha
//
// Depois são realizadas várias rodadas de SHA-256.
// ============================================================

function gerarHashSenha(
  senha,
  salt
) {

  senha =
    String(senha || "");


  salt =
    String(salt || "");


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
// CADASTRO DE USUÁRIO
// ============================================================

function cadastrarUsuario(dados) {

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


  // ----------------------------------------------------------
  // VALIDAÇÕES
  // ----------------------------------------------------------

  if (!ra) {

    return {
      sucesso: false,
      mensagem:
        "Informe o RA/ID."
    };
  }


  if (!nome) {

    return {
      sucesso: false,
      mensagem:
        "Informe o nome."
    };
  }


  if (!usuario) {

    return {
      sucesso: false,
      mensagem:
        "Informe o usuário."
    };
  }


  if (!senha) {

    return {
      sucesso: false,
      mensagem:
        "Informe a senha."
    };
  }


  if (senha.length < 6) {

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


  // ----------------------------------------------------------
  // ALUNO
  // ----------------------------------------------------------

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


      // RA
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


      // Usuário
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


    // --------------------------------------------------------
    // GERAR SALT E HASH
    // --------------------------------------------------------

    const salt =
      gerarSalt();


    const senhaHash =
      gerarHashSenha(
        senha,
        salt
      );


    // --------------------------------------------------------
    // SALVAR
    // --------------------------------------------------------

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


  // ----------------------------------------------------------
  // PROFESSOR
  // ----------------------------------------------------------

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


      // RA / ID
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


      // Usuário
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

        const senhaValida =
          verificarSenha(
            senha,
            senhaHash,
            salt
          );


        if (!senhaValida) {

          return {
            sucesso: false,
            mensagem:
              "Usuário ou senha incorretos."
          };
        }


        const usuario = {

          tipo:
            "ALUNO",

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

        const senhaValida =
          verificarSenha(
            senha,
            senhaHash,
            salt
          );


        if (!senhaValida) {

          return {
            sucesso: false,
            mensagem:
              "Usuário ou senha incorretos."
          };
        }


        const usuario = {

          tipo:
            "PROFESSOR",

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
    .map(
      function(
        linha,
        index
      ) {

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
      }
    )
    .filter(
      function(materia) {

        return (
          normalizar(
            materia.serieAno
          ) ===
          normalizar(
            usuario.serieAno
          )
        );
      }
    );
}


function atualizarMateria(
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
      "Acesso negado."
    );
  }


  const aba =
    getAba(
      ABAS.MATERIA
    );


  const linha =
    Number(
      dados.linha
    );


  const existente =
    aba
      .getRange(
        linha,
        1,
        1,
        8
      )
      .getValues()[0];


  if (
    String(existente[4]) !==
    String(professor.ra)
  ) {

    throw new Error(
      "Você não pode editar esta matéria."
    );
  }


  aba
    .getRange(
      linha,
      2,
      1,
      3
    )
    .setValues([[
      dados.nome,
      dados.conteudo,
      dados.links
    ]]);


  return {

    sucesso: true,

    mensagem:
      "Matéria atualizada."
  };
}


function excluirMateria(
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
      ABAS.MATERIA
    );


  const dados =
    aba
      .getRange(
        Number(linha),
        1,
        1,
        8
      )
      .getValues()[0];


  if (
    String(dados[4]) !==
    String(professor.ra)
  ) {

    throw new Error(
      "Você não pode excluir esta matéria."
    );
  }


  aba.deleteRow(
    Number(linha)
  );


  return {

    sucesso: true,

    mensagem:
      "Matéria excluída."
  };
}


// ============================================================
// ALUNOS DA SÉRIE DO PROFESSOR
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

    .map(
      function(linha) {

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
      }
    )

    .filter(
      function(aluno) {

        return (
          normalizar(
            aluno.serieAno
          ) ===
          normalizar(
            professor.serieAno
          )
        );
      }
    );
}


// ============================================================
// PROVAS
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


  // ----------------------------------------------------------
  // TODOS DA SÉRIE
  // ----------------------------------------------------------

  if (
    dados.todosSerie === true ||
    String(
      dados.todosSerie
    ) === "true"
  ) {

    alunosDestino =
      alunos.filter(
        function(linha) {

          return (
            normalizar(
              linha[5]
            ) ===
            normalizar(
              professor.serieAno
            )
          );
        }
      );

  } else {

    const aluno =
      alunos.find(
        function(linha) {

          return (
            String(linha[0]) ===
            String(
              dados.alunoRA
            )
          );
        }
      );


    if (!aluno) {

      throw new Error(
        "Aluno não encontrado pelo RA."
      );
    }


    if (
      normalizar(
        aluno[5]
      ) !==
      normalizar(
        professor.serieAno
      )
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
      ABAS.PROVA
    );


  const linhas = [];


  alunosDestino.forEach(
    function(aluno) {

      linhas.push([

        gerarID("PROVA"),

        dados.data,

        dados.titulo,

        String(
          aluno[0]
        ),

        String(
          aluno[1]
        ),

        String(
          aluno[5]
        ),

        dados.materia,

        dados.nota || "",

        dados.descricao || "",

        professor.ra,

        professor.nome,

        new Date()

      ]);
    }
  );


  aba
    .getRange(
      aba.getLastRow() + 1,
      1,
      linhas.length,
      linhas[0].length
    )
    .setValues(linhas);


  return {

    sucesso: true,

    quantidade:
      alunosDestino.length,

    mensagem:
      dados.todosSerie
        ? "Prova publicada para " +
          alunosDestino.length +
          " aluno(s) da Série/Ano."
        : "Prova cadastrada para o aluno."
  };
}


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

    .map(
      function(
        linha,
        index
      ) {

        return {

          linha:
            index + 2,

          id:
            String(linha[0]),

          data:
            formatarData(
              linha[1]
            ),

          titulo:
            String(linha[2]),

          alunoRA:
            String(linha[3]),

          aluno:
            String(linha[4]),

          serieAno:
            String(linha[5]),

          materia:
            String(linha[6]),

          nota:
            linha[7] === ""
              ? ""
              : String(
                  linha[7]
                ),

          descricao:
            String(
              linha[8] || ""
            ),

          professorRA:
            String(linha[9]),

          professor:
            String(linha[10])
        };
      }
    )

    .filter(
      function(prova) {

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
      }
    );
}


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
      "Você não pode excluir esta prova."
    );
  }


  aba.deleteRow(
    Number(linha)
  );


  return {

    sucesso: true,

    mensagem:
      "Prova excluída."
  };
}


// ============================================================
// TRABALHOS
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
      alunos.filter(
        function(linha) {

          return (
            normalizar(
              linha[5]
            ) ===
            normalizar(
              professor.serieAno
            )
          );
        }
      );

  } else {

    const aluno =
      alunos.find(
        function(linha) {

          return (
            String(
              linha[0]
            ) ===
            String(
              dados.alunoRA
            )
          );
        }
      );


    if (!aluno) {

      throw new Error(
        "Aluno não encontrado pelo RA."
      );
    }


    if (
      normalizar(
        aluno[5]
      ) !==
      normalizar(
        professor.serieAno
      )
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


  alunosDestino.forEach(
    function(aluno) {

      linhas.push([

        gerarID("TRAB"),

        dados.data,

        dados.titulo,

        String(
          aluno[0]
        ),

        String(
          aluno[1]
        ),

        String(
          aluno[5]
        ),

        dados.materia,

        dados.nota || "",

        dados.descricao || "",

        professor.ra,

        professor.nome,

        new Date()

      ]);
    }
  );


  aba
    .getRange(
      aba.getLastRow() + 1,
      1,
      linhas.length,
      linhas[0].length
    )
    .setValues(linhas);


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

    .map(
      function(
        linha,
        index
      ) {

        return {

          linha:
            index + 2,

          id:
            String(linha[0]),

          data:
            formatarData(
              linha[1]
            ),

          titulo:
            String(linha[2]),

          alunoRA:
            String(linha[3]),

          aluno:
            String(linha[4]),

          serieAno:
            String(linha[5]),

          materia:
            String(linha[6]),

          nota:
            linha[7] === ""
              ? ""
              : String(
                  linha[7]
                ),

          descricao:
            String(
              linha[8] || ""
            ),

          professorRA:
            String(linha[9]),

          professor:
            String(linha[10])
        };
      }
    )

    .filter(
      function(trabalho) {

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
      }
    );
}


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
      professores.find(
        function(linha) {

          return (
            String(
              linha[1]
            ) ===
            professorRA
          );
        }
      );
  }


  // ----------------------------------------------------------
  // PROFESSOR DA MESMA SÉRIE
  // ----------------------------------------------------------

  if (!professor) {

    professor =
      professores.find(
        function(linha) {

          return (
            normalizar(
              linha[6]
            ) ===
            normalizar(
              aluno.serieAno
            )
          );
        }
      );
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

      String(
        professor[1]
      ),

      String(
        professor[2]
      ),

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
// LISTAR PROFESSORES
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

      .map(
        function(linha) {

          return {

            ra:
              String(
                linha[1]
              ),

            nome:
              String(
                linha[2]
              ),

            materia:
              String(
                linha[7]
              ),

            serieAno:
              String(
                linha[6]
              )
          };
        }
      )

      .filter(
        function(professor) {

          return (
            normalizar(
              professor.serieAno
            ) ===
            normalizar(
              usuario.serieAno
            )
          );
        }
      );
  }


  return [];
}


// ============================================================
// LISTAR DÚVIDAS
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
    dados.map(
      function(
        linha,
        index
      ) {

        return {

          linha:
            index + 2,

          id:
            String(
              linha[0]
            ),

          data:
            formatarData(
              linha[1]
            ),

          alunoRA:
            String(
              linha[2]
            ),

          aluno:
            String(
              linha[3]
            ),

          professorRA:
            String(
              linha[4]
            ),

          professor:
            String(
              linha[5]
            ),

          assunto:
            String(
              linha[6]
            ),

          mensagem:
            String(
              linha[7]
            ),

          resposta:
            String(
              linha[8] || ""
            ),

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
      }
    );


  // ----------------------------------------------------------
  // SEGURANÇA
  // ----------------------------------------------------------

  if (
    usuario.tipo ===
    "ALUNO"
  ) {

    lista =
      lista.filter(
        function(item) {

          return (
            String(
              item.alunoRA
            ) ===
            String(
              usuario.ra
            )
          );
        }
      );

  } else {

    lista =
      lista.filter(
        function(item) {

          return (
            String(
              item.professorRA
            ) ===
            String(
              usuario.ra
            )
          );
        }
      );
  }


  // ----------------------------------------------------------
  // FILTRO
  // ----------------------------------------------------------

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
      lista.filter(
        function(item) {

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
        }
      );
  }


  if (
    filtro ===
    "LIDAS"
  ) {

    lista =
      lista.filter(
        function(item) {

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
        }
      );
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
      String(
        dados[i][0]
      ) ===
      String(id)
    ) {

      const linha =
        i + 2;


      if (
        usuario.tipo ===
        "PROFESSOR"
      ) {

        if (
          String(
            dados[i][4]
          ) !==
          String(
            usuario.ra
          )
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
          String(
            dados[i][2]
          ) !==
          String(
            usuario.ra
          )
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
      String(
        dados[i][0]
      ) ===
      String(id)
    ) {

      if (
        String(
          dados[i][4]
        ) !==
        String(
          professor.ra
        )
      ) {

        throw new Error(
          "Você não pode responder esta dúvida."
        );
      }


      const linha =
        i + 2;


      // ------------------------------------------------------
      // RESPOSTA
      // ------------------------------------------------------

      aba
        .getRange(
          linha,
          9
        )
        .setValue(
          resposta
        );


      // ------------------------------------------------------
      // ALUNO RECEBE COMO NÃO LIDA
      // ------------------------------------------------------

      aba
        .getRange(
          linha,
          10
        )
        .setValue(
          "NAO_LIDA"
        );


      // ------------------------------------------------------
      // PROFESSOR JÁ LEU
      // ------------------------------------------------------

      aba
        .getRange(
          linha,
          11
        )
        .setValue(
          "LIDA"
        );


      // ------------------------------------------------------
      // DATA
      // ------------------------------------------------------

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