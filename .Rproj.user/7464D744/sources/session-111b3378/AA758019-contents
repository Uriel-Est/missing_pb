
# ================================================================
# AUDITORIA DE QUALIDADE - PESSOAS DESAPARECIDAS
# POLICIA CIVIL - 2020 A 2026
# ================================================================
#
# O script NAO modifica o arquivo original.
#
# Saidas:
#   1. Diagnostico no console
#   2. Um XLSX consolidado (opcional)
#   3. README.md (opcional)
#
# ================================================================


# ================================================================
# 0. CONFIGURACOES
# ================================================================

ARQUIVO <- paste0(
  "C:/Users/uriel/Documents/UFPB Estatística/",
  "Freelance - Desaparecidos/",
  "Polícia Civil - Pessoas Desaparecidas ",
  "(2020-2026)_24Set2026 evelyne.xlsx"
)

PASTA_SAIDA <- dirname(ARQUIVO)

ANOS <- 2020:2026

GERAR_XLSX <- TRUE
GERAR_README <- TRUE

# Impressao no console
TOP_CATEGORIAS <- 15
MOSTRAR_QUALIDADE_COMPLETA <- TRUE

# Limites de verificacao
IDADE_MIN <- 0
IDADE_MAX <- 110

# Diferenca tolerada na comparacao entre idade
# declarada e idade calculada a partir do nascimento.
TOLERANCIA_IDADE <- 1

# Data de referencia do arquivo recebido.
# Utilizada para sinalizar datas posteriores a extracao.
DATA_CORTE <- as.Date("2026-09-24")

options(
  width = 180,
  max.print = 10000,
  scipen = 999
)


# ================================================================
# 1. PACOTES
# ================================================================

pacotes <- c(
  "readxl",
  "dplyr",
  "tidyr",
  "purrr",
  "stringr",
  "stringi",
  "tibble",
  "openxlsx"
)

faltantes <- pacotes[
  !vapply(
    pacotes,
    requireNamespace,
    logical(1),
    quietly = TRUE
  )
]

if (length(faltantes) > 0) {
  install.packages(faltantes)
}

suppressPackageStartupMessages({
  
  library(readxl)
  library(dplyr)
  library(tidyr)
  library(purrr)
  library(stringr)
  library(stringi)
  library(tibble)
  library(openxlsx)
  
})


# ================================================================
# 2. FUNCOES AUXILIARES
# ================================================================

titulo <- function(x) {
  
  cat("\n")
  cat(strrep("=", 75), "\n")
  cat(x, "\n")
  cat(strrep("=", 75), "\n")
  
}


normalizar <- function(x) {
  
  x <- as.character(x)
  
  x <- stringi::stri_trans_general(
    x,
    "Latin-ASCII"
  )
  
  x <- toupper(trimws(x))
  
  x <- gsub("[\r\n\t]+", " ", x)
  
  x <- gsub("[^A-Z0-9]+", " ", x)
  
  x <- gsub("\\s+", " ", x)
  
  trimws(x)
  
}


# Marcadores de ausencia informacional.
# Mantemos separados dos NA reais.

MARCADORES_NI <- c(
  "NI",
  "N I",
  "NAO INFORMADO",
  "NAO INFORMADA",
  "NAO CONSTA",
  "SEM INFORMACAO",
  "INDETERMINADO",
  "PREJUDICADO",
  "-"
)


classificar_ausencia <- function(x) {
  
  bruto <- trimws(as.character(x))
  
  vazio <- is.na(bruto) | bruto == ""
  
  # Identifica símbolos antes de normalizar.
  simbolos_ausencia <- bruto %in% c(
    "-",
    "--",
    "---"
  )
  
  norm <- normalizar(bruto)
  
  ni <- (
    !vazio &
      (
        simbolos_ausencia |
          norm %in% MARCADORES_NI
      )
  )
  
  resultado <- rep(
    "PREENCHIDO",
    length(bruto)
  )
  
  resultado[ni] <- "NAO_INFORMADO"
  resultado[vazio] <- "VAZIO"
  
  resultado
}

# ------------------------------------------------
# Mapeamento provisório dos nomes das variaveis
# ------------------------------------------------

identificar_variavel <- function(nome) {
  
  n <- normalizar(nome)
  
  if (grepl(
    "^DATA DA OCORRENCIA.*DATA DO DESAPARECIMENTO",
    n
  )) {
    return("data_desaparecimento")
  }
  
  if (n == "DATA DO REGISTRO DA OCORRENCIA") {
    return("data_registro")
  }
  
  if (n == "MUNICIPIO DO DESAPARECIMENTO") {
    return("municipio_desaparecimento")
  }
  
  if (n == "BAIRRO DO DESAPARECIMENTO") {
    return("bairro_desaparecimento")
  }
  
  if (n %in% c(
    "IDADE",
    "IDADE DO DESAPARECIDO"
  )) {
    return("idade")
  }
  
  if (n %in% c(
    "SEXO",
    "SEXO DO DESAPARECIDO"
  )) {
    return("sexo")
  }
  
  if (n %in% c(
    "COR RACA",
    "COR RACA DO DESAPARECIDO"
  )) {
    return("cor_raca")
  }
  
  if (n == "DATA DE NASCIMENTO DO DESAPARECIDO") {
    return("data_nascimento")
  }
  
  if (n == "RELACAO COM A VITIMA") {
    return("relacao_com_vitima")
  }
  
  if (n %in% c(
    "CONDICAO DE LOCALIZACAO",
    "DESFECHO DA OCORRENCIA",
    paste(
      "STATUS DA OCORRENCIA",
      "SITUACAO DA VITIMA DE DESAPARECIMENTO"
    )
  )) {
    return("status_desfecho")
  }
  
  if (n == "PESSOA LOCALIZADA") {
    return("pessoa_localizada")
  }
  
  if (n %in% c(
    "DATA FINAL",
    "DATA DO DESFECHO FINAL DA OCORRENCIA",
    "DATA DA LOCALIZACAO"
  )) {
    return("data_localizacao")
  }
  
  if (grepl(
    "DIAS DO REGISTRO ATE LOCALIZACAO",
    n
  )) {
    return("dias_ate_localizacao")
  }
  
  if (grepl(
    "^MOTIVACAO DO DESAPARECIMENTO",
    n
  )) {
    return("motivacao_desaparecimento")
  }
  
  if (grepl(
    "^TIPO DE LOCAL DO ENCONTRO",
    n
  )) {
    return("tipo_local_encontro")
  }
  
  if (grepl(
    "^MUNICIPIO DA LOCALIZACAO",
    n
  )) {
    return("municipio_localizacao")
  }
  
  if (grepl(
    "^BAIRRO DA LOCALIZACAO",
    n
  )) {
    return("bairro_localizacao")
  }
  
  # Colunas nao reconhecidas
  paste0(
    "nao_mapeada_",
    tolower(gsub(" ", "_", n))
  )
  
}


# ------------------------------------------------
# Conversao exploratoria de datas
# ------------------------------------------------

converter_data <- function(x) {
  
  x <- trimws(as.character(x))
  
  resultado <- rep(
    as.Date(NA),
    length(x)
  )
  
  # Data serial do Excel
  serial <- grepl(
    "^[0-9]{5}(\\.0+)?$",
    x
  )
  
  serial[is.na(serial)] <- FALSE
  
  if (any(serial)) {
    
    indices <- which(serial)
    
    numeros <- suppressWarnings(
      as.numeric(x[indices])
    )
    
    validos <- (
      !is.na(numeros) &
        numeros >= 20000 &
        numeros <= 80000
    )
    
    resultado[indices[validos]] <- as.Date(
      numeros[validos],
      origin = "1899-12-30"
    )
  }
  
  # ISO: YYYY-MM-DD ou YYYY-MM-DD HH:MM:SS
  iso <- grepl(
    "^[0-9]{4}-[0-9]{2}-[0-9]{2}",
    x
  )
  
  iso[is.na(iso)] <- FALSE
  
  if (any(iso)) {
    
    resultado[iso] <- suppressWarnings(
      as.Date(
        substr(x[iso], 1, 10),
        format = "%Y-%m-%d"
      )
    )
  }
  
  # Brasileiro: DD/MM/YYYY ou DD.MM.YYYY
  brasileiro <- grepl(
    "^[0-9]{1,2}[/.][0-9]{1,2}[/.][0-9]{4}$",
    x
  )
  
  brasileiro[is.na(brasileiro)] <- FALSE
  
  if (any(brasileiro)) {
    
    resultado[brasileiro] <- suppressWarnings(
      as.Date(
        gsub("\\.", "/", x[brasileiro]),
        format = "%d/%m/%Y"
      )
    )
  }
  
  resultado
  
}


converter_numero <- function(x) {
  
  x <- trimws(as.character(x))
  
  x <- gsub(",", ".", x, fixed = TRUE)
  
  suppressWarnings(as.numeric(x))
  
}


# ------------------------------------------------
# Classificar formato de data
# ------------------------------------------------

formato_data <- function(x) {
  
  x <- trimws(as.character(x))
  
  resultado <- rep("OUTRO", length(x))
  
  # ------------------------------------------------
  # 1. Identificar valores ausentes
  # ------------------------------------------------
  
  situacao <- classificar_ausencia(x)
  
  resultado[situacao == "VAZIO"] <- "VAZIO"
  
  resultado[situacao == "NAO_INFORMADO"] <- "NAO_INFORMADO"
  
  # ------------------------------------------------
  # 2. Identificar formatos de datas
  # ------------------------------------------------
  
  # Serial do Excel: 44837, 45294, 46021 etc.
  serial <- grepl(
    "^[0-9]{5}(\\.0+)?$",
    x
  )
  
  # Data ISO: 2022-12-15 ou 2022-12-15 00:00:00
  iso <- grepl(
    "^[0-9]{4}-[0-9]{2}-[0-9]{2}($|[ T])",
    x
  )
  
  # Data brasileira: 15/12/2022 ou 15.12.2022
  brasileiro <- grepl(
    "^[0-9]{1,2}[/.][0-9]{1,2}[/.][0-9]{4}$",
    x
  )
  
  serial[is.na(serial)] <- FALSE
  iso[is.na(iso)] <- FALSE
  brasileiro[is.na(brasileiro)] <- FALSE
  
  # Somente valores efetivamente preenchidos
  preenchido <- situacao == "PREENCHIDO"
  
  resultado[serial & preenchido] <- "SERIAL_EXCEL"
  
  resultado[iso & preenchido] <- "DATA_ISO"
  
  resultado[brasileiro & preenchido] <- "DATA_BR"
  
  # ------------------------------------------------
  # 3. Separar formatos reconhecidos de inválidos
  # ------------------------------------------------
  
  datas <- converter_data(x)
  
  invalida <- (
    preenchido &
      is.na(datas)
  )
  
  resultado[invalida] <- "NAO_INTERPRETAVEL"
  
  # ------------------------------------------------
  # 4. Retornar classificação
  # ------------------------------------------------
  
  return(resultado)
}


# ------------------------------------------------
# Impressao de tabelas
# ------------------------------------------------

mostrar <- function(df, nome, limite = Inf) {
  
  titulo(nome)
  
  if (nrow(df) == 0) {
    cat("Nenhuma observacao.\n")
    return(invisible(NULL))
  }
  
  n_mostrar <- if (is.infinite(limite)) {
    nrow(df)
  } else {
    min(nrow(df), limite)
  }
  
  print(
    as.data.frame(
      head(df, n_mostrar)
    ),
    row.names = FALSE,
    max = n_mostrar
  )
  
  invisible(NULL)
  
}


# ================================================================
# 3. IDENTIFICACAO DAS ABAS
# ================================================================

if (!file.exists(ARQUIVO)) {
  stop("Arquivo nao encontrado: ", ARQUIVO)
}

abas <- excel_sheets(ARQUIVO)

mapa_abas <- tibble(
  aba = abas,
  ano = as.integer(
    str_extract(abas, "202[0-6]")
  )
) %>%
  filter(ano %in% ANOS) %>%
  arrange(ano, aba)

if (nrow(mapa_abas) == 0) {
  stop("Nenhuma aba de 2020 a 2026 encontrada.")
}

mostrar(
  mapa_abas,
  "ABAS IDENTIFICADAS"
)

anos_ausentes <- setdiff(
  ANOS,
  unique(mapa_abas$ano)
)

if (length(anos_ausentes) > 0) {
  warning(
    "Anos sem aba: ",
    paste(anos_ausentes, collapse = ", ")
  )
}


# ================================================================
# 4. OBJETOS DE RESULTADO
# ================================================================

lista_resumo <- list()
lista_estrutura <- list()
lista_qualidade <- list()
lista_categorias <- list()
lista_problemas <- list()
lista_formatos <- list()

# Lista de registros com problemas especificos.
lista_casos <- list()


# ================================================================
# 5. PROCESSAMENTO DAS ABAS
# ================================================================

for (i in seq_len(nrow(mapa_abas))) {
  
  ano <- mapa_abas$ano[i]
  aba <- mapa_abas$aba[i]
  
  titulo(paste("ANALISANDO", aba))
  
  # Duas leituras:
  # - Texto: preservar valores e diagnosticar formatos
  # - Automatica: verificar tipos inferidos pelo readxl
  
  df <- read_excel(
    ARQUIVO,
    sheet = aba,
    col_types = "text",
    .name_repair = "minimal"
  )
  
  df_auto <- read_excel(
    ARQUIVO,
    sheet = aba,
    .name_repair = "minimal"
  )
  
  nomes_originais <- names(df)
  
  # Nomes internos por posicao
  names(df) <- paste0(
    "col_", seq_along(df)
  )
  
  n <- nrow(df)
  p <- ncol(df)
  
  variaveis <- vapply(
    nomes_originais,
    identificar_variavel,
    character(1)
  )
  
  # ------------------------------------------------------------
  # 5.1 ESTRUTURA
  # ------------------------------------------------------------
  
  estrutura_ano <- tibble(
    ano = ano,
    aba = aba,
    posicao = seq_len(p),
    nome_original = nomes_originais,
    nome_normalizado = normalizar(nomes_originais),
    variavel = variaveis,
    tipo_readxl = vapply(
      df_auto,
      function(x) paste(class(x), collapse = "/"),
      character(1)
    )
  )
  
  lista_estrutura[[aba]] <- estrutura_ano
  
  # ------------------------------------------------------------
  # 5.2 DUPLICADOS E RESUMO
  # ------------------------------------------------------------
  
  duplicados <- sum(duplicated(df))
  
  classificacoes <- lapply(
    df,
    classificar_ausencia
  )
  
  total_vazios <- sum(vapply(
    classificacoes,
    function(x) sum(x == "VAZIO"),
    numeric(1)
  ))
  
  total_ni <- sum(vapply(
    classificacoes,
    function(x) sum(x == "NAO_INFORMADO"),
    numeric(1)
  ))
  
  total_celulas <- n * p
  
  lista_resumo[[aba]] <- tibble(
    ano = ano,
    aba = aba,
    registros = n,
    colunas = p,
    pct_vazios = round(
      100 * total_vazios / max(total_celulas, 1),
      2
    ),
    pct_nao_informados = round(
      100 * total_ni / max(total_celulas, 1),
      2
    ),
    duplicados_integrais = duplicados
  )
  
  cat("Registros:", n, "\n")
  cat("Colunas:", p, "\n")
  cat("Duplicados integrais:", duplicados, "\n")
  
  # ------------------------------------------------------------
  # 5.3 QUALIDADE POR COLUNA
  # ------------------------------------------------------------
  
  qualidade_ano <- list()
  categorias_ano <- list()
  problemas_ano <- list()
  formatos_ano <- list()
  casos_ano <- list()
  
  adicionar_problema <- function(
    variavel, problema, quantidade
  ) {
    
    problemas_ano[[length(problemas_ano) + 1]] <<-
      tibble(
        ano = ano,
        aba = aba,
        variavel = variavel,
        problema = problema,
        quantidade = as.integer(quantidade)
      )
  }
  
  adicionar_casos <- function(
    variavel, problema, indices, valores
  ) {
    
    indices <- which(indices %in% TRUE)
    
    if (length(indices) == 0) {
      return(invisible(NULL))
    }
    
    casos_ano[[length(casos_ano) + 1]] <<-
      tibble(
        ano = ano,
        aba = aba,
        linha_excel = indices + 1L,
        variavel = variavel,
        problema = problema,
        valor_original = as.character(valores[indices])
      )
    
  }
  
  for (j in seq_len(p)) {
    
    x <- as.character(df[[j]])
    
    variavel <- variaveis[j]
    nome <- nomes_originais[j]
    
    situacao <- classificar_ausencia(x)
    
    preenchido <- situacao == "PREENCHIDO"
    
    qualidade_ano[[j]] <- tibble(
      ano = ano,
      aba = aba,
      variavel = variavel,
      coluna_original = nome,
      registros = n,
      vazios = sum(situacao == "VAZIO"),
      nao_informados = sum(
        situacao == "NAO_INFORMADO"
      ),
      preenchidos = sum(preenchido),
      pct_vazios = round(
        100 * sum(situacao == "VAZIO") /
          max(n, 1),
        2
      ),
      pct_sem_informacao = round(
        100 * sum(!preenchido) /
          max(n, 1),
        2
      ),
      valores_distintos = length(
        unique(x[preenchido])
      )
    )
    
    # ----------------------------------------------------------
    # Categorias e frequencias
    # ----------------------------------------------------------
    
    variaveis_categoricas <- c(
      "sexo",
      "cor_raca",
      "status_desfecho",
      "pessoa_localizada",
      "relacao_com_vitima",
      "motivacao_desaparecimento",
      "tipo_local_encontro",
      "municipio_desaparecimento",
      "municipio_localizacao"
    )
    
    if (variavel %in% variaveis_categoricas) {
      
      tabela_frequencia <- tibble(
        valor = ifelse(
          is.na(x) | trimws(x) == "",
          "<VAZIO>",
          x
        )
      ) %>%
        count(
          valor,
          name = "frequencia",
          sort = TRUE
        ) %>%
        mutate(
          ano = ano,
          aba = aba,
          variavel = variavel,
          valor_normalizado = normalizar(valor),
          percentual = round(
            100 * frequencia / max(n, 1),
            2
          )
        ) %>%
        select(
          ano,
          aba,
          variavel,
          valor,
          valor_normalizado,
          frequencia,
          percentual
        )
      
      categorias_ano[[length(categorias_ano) + 1]] <-
        tabela_frequencia
    }
    
    # ----------------------------------------------------------
    # Idade
    # ----------------------------------------------------------
    
    if (variavel == "idade") {
      
      idade <- converter_numero(x)
      
      nao_numerica <- (
        preenchido &
          is.na(idade)
      )
      
      fora_faixa <- (
        !is.na(idade) &
          (idade < IDADE_MIN | idade > IDADE_MAX)
      )
      
      decimal <- (
        !is.na(idade) &
          abs(idade - round(idade)) > 0.000001
      )
      
      adicionar_problema(
        variavel,
        "Idade nao numerica",
        sum(nao_numerica)
      )
      
      adicionar_problema(
        variavel,
        "Idade fora da faixa",
        sum(fora_faixa)
      )
      
      adicionar_problema(
        variavel,
        "Idade decimal",
        sum(decimal)
      )
      
      adicionar_casos(
        variavel,
        "Idade nao numerica",
        nao_numerica,
        x
      )
      
      adicionar_casos(
        variavel,
        "Idade fora da faixa",
        fora_faixa,
        x
      )
    }
    
    # ----------------------------------------------------------
    # Datas
    # ----------------------------------------------------------
    
    variaveis_data <- c(
      "data_desaparecimento",
      "data_registro",
      "data_localizacao",
      "data_nascimento"
    )
    
    if (variavel %in% variaveis_data) {
      
      datas <- converter_data(x)
      
      formatos <- formato_data(x)
      
      tab_formatos <- tibble(
        formato = formatos
      ) %>%
        count(
          formato,
          name = "quantidade"
        ) %>%
        mutate(
          ano = ano,
          aba = aba,
          variavel = variavel
        ) %>%
        select(
          ano,
          aba,
          variavel,
          formato,
          quantidade
        )
      
      formatos_ano[[length(formatos_ano) + 1]] <-
        tab_formatos
      
      data_invalida <- (
        preenchido &
          is.na(datas)
      )
      
      data_futura <- (
        !is.na(datas) &
          datas > DATA_CORTE
      )
      
      adicionar_problema(
        variavel,
        "Data nao interpretavel",
        sum(data_invalida)
      )
      
      adicionar_problema(
        variavel,
        "Data posterior ao corte",
        sum(data_futura)
      )
      
      adicionar_casos(
        variavel,
        "Data nao interpretavel",
        data_invalida,
        x
      )
      
      adicionar_casos(
        variavel,
        "Data posterior ao corte",
        data_futura,
        x
      )
    }
    
    # ----------------------------------------------------------
    # Categorias potencialmente invalidas
    # ----------------------------------------------------------
    
    if (variavel == "sexo") {
      
      validos <- c(
        "MASCULINO",
        "FEMININO",
        "NAO INFORMADO",
        "NAO INFORMADA",
        "NI"
      )
      
      suspeito <- (
        preenchido &
          !normalizar(x) %in% validos
      )
      
      adicionar_problema(
        variavel,
        "Sexo fora das categorias esperadas",
        sum(suspeito)
      )
      
      adicionar_casos(
        variavel,
        "Sexo fora das categorias esperadas",
        suspeito,
        x
      )
    }
    
    if (variavel == "cor_raca") {
      
      validos <- c(
        "BRANCA", "BRANCO",
        "PRETA", "PRETO",
        "PARDA", "PARDO",
        "AMARELA", "AMARELO",
        "INDIGENA",
        "NEGRA", "NEGRO",
        "NAO INFORMADA",
        "NAO INFORMADO",
        "NI"
      )
      
      # Retira prefixos numericos, como 2-PARDA.
      cor_limpa <- normalizar(x)
      
      cor_limpa <- gsub(
        "^[0-9]+ ",
        "",
        cor_limpa
      )
      
      suspeito <- (
        preenchido &
          !cor_limpa %in% validos
      )
      
      adicionar_problema(
        variavel,
        "Cor/raca fora das categorias esperadas",
        sum(suspeito)
      )
      
      adicionar_casos(
        variavel,
        "Cor/raca fora das categorias esperadas",
        suspeito,
        x
      )
    }
    
    if (variavel == "pessoa_localizada") {
      
      validos <- c("SIM", "NAO")
      
      suspeito <- (
        preenchido &
          !normalizar(x) %in% validos
      )
      
      adicionar_problema(
        variavel,
        "Indicador de localizacao nao binario",
        sum(suspeito)
      )
      
      adicionar_casos(
        variavel,
        "Indicador de localizacao nao binario",
        suspeito,
        x
      )
    }
  }
  
  # ------------------------------------------------------------
  # 5.4 CONSISTENCIA ENTRE DATAS
  # ------------------------------------------------------------
  
  obter_coluna <- function(variavel) {
    
    indice <- which(variaveis == variavel)
    
    if (length(indice) == 0) {
      return(rep(NA_character_, n))
    }
    
    as.character(df[[indice[1]]])
  }
  
  desaparecimento <- converter_data(
    obter_coluna("data_desaparecimento")
  )
  
  registro <- converter_data(
    obter_coluna("data_registro")
  )
  
  localizacao <- converter_data(
    obter_coluna("data_localizacao")
  )
  
  nascimento <- converter_data(
    obter_coluna("data_nascimento")
  )
  
  idade <- converter_numero(
    obter_coluna("idade")
  )
  
  # Registro anterior ao desaparecimento
  
  registro_anterior <- (
    !is.na(desaparecimento) &
      !is.na(registro) &
      registro < desaparecimento
  )
  
  adicionar_problema(
    "cronologia",
    "Registro anterior ao desaparecimento",
    sum(registro_anterior)
  )
  
  adicionar_casos(
    "cronologia",
    "Registro anterior ao desaparecimento",
    registro_anterior,
    obter_coluna("data_registro")
  )
  
  # Localizacao anterior ao desaparecimento
  
  localizacao_anterior <- (
    !is.na(desaparecimento) &
      !is.na(localizacao) &
      localizacao < desaparecimento
  )
  
  adicionar_problema(
    "cronologia",
    "Localizacao anterior ao desaparecimento",
    sum(localizacao_anterior)
  )
  
  adicionar_casos(
    "cronologia",
    "Localizacao anterior ao desaparecimento",
    localizacao_anterior,
    obter_coluna("data_localizacao")
  )
  
  # Registro muito posterior ao desaparecimento
  # Apenas sinalizacao exploratoria (>365 dias).
  
  atraso_registro <- as.numeric(
    registro - desaparecimento
  )
  
  registro_tardio <- (
    !is.na(atraso_registro) &
      atraso_registro > 365
  )
  
  adicionar_problema(
    "cronologia",
    "Registro mais de 365 dias depois",
    sum(registro_tardio)
  )
  
  # ------------------------------------------------------------
  # 5.5 IDADE E NASCIMENTO
  # ------------------------------------------------------------
  
  if ("data_nascimento" %in% variaveis) {
    
    referencia <- desaparecimento
    
    # Nao substituir a data de desaparecimento
    # pela de registro para evitar comparacoes
    # entre conceitos diferentes.
    
    idade_calculada <- as.numeric(
      referencia - nascimento
    ) / 365.2425
    
    nascimento_posterior <- (
      !is.na(nascimento) &
        !is.na(referencia) &
        nascimento > referencia
    )
    
    divergencia <- (
      !is.na(idade) &
        !is.na(idade_calculada) &
        abs(idade - idade_calculada) >
        TOLERANCIA_IDADE
    )
    
    adicionar_problema(
      "idade",
      "Nascimento posterior ao desaparecimento",
      sum(nascimento_posterior)
    )
    
    adicionar_problema(
      "idade",
      "Divergencia entre idade e nascimento",
      sum(divergencia)
    )
    
    adicionar_casos(
      "idade",
      "Divergencia entre idade e nascimento",
      divergencia,
      obter_coluna("idade")
    )
  }
  
  # ------------------------------------------------------------
  # 5.6 CONSISTENCIA DO INDICADOR DE LOCALIZACAO
  # ------------------------------------------------------------
  
  if ("pessoa_localizada" %in% variaveis) {
    
    localizada <- normalizar(
      obter_coluna("pessoa_localizada")
    )
    
    status <- normalizar(
      obter_coluna("status_desfecho")
    )
    
    status_localizado <- grepl(
      "LOCALIZADO|ENCONTRADO",
      status
    )
    
    contradicao <- (
      localizada == "NAO" &
        status_localizado
    )
    
    contradicao[is.na(contradicao)] <- FALSE
    
    adicionar_problema(
      "pessoa_localizada",
      "Nao localizado com status de encontrado",
      sum(contradicao)
    )
    
    adicionar_casos(
      "pessoa_localizada",
      "Nao localizado com status de encontrado",
      contradicao,
      obter_coluna("pessoa_localizada")
    )
  }
  
  # ------------------------------------------------------------
  # 5.7 CONSOLIDAR RESULTADOS DA ABA
  # ------------------------------------------------------------
  
  lista_qualidade[[aba]] <- bind_rows(
    qualidade_ano
  )
  
  lista_categorias[[aba]] <- bind_rows(
    categorias_ano
  )
  
  lista_problemas[[aba]] <- bind_rows(
    problemas_ano
  )
  
  lista_formatos[[aba]] <- bind_rows(
    formatos_ano
  )
  
  lista_casos[[aba]] <- bind_rows(
    casos_ano
  )
  
  # ------------------------------------------------------------
  # 5.8 IMPRESSAO RESUMIDA
  # ------------------------------------------------------------
  
  mostrar(
    lista_estrutura[[aba]] %>%
      select(
        posicao,
        nome_original,
        variavel,
        tipo_readxl
      ),
    paste("ESTRUTURA:", ano)
  )
  
  mostrar(
    lista_qualidade[[aba]] %>%
      select(
        variavel,
        pct_sem_informacao,
        valores_distintos
      ),
    paste("COMPLETUDE:", ano)
  )
  
  mostrar(
    lista_problemas[[aba]] %>%
      filter(quantidade > 0) %>%
      arrange(desc(quantidade)),
    paste("PROBLEMAS:", ano)
  )
}


# ================================================================
# 6. CONSOLIDACAO
# ================================================================

resumo <- bind_rows(lista_resumo)

estrutura <- bind_rows(lista_estrutura)

qualidade <- bind_rows(lista_qualidade)

categorias <- bind_rows(lista_categorias)

problemas <- bind_rows(lista_problemas)

formatos <- bind_rows(lista_formatos)

casos <- bind_rows(lista_casos)


# ================================================================
# 7. MATRIZ DE PRESENCA
# ================================================================

presenca <- estrutura %>%
  distinct(
    ano,
    variavel
  ) %>%
  mutate(presente = "SIM") %>%
  pivot_wider(
    names_from = ano,
    values_from = presente,
    values_fill = "NAO",
    names_prefix = "ano_"
  )

for (ano in ANOS) {
  
  coluna <- paste0("ano_", ano)
  
  if (!coluna %in% names(presenca)) {
    presenca[[coluna]] <- "SEM ABA"
  }
}

presenca <- presenca %>%
  select(
    variavel,
    all_of(paste0("ano_", ANOS))
  ) %>%
  arrange(variavel)


# ================================================================
# 8. COMPARACAO DA COMPLETUDE
# ================================================================

# Mantem colunas separadas quando eventualmente
# mais de uma corresponder a mesma variavel/ano.

comparacao_qualidade <- qualidade %>%
  group_by(
    ano,
    variavel
  ) %>%
  summarise(
    pct_sem_informacao = round(
      weighted.mean(
        pct_sem_informacao,
        w = registros,
        na.rm = TRUE
      ),
      2
    ),
    .groups = "drop"
  ) %>%
  pivot_wider(
    names_from = ano,
    values_from = pct_sem_informacao,
    names_prefix = "ano_"
  )


# ================================================================
# 9. SINTETIZAR PROBLEMAS
# ================================================================

problemas_resumo <- problemas %>%
  group_by(
    ano,
    variavel,
    problema
  ) %>%
  summarise(
    quantidade = sum(quantidade),
    .groups = "drop"
  ) %>%
  arrange(
    ano,
    desc(quantidade)
  )


# ================================================================
# 10. COMPARACAO DE CATEGORIAS
# ================================================================

categorias_comparacao <- categorias %>%
  group_by(
    variavel,
    valor_normalizado,
    ano
  ) %>%
  summarise(
    frequencia = sum(frequencia),
    .groups = "drop"
  ) %>%
  pivot_wider(
    names_from = ano,
    values_from = frequencia,
    values_fill = 0,
    names_prefix = "ano_"
  ) %>%
  arrange(
    variavel,
    valor_normalizado
  )


# ================================================================
# 11. RESULTADOS FINAIS NO CONSOLE
# ================================================================

mostrar(
  resumo,
  "RESUMO GERAL - 2020 A 2026"
)

mostrar(
  presenca,
  "MATRIZ DE PRESENCA DAS VARIAVEIS"
)

mostrar(
  comparacao_qualidade,
  "COMPARACAO DA AUSENCIA DE INFORMACOES"
)

mostrar(
  formatos,
  "FORMATOS DAS DATAS"
)

mostrar(
  problemas_resumo %>%
    filter(quantidade > 0),
  "PROBLEMAS IDENTIFICADOS"
)

titulo("CATEGORIAS MAIS FREQUENTES POR ANO")

for (ano in sort(unique(categorias$ano))) {
  
  cat("\nANO:", ano, "\n")
  
  variaveis_ano <- unique(
    categorias$variavel[categorias$ano == ano]
  )
  
  for (v in variaveis_ano) {
    
    cat("\nVariavel:", v, "\n")
    
    tab <- categorias %>%
      filter(
        ano == !!ano,
        variavel == v
      ) %>%
      arrange(desc(frequencia))
    
    print(
      as.data.frame(
        head(tab, TOP_CATEGORIAS)
      )[, c("valor", "frequencia", "percentual")],
      row.names = FALSE
    )
  }
}


# ================================================================
# 12. EXPORTACAO - UM UNICO XLSX
# ================================================================

if (GERAR_XLSX) {
  
  caminho_xlsx <- file.path(
    PASTA_SAIDA,
    "auditoria_qualidade_desaparecidos_2020_2026.xlsx"
  )
  
  tabelas <- list(
    Resumo = resumo,
    Estrutura = estrutura,
    Presenca = presenca,
    Qualidade = qualidade,
    Comparacao = comparacao_qualidade,
    Categorias = categorias,
    Categorias_Anos = categorias_comparacao,
    Formatos_Datas = formatos,
    Problemas = problemas_resumo,
    Casos = casos
  )
  
  wb <- createWorkbook()
  
  for (nome in names(tabelas)) {
    
    addWorksheet(wb, nome)
    
    writeData(
      wb,
      sheet = nome,
      x = tabelas[[nome]],
      withFilter = nrow(tabelas[[nome]]) > 0
    )
    
    freezePane(
      wb,
      sheet = nome,
      firstRow = TRUE
    )
    
    setColWidths(
      wb,
      sheet = nome,
      cols = seq_len(ncol(tabelas[[nome]])),
      widths = "auto"
    )
  }
  
  saveWorkbook(
    wb,
    file = caminho_xlsx,
    overwrite = TRUE
  )
  
  cat("\nXLSX salvo em:\n", caminho_xlsx, "\n")
}


# ================================================================
# 13. README.MD
# ================================================================

if (GERAR_README) {
  
  caminho_readme <- file.path(
    PASTA_SAIDA,
    "README.md"
  )
  
  texto <- c(
    "# Auditoria da qualidade dos dados de pessoas desaparecidas",
    "",
    "## Fonte",
    "",
    basename(ARQUIVO),
    "",
    "## Período analisado",
    "",
    "2020 a 2026.",
    "",
    "## Objetivo",
    "",
    paste(
      "Investigar a estrutura, a completude e a consistência",
      "das informações anuais de pessoas desaparecidas."
    ),
    "",
    "## Procedimentos",
    "",
    "1. Inventário das colunas e tipos.",
    "2. Comparação estrutural entre os anos.",
    "3. Identificação de valores vazios e não informados.",
    "4. Análise de categorias e frequências.",
    "5. Verificação de formatos de datas.",
    "6. Avaliação de idade e data de nascimento.",
    "7. Consistência cronológica.",
    "8. Identificação de registros integralmente duplicados.",
    "9. Registro de casos suspeitos.",
    "",
    "## Observações metodológicas",
    "",
    paste(
      "As correspondências entre colunas são provisórias.",
      "A equivalência conceitual precisa ser confirmada."
    ),
    "",
    paste(
      "Os problemas identificados são sinalizações",
      "e não implicam exclusão automática de registros."
    ),
    "",
    paste(
      "As idades decimais não são consideradas inválidas",
      "apenas por apresentarem casas decimais."
    ),
    "",
    paste(
      "A ausência de informações de localização pode",
      "decorrer do status da ocorrência e precisa",
      "ser interpretada condicionalmente."
    ),
    "",
    paste(
      "A comparação de idades com nascimento utiliza",
      "a data do desaparecimento como referência."
    ),
    "",
    paste(
      "O total de registros corresponde às linhas das",
      "abas e não necessariamente a pessoas únicas."
    ),
    "",
    "## Resultados gerais",
    "",
    paste(
      "Abas analisadas:",
      nrow(mapa_abas)
    ),
    paste(
      "Registros:",
      sum(resumo$registros)
    ),
    paste(
      "Variáveis propostas:",
      nrow(presenca)
    ),
    paste(
      "Tipos de problema com ocorrência:",
      nrow(problemas_resumo %>%
             filter(quantidade > 0))
    ),
    "",
    "## Próxima etapa",
    "",
    paste(
      "Validar o dicionário de variáveis, estabelecer",
      "regras de harmonização e documentar os problemas",
      "que afetam a comparabilidade temporal."
    )
  )
  
  writeLines(
    texto,
    caminho_readme,
    useBytes = TRUE
  )
  
  cat("\nREADME salvo em:\n", caminho_readme, "\n")
}


# ================================================================
# 14. FINALIZACAO
# ================================================================

titulo("AUDITORIA CONCLUIDA")

cat(
  "Anos analisados:",
  paste(
    sort(unique(resumo$ano)),
    collapse = ", "
  ),
  "\n"
)

cat(
  "Total de registros:",
  sum(resumo$registros),
  "\n"
)

cat(
  "Total de ocorrencias sinalizadas:",
  sum(problemas_resumo$quantidade),
  "\n"
)

cat(
  "\nO Excel original foi preservado.\n"
)