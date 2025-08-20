library(pacman)
pacman::p_load(tidyverse,
               stopwords,
               xml2,
               httr,
               rvest,
               tidytext,
               wordcloud2,
               officer,
               lubridate,
               stringi,
               writexl,
               officer,
               zip)
# Carregar pacotes necessários

# Função para analisar o feed RSS e criar um DataFrame
get_rss_data <- function(url) {
  feed <- read_xml(url)
  items <- xml_find_all(feed, "//item")
  
  data <- tibble(
    Title = xml_text(xml_find_all(items, "title")),
    Link = xml_text(xml_find_all(items, "link")),
    Published = xml_text(xml_find_all(items, "pubDate")),
    Description = xml_text(xml_find_all(items, "description"))
  )
  return(data)
}

# URLs dos feeds RSS
urls <- c(
  "https://portal.who.int/eios/API/News/Monitoring/getBoardRssFeed?queryId=12760",
  "https://portal.who.int/eios/API/News/Monitoring/getBoardRssFeed?queryId=13428",
  "https://portal.who.int/eios/API/News/Monitoring/getBoardRssFeed?queryId=12847",
  "https://portal.who.int/eios/API/News/Monitoring/getBoardRssFeed?queryId=13620",
  "https://portal.who.int/eios/API/News/Monitoring/getBoardRssFeed?queryId=11151",
  "https://portal.who.int/eios/API/News/Monitoring/getBoardRssFeed?queryId=13462",
  "https://portal.who.int/eios/API/News/Monitoring/getBoardRssFeed?queryId=2478",
  "https://portal.who.int/eios/API/News/Monitoring/getBoardRssFeed?queryId=14901",
  "https://portal.who.int/eios/API/News/Monitoring/getBoardRssFeed?queryId=13891"
)

# Nomes dos termos associados a cada URL
termos <- c(
  'Surtos e Emergências',
  'Secas e Queimadas',
  'Vírus Respiratorios',
  'Ameaças Naturais',
  'Arboviroses',
  'Refugiados Israel',
  'Outros',
  'Doenças Gerais',
  'Sintomas'
)

# Baixar e combinar os dados dos feeds RSS
df_list <- map(urls, get_rss_data)
df_list <- map2(df_list, termos, ~ mutate(.x, Termo = .y))
df_final <- bind_rows(df_list)

# Função para limpar texto
limpar_texto <- function(texto) {
  texto <- stri_trans_general(texto, "Latin-ASCII")
  texto <- str_remove_all(texto, "[^a-zA-Z0-9\\s]")
  palavras <- str_split(texto, "\\s+")[[1]]
  palavras <- palavras[!tolower(palavras) %in% stopwords("pt")]
  return(paste(palavras, collapse = " "))
}

# Limpar o título
df_final <- df_final %>%
  mutate(Title_limpo = map_chr(Title, limpar_texto))

# Lista de palavras-chave
palavras_chave <- c("casos", "contra", "alerta", "doenca", "aumento", "grande", "emergencia", "numero", "situacao",
                    "mortes", "caso", "vacina", "registra", "maior", "chuvas", "dados", "surto", "confirmados",
                    "confirma", "morte","obitos","obito", "virus", "alta", "risco", "vigilancia", "seca", "leitos", "registrou",
                    "chuva", "causa","queimada","queimadas","estiagem","quente","hospitalizacao","internacao","SRAG")

# Função para calcular a pontuação com base nas palavras-chave
calcular_pontuacao <- function(texto) {
  texto <- stri_trans_general(texto, "Latin-ASCII")
  texto <- str_remove_all(texto, "[^a-zA-Z0-9\\s]")
  pontuacao <- sum(str_detect(texto, palavras_chave))
  return(pontuacao)
}

# Calcular pontuação e filtrar resultados
df_final <- df_final %>%
  mutate(Pontuacao = map_int(Title_limpo, calcular_pontuacao)) %>%
  arrange(desc(Pontuacao))

df_final_filtrado <- df_final %>% filter(Pontuacao >= 3)


# Lista de estados, capitais, siglas e outros termos em uma linha
estados_capitais_siglas <- c(
  'Acre', 'Rio Branco', 'AC', 'Alagoas', 'Maceio', 'AL', 'Amapa', 'Macapa', 'AP', 'Amazonas', 'Manaus', 'AM',
  'Bahia', 'Salvador', 'BA', 'Ceara', 'Fortaleza', 'CE', 'Distrito Federal', 'Brasilia', 'DF', 'Espirito Santo', 'Vitoria', 'ES',
  'Goias', 'Goiania', 'GO', 'Maranhao', 'Sao Luis', 'MA', 'Mato Grosso', 'Cuiaba', 'MT', 'Mato Grosso do Sul', 'Campo Grande', 'MS',
  'Minas Gerais', 'Belo Horizonte', 'MG', 'Para', 'Belem', 'PA', 'Paraiba', 'Joao Pessoa', 'PB', 'Parana', 'Curitiba', 'PR',
  'Pernambuco', 'Recife', 'PE', 'Piaui', 'Teresina', 'PI', 'Rio de Janeiro', 'Rio de Janeiro', 'RJ', 'Rio Grande do Norte', 'Natal', 'RN',
  'Rio Grande do Sul', 'Porto Alegre', 'RS', 'Rondonia', 'Porto Velho', 'RO', 'Roraima', 'Boa Vista', 'RR', 'Santa Catarina', 'Florianopolis', 'SC',
  'Sao Paulo', 'Sao Paulo', 'SP', 'Sergipe', 'Aracaju', 'SE', 'Tocantins', 'Palmas', 'TO', 'Brasil', 'Brasilia', 'BR', 'Ministerio da Saude','Governo'
)

# Função para verificar se o título contém algum estado, capital ou sigla
contém_termos <- function(titulo, termos) {
  any(str_detect(titulo, termos))
}

# Atualizar o DataFrame com a coluna booleana
df_final_filtrado <- df_final_filtrado %>%
  mutate(
    Contém_Termos = map_lgl(Title_limpo, ~ contém_termos(.x, estados_capitais_siglas))
  ) %>%
  filter(Contém_Termos) %>%
  select(-Contém_Termos) %>%  # Remover a coluna booleana se não for mais necessária
  distinct(Title_limpo, .keep_all = TRUE)  # Remover duplicatas com base na coluna Title_limpo



library(officer)
library(dplyr)
library(stringr)

library(officer)
library(dplyr)
library(stringr)

# Carregar o modelo existente
doc <- read_docx("Modelo_word.docx")

# Configurar o estilo para a data com Franklin Gothic e negrito
data_estilo <- fp_text(font.size = 12, font.family = "Franklin Gothic", bold = TRUE)

# Obter a data atual e formatá-la com o mês em maiúsculas
data_atual <- format(Sys.Date(), "%d de %B")
data_atual <- str_to_upper(data_atual)  # Colocar o mês em maiúsculas

# Adicionar a data antes das notícias, alinhada à direita, com o estilo configurado
doc <- doc %>%
  body_add_fpar(
    fpar(
      ftext(data_atual, prop = data_estilo),
      fp_p = fp_par(text.align = "right")
    )
  ) %>%
  body_add_par("", style = "Normal")  # Espaço após a data

# Imagens específicas para cada categoria
imagens_categorias <- list(
  'Surtos e Emergências' = "logo_surtos.png",
  'Secas e Queimadas' = "logo_secas_queimadas.png",
  'Vírus Respiratorios' = "logo_virus_respiratorios.png",
  'Ameaças Naturais' = "logo_ameacas_naturais.png",
  'Arboviroses' = "logo_arboviroses.png",
  'Refugiados Israel' = "logo_refugiados_israel.png",
  'Outros' = "logo_outros.png",
  'Doenças Gerais' = "logo_doencas_gerais.png",
  'Sintomas' = "logo_sintomas.png"
)

# Largura e altura desejada para as imagens
largura_imagem_cm <- 15.0  # Ajuste a largura conforme necessário (em cm)
altura_imagem_cm <- 1.75   # Altura desejada em cm

# Função para adicionar imagens com tamanho fixo
adicionar_imagem <- function(doc, imagem_path, largura_cm, altura_cm) {
  largura_inch <- largura_cm / 2.54
  altura_inch <- altura_cm / 2.54
  doc <- body_add_img(doc, src = imagem_path, width = largura_inch * 72, height = altura_inch * 72)
  return(doc)
}

# Adicionar as notícias divididas por categorias
for (termo in termos) {
  # Filtrar as notícias dessa categoria
  noticias_categoria <- df_final_filtrado %>% filter(Termo == termo)
  
  # Se houver notícias na categoria, adiciona a imagem da categoria e as notícias
  if (nrow(noticias_categoria) > 0) {
    # Adicionar a imagem correspondente à categoria
    imagem_categoria <- imagens_categorias[[termo]]
    if (!is.null(imagem_categoria)) {
      doc <- adicionar_imagem(doc, imagem_categoria, largura_imagem_cm, altura_imagem_cm)
    }
    
    for (i in seq_len(nrow(noticias_categoria))) {
      noticia <- noticias_categoria[i, ]
      
      # Adicionar título, descrição e link com estilos personalizados
      doc <- doc %>%
        body_add_fpar(fpar(ftext(noticia$Title, prop = titulo_estilo))) %>%
        body_add_fpar(fpar(ftext(noticia$Description, prop = descricao_estilo))) %>%
        body_add_fpar(fpar(ftext("Link", prop = link_estilo), hyperlink = noticia$Link)) %>%
        body_add_par("", style = "Normal")  # Adiciona uma linha em branco para separação
    }
  }
}

# Salvar o documento modificado
print(doc, target = "noticias_filtradas_modelo.docx")

# Carregar o modelo existente
doc <- read_docx("Modelo_word.docx")

# Configurar o estilo para a data com Franklin Gothic e negrito
data_estilo <- fp_text(font.size = 12, font.family = "Franklin Gothic", bold = TRUE)

# Configurar o estilo para o título
titulo_estilo <- fp_text(font.size = 15, font.family = "Franklin Gothic", bold = TRUE)

# Configurar o estilo para a descrição
descricao_estilo <- fp_text(font.size = 11, font.family = "Franklin Gothic")

# Configurar o estilo para o link
link_estilo <- fp_text(font.size = 10, font.family = "Franklin Gothic", color = "blue", underline = TRUE)

# Obter a data atual e formatá-la com o mês em maiúsculas
data_atual <- format(Sys.Date(), "%d de %B")
data_atual <- str_to_upper(data_atual)  # Colocar o mês em maiúsculas

# Adicionar a data antes das notícias, alinhada à direita, com o estilo configurado
doc <- doc %>%
  body_add_fpar(
    fpar(
      ftext(data_atual, prop = data_estilo),
      fp_p = fp_par(text.align = "right")
    )
  ) %>%
  body_add_par("", style = "Normal")  # Espaço após a data

# Mapear categorias para imagens
imagens <- c(
  'Surtos e Emergências' = "logo_surtos.png",
  'Secas e Queimadas' = "logo_secas_queimadas.png",
  'Vírus Respiratorios' = "logo_virus_respiratorios.png",
  'Ameaças Naturais' = "logo_ameacas_naturais.png",
  'Arboviroses' = "logo_arboviroses.png",
  'Refugiados Israel' = "logo_refugiados_israel.png",
  'Outros' = "logo_outros.png",
  'Doenças Gerais' = "logo_doencas_gerais.png",
  'Sintomas' = "logo_sintomas.png"
)

# Adicionar as notícias divididas por categorias
for (termo in termos) {
  # Filtrar as notícias dessa categoria
  noticias_categoria <- df_final_filtrado %>% filter(Termo == termo)
  
  # Se houver notícias na categoria, adiciona a imagem ou o título e as notícias
  if (nrow(noticias_categoria) > 0) {
    if (termo %in% names(imagens)) {
      # Adicionar a imagem correspondente à categoria
      imagem_path <- imagens[[termo]]
      doc <- doc %>%
        body_add_fpar(
          fpar(
            external_img(src = imagem_path, width = 6.5, height = 0.75)  # Ajuste a altura conforme necessário
          )
        )
    } else {
      # Adicionar o título para categorias não listadas nas imagens
      doc <- doc %>%
        body_add_par(termo, style = "heading 1")
    }
    
    doc <- doc %>%
      body_add_par("", style = "Normal") # Espaço após o título ou imagem
    
    for (i in seq_len(nrow(noticias_categoria))) {
      noticia <- noticias_categoria[i, ]
      
      # Adicionar título, descrição e link com estilos personalizados
      doc <- doc %>%
        body_add_fpar(fpar(ftext(noticia$Title, prop = titulo_estilo))) %>%
        body_add_fpar(fpar(ftext(noticia$Description, prop = descricao_estilo))) %>%
        body_add_fpar(
          fpar(
            ftext("Link", prop = link_estilo),
            hyperlink = noticia$Link  # Adiciona o hyperlink diretamente ao URL
          )
        ) %>%
        body_add_par("", style = "Normal")  # Adiciona uma linha em branco para separação
    }
  }
}

# Salvar o documento modificado
print(doc, target = "noticias_filtradas_modelo.docx")







# Salvar
excel_file <- "feed_data_filtrado.xlsx"
write_xlsx(df_final_filtrado, excel_file)




