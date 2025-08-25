library(pacman)
pacman::p_load(purrr,
  dplyr,#manipulação de dataframe
               stopwords,#limpeza de texto, deixar em pt
               xml2, #trabalhar com HTML e XML
               tidytext,#mineração de texto
               lubridate,#manipulação de data
               stringi,#manipulação de string/texto
               stringr,#manipulação de string/texto
               writexl,#cria tabela
               openxlsx,#estruturar a tabela 
               )


# Carregar pacotes necessários

# Função para analisar o feed RSS e criar um DataFrame
get_rss_data <- function(url) {
  feed <- xml2::read_xml(url)
  items <- xml2::xml_find_all(feed, "//item")
  
  data <- tibble(
    Title = xml2::xml_text(xml_find_all(items, "title")),
    Link = xml2::xml_text(xml_find_all(items, "link")),
    Published = xml2::xml_text(xml_find_all(items, "pubDate")),
    Description = xml2::xml_text(xml_find_all(items, "description"))
  )
  return(data)
}

# URLs dos feeds RSS
urls <- c(
  "https://portal.who.int/eios/API/News/Monitoring/getBoardRssFeed?queryId=1780",
  "https://portal.who.int/eios/API/News/Monitoring/getBoardRssFeed?queryId=1781",
  "https://portal.who.int/eios/API/News/Monitoring/getBoardRssFeed?queryId=10922",
)



# Nomes dos termos associados a cada URL
termos <- c(
  '24H Nacional',
  '24H Internacional',
  'plantão_CIEVS'
)




# Baixar e combinar os dados dos feeds RSS
df_list <- purrr::map(urls, get_rss_data)
df_list <- purrr::map2(df_list, termos, ~ mutate(.x, Termo = .y))
df_final <- bind_rows(df_list)

# Função para limpar texto
limpar_texto <- function(texto) {
  texto <- stringi::stri_trans_general(texto, "Latin-ASCII")
  texto <- stringr::str_remove_all(texto, "[^a-zA-Z0-9\\s]")
  palavras <- stringr::str_split(texto, "\\s+")[[1]]
  palavras <- palavras[!tolower(palavras) %in% stopwords("pt")]
  return(paste(palavras, collapse = " "))
}

# Limpar o título
df_final <- df_final %>%
  dplyr::mutate(Title_limpo = map_chr(Title, limpar_texto))

# Lista de palavras-chave
palavras_chave <- c("casos", "contra", "alerta", "doenca", "aumento", "grande", "emergencia", "numero", "situacao",
                    "mortes", "caso", "vacina","vacinacao","campanha vacinacao","doses", "registra", "maior", "chuvas", "dados", "surto", "confirmados","gripe","virus respiratorios",
                    "confirma", "morte", "obitos", "obito", "virus", "alta", "risco", "vigilancia", "seca", "leitos", 
                    "registrou", "chuva", "causa", "queimada", "queimadas", "estiagem", "quente", "hospitalizacao", 
                    "internacao", "SRAG","VSR", "influenza aviaria","aviaria","H5N1","rinovirus","covid","infogripe","gripe aviária","EUA","Síndrome Respiratória Aguda Grave","sobrecarregam","sobrecarregar")


# Filter news based on keywords in the title only

# df_final <- df_final[grepl(paste(palavras_chave, collapse = "|"), df_final$Title, ignore.case = TRUE), ]


# Função para calcular a pontuação com base nas palavras-chave
calcular_pontuacao <- function(texto) {
  texto <- stringi::stri_trans_general(texto, "Latin-ASCII")
  texto <- stringr::str_remove_all(texto, "[^a-zA-Z0-9\\s]")
  pontuacao <- sum(str_detect(texto, palavras_chave))
  return(pontuacao)
}

# Calcular pontuação, filtrar resultados e remover duplicatas
df_final <- df_final %>%
  dplyr::mutate(Pontuacao = map_int(Title_limpo, calcular_pontuacao)) %>%
  dplyr::arrange(desc(Pontuacao))


df_final_filtrado <- df_final %>%
  dplyr::filter(Pontuacao >= 1) %>%
  dplyr::distinct(Title, .keep_all = TRUE)



# Atualizar nomes das colunas e remover a coluna 'Title_limpo'
df_final_filtrado <- df_final_filtrado %>%
  dplyr::select(
    Título = Title,
    Link,
    `Data da Publicação` = Published,
    Descrição = Description,
    Categoria = Termo)




# # Assuming df_final_filtrado is your existing dataframe
# titles <- df_final_filtrado$Título
# 
# # Preprocess the text
# preprocess_text <- function(text) {
#   text <- tolower(text)
#   text <- removePunctuation(text)
#   text <- removeWords(text, stopwords("portuguese"))
#   return(text)
# }
# 
# titles_clean <- sapply(titles, preprocess_text)
# 
# # Calculate Jaccard similarity
# similarity_matrix <- stringdistmatrix(titles_clean, titles_clean, method = "jaccard")
# 
# # Set a threshold for similarity (e.g., 0.2)
# threshold <- 0.2
# 
# # Find duplicates
# duplicates <- which(similarity_matrix < threshold, arr.ind = TRUE)
# duplicates <- duplicates[duplicates[,1] != duplicates[,2],]
# 
# # Include length of the titles in the comparison
# title_lengths <- nchar(titles)
# 
# # Function to keep the longer title
# keep_longer_title <- function(dup_indices, lengths) {
#   to_remove <- c()
#   for (i in 1:nrow(dup_indices)) {
#     if (lengths[dup_indices[i, 1]] >= lengths[dup_indices[i, 2]]) {
#       to_remove <- c(to_remove, dup_indices[i, 2])
#     } else {
#       to_remove <- c(to_remove, dup_indices[i, 1])
#     }
#   }
#   return(unique(to_remove))
# }
# 
# # Get indices of titles to remove
# to_remove <- keep_longer_title(duplicates, title_lengths)
# 
# # Remove duplicates
# unique_titles <- titles[-to_remove]

# Create a new dataframe with unique titles
# df_unique <- df_final_filtrado %>% filter(Título %in% unique_titles)


df_final_filtrado <- df_final_filtrado[, c("Categoria", "Título", "Link", "Data da Publicação", "Descrição")]




df_final_filtrado$Descrição <- stringr::str_wrap(df_final_filtrado$Descrição, width = 60)

wb <- openxlsx::createWorkbook()
openxlsx::addWorksheet(wb, "Sheet1")

openxlsx::writeData(wb, "Sheet1", df_final_filtrado, startCol = 1, startRow = 1)


for (i in seq_along(df_final_filtrado$Link)) {
  openxlsx::writeFormula(
    wb, "Sheet1",
    startRow = i + 1, startCol = 3,
    x = sprintf('=HYPERLINK("%s", "%s")',  df_final_filtrado$Link[i], df_final_filtrado$Link[i])
  )
}



data_atual <- format(Sys.Date(), "%d-%m-%Y")

# Definir o caminho completo do diretório e nome do arquivo
caminho <- "C:/Users/andradecle/OneDrive - Pan American Health Organization/General - PHE BRA/Rumores/Automatização/"

nome_arquivo <- paste0(caminho, "Rumores_", data_atual, ".xlsx")

# Exportar o data frame para o arquivo CSV no caminho especificado

openxlsx::saveWorkbook(wb, nome_arquivo, overwrite = TRUE)
# export(df_unique, nome_arquivo)





# # Defina o arquivo original
# excel_file <- "feed_data__filtrado.xlsx"
# 
# # Defina o caminho de destino, incluindo a data do dia no nome do arquivo
# today_date <- Sys.Date()
# destination <- paste0("C:/Users/andradecle/OneDrive - Pan American Health Organization/General - PHE BRA/Rumores/Automatização/feed_data_filtrado_", today_date, ".xlsx")
# write_xlsx(df_unique, excel_file)
# # Copie o arquivo para o destino
# file.copy(excel_file, destination)
