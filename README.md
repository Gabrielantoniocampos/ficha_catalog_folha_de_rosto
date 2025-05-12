# Geração automática do README.txt explicativo
conteudo_readme = """
README - Geração Automatizada de Fichas Catalográficas e Modelos Word
Autor: [Seu Nome ou Equipe]
Versão: 1.0
Data: [Data de geração]

Objetivo
--------
Este script automatiza a criação de fichas catalográficas e documentos modelo a partir de uma planilha Excel contendo dados bibliográficos. Utiliza dois arquivos .docx como templates, substitui os placeholders com dados reais e salva os documentos organizados em pastas nomeadas por autor e título.

Requisitos
----------
Instale os pacotes necessários com:
    pip install pandas openpyxl python-docx

Arquivos Utilizados
-------------------
- BASE_DE_FICHAS_PARA_AUTOMAÇÃO.xlsx: Dados bibliográficos por linha.
- Modelo_WBA0048_v3_Liderança_FC.docx: Template da ficha catalográfica.
- Modelo_WBA0048_v3_Liderança.docx: Template do documento base.
- fichas_final_com_tabelas/: Pasta onde são salvos os documentos gerados.
- fichas_e_modelos.zip: Arquivo ZIP contendo todos os documentos criados.

Descrição Técnica do Processo
-----------------------------
1. A planilha é carregada e lida com pandas.
2. O nome do autor é formatado como "Sobrenome, Nome".
3. Os dados são usados para substituir placeholders como <AUTOR>, <TITULO>, <SUBTITULO> etc.
4. Campos como subtítulo e volume são tratados para evitar formatações quebradas.
5. Uma pasta é criada para cada linha da planilha com os arquivos:
    - Ficha Catalográfica.docx
    - Modelo Preenchido.docx
6. Ao final, todos os arquivos são compactados em fichas_e_modelos.zip.

Observações
-----------
- Volume tratado como inteiro, ex.: ": v. 2"
- Subtítulo e volume são ocultados se estiverem vazios.
- Os nomes de pastas e arquivos são higienizados para evitar erros de sistema.

Contato
-------
[Insira aqui e-mail ou referência do autor, se necessário]
"""

# Salva o README na pasta de saída
with open(os.path.join(pasta_saida, "README.txt"), "w", encoding="utf-8") as f:
    f.write(conteudo_readme.strip())
