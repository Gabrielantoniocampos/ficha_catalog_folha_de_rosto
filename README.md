# 📚 Automação de Fichas Catalográficas e Modelos Word

Este projeto automatiza a geração de fichas catalográficas e documentos modelo Word com base em dados provenientes de uma planilha Excel. Ele é útil para instituições de ensino, bibliotecas ou editoras que precisam gerar um grande volume de fichas catalográficas padronizadas.

---

## 🎯 Objetivo

A ferramenta tem como finalidade:

- Substituir automaticamente campos (placeholders) em arquivos `.docx` usando dados da planilha.
- Gerar uma ficha catalográfica e um modelo de conteúdo para cada item (linha) da planilha.
- Organizar os documentos em pastas com base no nome do autor e título.
- Compactar todo o resultado final em um único arquivo `.zip`.

---

## 🛠️ Tecnologias e Bibliotecas Usadas

- Python 3.x
- `pandas`
- `openpyxl`
- `python-docx`
- `os`, `zipfile`, `re`, `shutil`

Instalação das dependências:

```bash
pip install pandas openpyxl python-docx


# Salva o README na pasta de saída

📁 fichas_final_com_tabelas/
├── 📁 Silva, João - Liderança Transformadora/
│   ├── Ficha Catalográfica.docx
│   └── Modelo Preenchido.docx
├── 📁 Souza, Maria - Gestão de Projetos/
│   ├── Ficha Catalográfica.docx
│   └── Modelo Preenchido.docx
└── 📄 README.txt

📦 fichas_e_modelos.zip

with open(os.path.join(pasta_saida, "README.txt"), "w", encoding="utf-8") as f:
    f.write(conteudo_readme.strip())


📁 Estrutura dos Arquivos
BASE_DE_FICHAS_PARA_AUTOMAÇÃO.xlsx — Planilha com dados bibliográficos por linha.

Modelo_WBA0048_v3_Liderança_FC.docx — Template da ficha catalográfica com placeholders.

Modelo_WBA0048_v3_Liderança.docx — Template do documento modelo.

fichas_final_com_tabelas/ — Pasta gerada com os documentos individuais.

fichas_e_modelos.zip — Arquivo final compactado contendo todos os documentos.

README.txt — Explicação técnica para documentação interna.

🔄 Funcionamento
A planilha Excel é lida linha por linha.

O nome do autor é formatado no estilo “Sobrenome, Nome”.

Substituições são feitas nos templates .docx, respeitando os placeholders como:

<AUTOR>, <TITULO>, <SUBTITULO>, <COORDENADOR>, <ISBN>, etc.

Campos como :<SUBTITULO> ou : v. <VOLUME> são removidos automaticamente se estiverem vazios.

Dois arquivos .docx são gerados por linha:

Ficha Catalográfica.docx

Modelo Preenchido.docx

Os arquivos são salvos em subpastas nomeadas como: AUTOR - TITULO

Todo o conteúdo gerado é compactado no arquivo fichas_e_modelos.zip.
