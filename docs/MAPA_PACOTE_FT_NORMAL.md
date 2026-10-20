> Documento histórico. O estado atual do pacote está em REVISAO_CONTEUDO_PUBLICO.md e LEIA-ME.md.

> Documento histórico. O estado atual do pacote está em REVISAO_CONTEUDO_PUBLICO.md e LEIA-ME.md.

# Mapa do pacote — FT normal

Proposta para o primeiro caso de demonstração do portfólio, conferida em 20/10/2026. Este documento define o pacote; não reorganiza, traduz ou modifica os arquivos Excel, PDFs ou macros. Nenhuma macro foi executada nesta etapa.

## Escopo

Uma versão funcional em inglês britânico, com `README.md` em inglês e `README.pt-BR.md` em português. A tradução será uma etapa posterior. O PDF fictício permanece inicialmente no formato português que o interpretador espera.

Fluxo demonstrado: PDF fictício FT normal → cópia pelo Adobe Acrobat Reader → colagem como Texto em A2 → Criar → conferência humana do consolidado → exportação para distribuição e Infolog, com geração de Ponto Focal.

A decisão operacional de validar e exportar pertence ao utilizador. As verificações técnicas do exemplo não substituem essa conferência.

## Arquivos que entram

| Arquivo ou grupo atual | Papel no pacote | Decisão |
| --- | --- | --- |
| `criacao_prog_v5_6_testes.xlsm` | Gerador e engine de Ponto Focal | Incluir uma cópia limpa, preservando o código validado pelo utilizador |
| `modelo_consolidado.xlsm` | Modelo principal de consolidado | Incluir |
| `backup modelo_consolidado/modelo_consolidado.xlsm` | Modelo que o FT normal abre efetivamente | Incluir; o nome da pasta não significa que seja dispensável |
| `modelo_template.xlsm` | Modelo usado pela exportação Infolog | Incluir |
| `Tbl_Dados_Lojas_Transportador/Tbl_Dados_Lojas_Transportador.xlsx` | Cadastro com lojas fictícias e dados de consulta | Incluir a versão fictícia; conferir também tabelas e abas auxiliares antes da publicação |
| `Tbl_PontoFocal_Index.xlsx` | Controle persistente de chaves já processadas | Incluir um estado inicial preparado para demonstração, mantendo estrutura, fórmulas e configurações necessárias; não distribuir inadvertidamente o índice já processado do teste |
| `assets/` | Imagens usadas pela interface | Incluir os seis arquivos existentes; `fundo_userform.jpg` tem referência externa confirmada no código |
| `arquivos_criacao_prog/Preliminar FT normal - Portfolio ficticio.pdf` | Entrada fictícia validada no Acrobat | Incluir como entrada do exemplo |
| Código VBA em `.bas`, `.cls` e `.frm` com respectivos `.frx` | Permitir leitura e revisão no GitHub sem abrir o Excel | Exportar em etapa posterior, preservando atributos, codificação e arquivos binários dos formulários; não substituir os XLSM executáveis |
| READMEs, guia do fluxo, mapa de arquivos e imagens da demonstração | Apresentar o problema, a solução e a execução | Produzir posteriormente; não publicar o diário técnico como se fosse o guia final |

## Resultados de exemplo

Guardar um conjunto coerente em `examples/expected-output/`, separado das pastas de trabalho:

- Um consolidado FT normal escolhido e conferido pelo utilizador, entre as bases existentes. Não escolher automaticamente entre o arquivo original e a cópia `(1)`.
- A distribuição FT normal correspondente. No inventário atual ela está em `Tbl_Prog_Testes/Backup/Tbl_Prog - Backup/2026/Tbl_DistrMCD_FT_Sem40.xlsx`; conferir qual versão corresponde ao consolidado escolhido antes de copiar.
- O Template Infolog correspondente.
- A tabela de Ponto Focal correspondente, com 32 rotas no teste documentado.
- Capturas da entrada, consolidado e exportações, acompanhadas da sequência de uso.

Referência do caso já testado: 78 lojas Z00–Z77, 32 rotas, 14.770 caixas e 109 registros de lojas/CDs na distribuição e no template. Alterações feitas pelo utilizador após a geração devem fazer parte do conjunto escolhido, sem misturar resultados de execuções diferentes.

As bases têm lógica de expiração. Exemplos XLSM não devem ser apresentados como arquivos utilizáveis indefinidamente: documentar esse comportamento e oferecer capturas como evidência permanente. A execução reproduzível deve partir do PDF e dos modelos. Não modificar a regra de expiração nesta organização.

## Pastas de trabalho e estado inicial

Manter na distribuição a estrutura que o VBA espera:

| Pasta | Estado inicial proposto |
| --- | --- |
| `Templates/` | Vazia para novas exportações |
| `Tbl_Prog_Testes/` | Vazia, com subpastas necessárias preservadas |
| `Tbl_PontoFocal/` | Vazia; a rotina cria ano e semana |
| `Tbl_Prog/` e `Backup/Tbl_Prog - Backup/` | Estrutura preservada para as referências existentes; o exemplo usa o modo de teste |

Preservar pastas vazias com `.gitkeep` ou documentação. Ignorar saídas geradas, arquivos temporários do Office (`~$*`) e backups de execução no Git. As exceções serão os resultados escolhidos dentro de `examples/expected-output/`.

Preparar o índice inicial e as pastas vazias em uma cópia de distribuição. Não limpar o ambiente validado do utilizador. Não criar agora um botão ou macro de reinicialização.

## Arquivos que ficam fora da primeira publicação

- PDFs reais, incluindo o FT normal original, pallets dedicados e NE.
- ZIPs e versões anteriores da raiz de Portifolio.
- `.portfolio-work/`, snapshots, scripts locais de automação e diagnósticos da adaptação.
- Bases duplicadas e todos os resultados de testes que não forem escolhidos como exemplos.
- `backup criacao_prog/` e `backup modelo_template/`: não são dependências do fluxo FT normal nas referências externas examinadas. Manter localmente; não incluir apenas por serem backups.
- Dados externos de Monta Palete e o modelo legado ausente `Template Transportes Infolog 2.0.xlsx`: pertencem a fluxos fora deste exemplo.
- O estado processado do índice de Ponto Focal como estado inicial da demonstração.

## Componentes de outros fluxos: preservar compatibilidade

`modelo_consolidado_OUTROS.xlsm` e sua cópia em `backup modelo_consolidado/` são referenciados por outros recursos do gerador. Proposta: mantê-los inicialmente como arquivos de suporte, identificados como fora do caso validado. Não remover módulos, botões nem arquivos referenciados apenas para reduzir o pacote. A presença desses componentes não significa que os fluxos foram validados para publicação.

O arquivo `Redistribuir_Ponto_Focal.xlsm`, atualmente dois níveis acima de `v5.6`, é uma dependência da atualização manual do cache de Pontos Focais. O FT normal usa o cache existente do gerador; essa atualização externa não é necessária para repetir o caso testado. Proposta para a primeira publicação: documentar o cache incluído e a atualização externa como fora do escopo. Se decidirmos oferecer essa função no pacote, será necessário incluir uma versão adequada da dependência e adaptar seu caminho em uma etapa autorizada. Não anunciar o pacote como completo para todos os botões existentes.

## Estrutura proposta do repositório

```text
/
├── README.md
├── README.pt-BR.md
├── .gitignore
├── docs/                          # fluxo, arquitetura, dados e capturas
├── src/vba/                       # código exportado para leitura
├── examples/
│   └── expected-output/            # conjunto estático escolhido
└── app/
    ├── criacao_prog_v5_6_testes.xlsm
    ├── modelo_consolidado.xlsm
    ├── modelo_consolidado_OUTROS.xlsm
    ├── modelo_template.xlsm
    ├── Tbl_PontoFocal_Index.xlsx   # estado inicial
    ├── backup modelo_consolidado/
    │   ├── modelo_consolidado.xlsm
    │   └── modelo_consolidado_OUTROS.xlsm
    ├── Tbl_Dados_Lojas_Transportador/
    │   └── Tbl_Dados_Lojas_Transportador.xlsx
    ├── arquivos_criacao_prog/
    │   └── Preliminar FT normal - Portfolio ficticio.pdf
    ├── assets/
    ├── Templates/
    ├── Tbl_Prog_Testes/
    ├── Tbl_PontoFocal/
    ├── Tbl_Prog/
    └── Backup/
```

Os nomes técnicos internos estão mantidos nesta proposta para não quebrar referências. A futura tradução deve distinguir textos da interface de nomes de arquivos, abas, colunas, fórmulas e palavras reconhecidas no PDF. Os nomes definitivos do projeto e do repositório ainda serão escolhidos.

`PortfolioPath` identifica a raiz pelos nomes do gerador e de `modelo_template.xlsm`, permitindo colocar esse conjunto em `app/` sem fixar o nome do utilizador. A mudança real de pasta ainda deverá ser testada na cópia de distribuição. O GitHub apresenta e distribui o projeto; a execução continua no Excel desktop para Windows e a entrada validada continua a depender da cópia pelo Acrobat.

## Próximas etapas

1. Escolher o conjunto de resultados de exemplo e preparar uma cópia de distribuição com o estado inicial correto.
2. Conferir o conteúdo público: dados fictícios, células/abas ocultas, metadados, links externos e imagens efetivamente incluídos. Preservar as marcas e demais dados que o utilizador já autorizou manter; confirmar separadamente que o fonte a publicar não contém credenciais ou caminhos pessoais desnecessários.
3. Exportar o VBA para revisão e documentar o contrato de entrada do PDF, a arquitetura e as limitações de escopo.
4. Adaptar os textos para inglês britânico e escrever os dois READMEs, em etapa posterior.
5. Validar o pacote em uma nova pasta, com saídas vazias e índice inicial, antes de publicar.

Os erros 52 e 1004 são considerados resolvidos conforme correção e reteste informados pelo utilizador. O histórico em `CAMINHOS.md` descreve tentativas anteriores e deverá ser consolidado no guia final; não é uma pendência reaberta por este mapa.
