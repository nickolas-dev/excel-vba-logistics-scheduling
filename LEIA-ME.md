# Cópia de distribuição — FT normal

Aplicação em português, conferida após a limpeza do conteúdo público. A apresentação principal está no [README em inglês britânico](README.md), com uma [versão em português](README.pt-BR.md). Código, interface e relatório permanecem em português. Nenhuma publicação foi realizada.

## Estrutura

- `app/`: gerador, modelos, cadastro fictício, índice inicial, imagens e PDF fictício. As pastas de saída contêm somente `.gitkeep`.
- `examples/expected-output/`: quatro exemplos oficiais com 78 lojas, 32 rotas e 14.770 caixas; 109 registros de lojas/CDs na distribuição e no template, e 32 registros de Ponto Focal.
- `src/vba/`: 81 arquivos de código para leitura, com índice de origem e hashes. Os .frm contêm eventos, sem descritor visual ou .frx de um export nativo. Os formulários executáveis permanecem nos XLSM.
- `docs/`: mapa, histórico de validação, conferência do conteúdo e manifesto atual.

## Executar

1. Use Excel desktop para Windows e Adobe Acrobat Reader. Preserve a estrutura de app e os nomes dos arquivos.
2. Abra o PDF fictício de app/arquivos_criacao_prog no Acrobat e copie todo o texto.
3. Abra app/criacao_prog_v5_6_testes.xlsm, cole como Texto em A2 da Planilha1 e clique em Criar.
4. Confira o consolidado. A validação operacional e a decisão de exportar pertencem ao utilizador.
5. Após validar, exporte para Tbl_Distribuição + Template Infolog.

Os destinos são relativos: app/Tbl_Prog_Testes, app/Templates e app/Tbl_PontoFocal. O índice inicial contém apenas os cabeçalhos e uma linha vazia formatada. Novas execuções atualizam o índice e produzem ou arquivam saídas conforme as regras existentes.

O compartilhamento legado da distribuição foi dispensado nesta cópia. A chamada interna do gerador como engine suspende somente os eventos dessa abertura, evitando manutenção durante a exportação. A abertura normal mantém as regras de manutenção.

## Exemplos e estado inicial

A base de referência em examples pode ser aberta para consulta: seu evento de abertura está inativo e não aplica expiração. Para executar o fluxo, use app; os botões do exemplo não foram adaptados para funcionar fora da aplicação.

O teste posterior com horários ajustados e rotas AM removidas tinha 72 lojas e 31 rotas. Esses arquivos foram preservados fora do pacote público, sem substituir os exemplos oficiais. O estado inicial foi restaurado depois dessa validação.

## Conferência

Consulte docs/REVISAO_CONTEUDO_PUBLICO.md. O endereço do vídeo interno foi removido do gerador e dos fontes, e os seletores não dependem de G:. Metadados internos e histórico de revisões foram limpos, preservando os dados. O PDF ficou visualmente idêntico nas 15 páginas. Os 12 arquivos Excel abriram em leitura no Excel; Criar e exportações não foram repetidos nesta etapa.

O escopo é FT normal. Os modelos OUTROS permanecem como suporte. A atualização manual de cache por Redistribuir_Ponto_Focal.xlsm e as fontes de Monta Palete não estão incluídas; o caso demonstrado usa o cache já presente.

Os dois READMEs estão preparados para revisão. Por decisão do autor, a aplicação continuará em português, incluindo código, comentários, interface e nomes técnicos internos. Não há tradução da aplicação prevista nesta etapa.
