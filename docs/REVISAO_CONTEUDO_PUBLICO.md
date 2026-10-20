# Conferência do conteúdo público — 20/10/2026

A limpeza autorizada foi concluída somente na distribuição. Nenhuma tradução ou publicação foi realizada. O projeto original foi preservado. Os resultados do teste com 72 lojas e 31 rotas estão guardados fora do pacote público; os quatro exemplos oficiais continuam com 78 lojas e 32 rotas.

## Tratamentos concluídos

- O endereço corporativo ainda estava presente no gerador gravado da distribuição na conferência inicial. Foi retirado. O botão de tutorial agora informa que o vídeo não está incluído. O fonte exportado também ficou sem o endereço.
- As duas chamadas fixas a `ChDrive "G"` foram substituídas pela unidade resolvida por `PortfolioPath`. As expressões de unidade e pasta dos seletores foram verificadas no Excel.
- Foram removidos os elementos `absPath` com caminhos anteriores de gravação dos arquivos Excel. O nome do autor e as marcas autorizadas foram mantidos.
- O PDF guarda apenas título e autor fictícios nos metadados. As páginas não foram redesenhadas.
- O exemplo de distribuição ficou sem logs de revisões e utilizadores, relações de pacote, tipos de conteúdo associados e ponteiro de revisão.
- As saídas do teste e a base de trabalho foram preservadas fora do pacote. As pastas de saída contêm apenas marcadores `.gitkeep`, e o índice voltou a zero registros.
- Os 81 arquivos de código foram regenerados em UTF-8 e tiveram os hashes conferidos.

## Evidências

- Os 12 arquivos Excel do pacote abriram em leitura no Excel desktop, com carregamento normal solicitado e eventos desativados. Não foram executados Criar, exportações ou manutenção de abertura.
- A função `PortfolioPath` e as expressões de unidade e pasta usadas pelos seletores foram executadas para conferir o destino local. Não foi importado nenhum arquivo nesses seletores.
- No gerador, somente o código de `UserForm1` e `frmInstrucoes` foi alterado. As demais partes do XLSM foram preservadas.
- Os arquivos ZIP e XML foram relidos. Células, estilos, textos e valores de cálculo dos exemplos oficiais permaneceram iguais aos anteriores à limpeza. O VBA desses exemplos não mudou nesta etapa.
- O texto extraído e os fluxos de conteúdo das 15 páginas do PDF permaneceram iguais. As renderizações com Poppler ficaram visualmente idênticas nas 15 páginas.
- O PDF mantém Z00–Z77 e DEMOUSER. A aba Index mantém Chave e DataFaturamento e não contém registros processados.
- As buscas finais não encontraram o link interno, as chamadas fixas à unidade G nem os caminhos pessoais de última gravação pesquisados nos pacotes XML.
- O histórico removido não permanece como partes órfãs: relações e tipos de conteúdo foram tratados junto com os logs.

## Escopo e limites

Empresas, marcas, CNPJs, nome do autor, CDs, comentários e abas auxiliares foram mantidos conforme autorizado. A revisão anterior pesquisou os códigos reais conhecidos do FT normal em células, fórmulas, tabelas internas de textos e revisões; não houve candidatos. Isso não é uma certificação de ausência de todo identificador operacional possível.

O pacote cobre FT normal. Modelos OUTROS são suporte, sem ampliar o escopo validado. A atualização manual do cache de Ponto Focal usa uma dependência externa não incluída; o exemplo usa o cache existente. A comparação fixa de utilizador na abertura dos modelos permanece como regra de somente leitura, não como credencial de login.

Os exemplos são referências estáticas. A base de referência não expira, mas seus botões não foram adaptados para executar fora de app. Os fontes .frm contêm código dos eventos, sem descritor visual ou .frx para reconstrução por importação; formulários completos permanecem nos XLSM.

Não foi repetida a colagem pelo Acrobat e a execução completa após a retirada dos metadados: conteúdo, texto e renderizações do PDF foram comparados diretamente. Após esta conferência, o autor decidiu preservar código e interface em português. Foram preparados README.md em inglês britânico e README.pt-BR.md em português, sem alterar a aplicação. A publicação continua pendente.
