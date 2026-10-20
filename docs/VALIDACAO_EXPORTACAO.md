> Documento histórico. O estado atual do pacote está em REVISAO_CONTEUDO_PUBLICO.md e LEIA-ME.md.

> Documento histórico. O estado atual do pacote está em REVISAO_CONTEUDO_PUBLICO.md e LEIA-ME.md.

# Validação da exportação na distribuição — 20/10/2026

A execução partiu de `app/Base FT - 01.10.26.xlsm`, após os ajustes de horários e a remoção das rotas AM feitos pelo utilizador. Este teste é diferente do conjunto oficial estático preservado em `examples/expected-output/`.

## Ajustes autorizados, somente na distribuição

- O segundo `SaveAs`, usado para ativar o compartilhamento legado da tabela de distribuição, foi substituído por `Save`. O compartilhamento legado foi dispensado pelo utilizador nesta cópia. O teste isolado reproduziu o erro 1004 na segunda gravação local após o Excel identificar o arquivo pela URL do OneDrive.
- Após criar a tabela, a rotina reutiliza `Workbooks(nomeArquivo)` em vez de tentar abrir novamente o arquivo já aberto pelo caminho local.
- Os eventos são suspensos apenas durante a abertura interna do gerador como engine de Ponto Focal e restaurados logo depois. Isso impede que a manutenção de abertura arquive a tabela histórica antes do Ponto Focal. A abertura normal do gerador mantém suas regras existentes.

As alterações estão no módulo de exportação da base atual e dos dois modelos consolidados em `app/` e `app/backup modelo_consolidado/`. As demais partes desses três XLSM, incluindo planilhas, fórmulas, valores, estilos e desenhos, foram preservadas a partir do snapshot anterior ao teste. O projeto de trabalho original em `Criação das Programações/v5.6` e o conjunto oficial em `examples/expected-output` não foram alterados.

## Resultado do reteste

O teste começou sem arquivos de saída ativos e com índice vazio. Não foi necessário suspender globalmente a manutenção no ambiente de teste: a própria chamada interna corrigida protege a etapa de exportação.

| Resultado | Local em app | Conferência |
| --- | --- | --- |
| Distribuição | `Tbl_Prog_Testes/Tbl_DistrMCD_FT_Sem40.xlsx` | 103 registros, 72 lojas, 31 rotas, 14.714 caixas |
| Infolog | `Templates/Template FT - 01.10.26.xlsm` | 103 registros; horários, origens, transportadoras e caixas coincidentes com a base |
| Ponto Focal | `Tbl_PontoFocal/2026/Sem40/Tbl_PontoFocal_Fat01_10_2026.xlsx` | 31 registros; turnos coincidentes com o template |
| Índice | `Tbl_PontoFocal_Index.xlsx` | Atualizado; 31 chaves únicas de rota |

O Excel apresentou a mensagem de exportação concluída. O template foi salvo após o preenchimento. Os arquivos foram comparados em disco, sem executar novas macros: nenhuma diferença nos campos verificados da distribuição e do template, nenhum erro de fórmula nas primeiras 500 linhas e 35 colunas das abas das saídas. As conferências usam o conteúdo da base validada pelo utilizador, não os totais do exemplo oficial anterior.

Os arquivos de saída e o índice processado foram deixados em `app/` para conferência. Assim, a pasta já não está no estado inicial vazio de distribuição. Antes da publicação ou de um novo teste de primeira execução, preparar novamente o estado inicial em etapa própria, preservando estes resultados.

Os resultados parciais de tentativas anteriores foram guardados fora do pacote público, na área local de trabalho. Nenhuma tradução foi realizada.
