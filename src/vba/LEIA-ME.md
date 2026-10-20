# VBA para leitura no GitHub

Extração estática em UTF-8, sem tradução e sem execução de macros. Os XLSM continuam a ser os arquivos executáveis. Esta pasta foi criada para leitura, navegação e revisão das rotinas.

**Estado: fontes sincronizados após a limpeza local. O link interno foi removido e os seletores usam o caminho do projeto. A conferência está em docs/REVISAO_CONTEUDO_PUBLICO.md. Sem tradução ou publicação.**

| Pasta | Origem | Arquivos de código |
| --- | --- | ---: |
| `gerador/` | `app/criacao_prog_v5_6_testes.xlsm` | 34 |
| `consolidado/` | `app/modelo_consolidado.xlsm` | 8 |
| `consolidado_outros/` | `app/modelo_consolidado_OUTROS.xlsm` | 11 |
| `template_infolog/` | `app/modelo_template.xlsm` | 10 |
| `exemplo_consolidado/` | `examples/expected-output/Base FT - 01.10.26.xlsm` | 8 |
| `exemplo_template/` | `examples/expected-output/Template FT - 01.10.26.xlsm` | 10 |

O total é de 81 arquivos de código. `INDICE.json` relaciona cada módulo ao arquivo de origem e registra seu hash. As cópias em `backup modelo_consolidado` e a base atual de `app` têm código equivalente ao modelo correspondente, desconsiderando apenas os identificadores internos dos formulários. Isso não significa que os arquivos Excel ou os desenhos dos formulários sejam idênticos.

## Limites da extração

- `.bas`: código dos módulos padrão.
- `.cls`: código de classes e módulos de documento, como `EstaPastaDeTrabalho` e planilhas.
- `.frm`: somente o código dos eventos dos formulários extraído do projeto. Não contém o descritor visual de um export nativo, nem arquivos `.frx`.

**Esta pasta não é um pacote para reconstruir o projeto por importação no editor VBA.** Os formulários completos, referências e recursos binários continuam preservados nos XLSM. Uma eventual exportação nativa para reconstrução é outra etapa.

O texto foi mantido em português ou inglês conforme já estava no arquivo de origem. Apenas a codificação de saída e as quebras de linha foram normalizadas para UTF-8 e LF. Os atributos dos módulos foram mantidos. Nenhum comentário, regra, literal ou credencial foi silenciosamente removido do código extraído.

## Por onde começar

1. `gerador/Módulo2.bas`: entrada do botão Criar.
2. `gerador/Módulo1.bas` e `gerador/Módulo4.bas`: etapas de criação e montagem da base.
3. `consolidado/UserForm1.frm`: escolha do destino de exportação.
4. `consolidado/Módulo1.bas`: distribuição e chamada do engine de Ponto Focal.
5. `consolidado/Módulo2.bas`: Template Infolog e preenchimento dos turnos.
6. `gerador/modpontoFocal.bas` e `gerador/modIndicePontoFocal.bas`: geração e controle de Ponto Focal.
7. `template_infolog/modDadosExternos.bas`: consultas ao cadastro.

A conferência operacional e a decisão de exportar permanecem com o utilizador. Os módulos OUTROS são suporte de funcionalidades fora do caso FT normal validado. Os exemplos refletem uma execução anterior, e a base de exemplo tem `Workbook_Open` inativo para consulta permanente.
