# Dashboard de Auditoria de Atendimento

Painel web, em Google Apps Script, que lê direto as planilhas de respostas do Google Forms (Filiais e Central de Relacionamento). Com ele o gestor **baixa o PDF da auditoria de um AR com um clique**.

- **Auditorias**: lista com os botões **Ver** (prévia do relatório) e **PDF**, além de **Baixar PDF das filtradas** (um arquivo, uma auditoria por página).
- **Por AR**: média, menor nota e última auditoria de cada AR. O botão **PDF** baixa todas as auditorias daquele AR.
- **Perguntas**: aproveitamento de cada item (do pior para o melhor) e evolução mensal da nota.
- **Evolução por Filial** (ou por Equipe, na Central): tabela filial × mês e ranking com a linha da meta.
- Filtros: Regional, Filial, AR, Auditor, Mês e período. Cada lista mostra só o que existe dentro dos outros filtros.
- O relatório PDF tem layout próprio para impressão. Os **Pontos de Auditoria** mantêm as quebras de linha, os títulos ("Pontos fortes:", "1. Diagnóstico…") ficam em negrito e os itens com "-" viram marcadores.

Não tem custo nem servidor: o painel roda na conta Google de quem publica.

## Arquivos

| Arquivo | O que é |
|---|---|
| `Code.gs` | Lê a planilha, calcula as notas e gera o PDF. **O `CONFIG` fica no topo.** |
| `Index.html` | O painel (abrindo direto no navegador, entra em modo demonstração com dados fictícios). |
| `appsscript.json` | Manifesto: fuso de São Paulo e web app restrito ao domínio. |

## Como publicar (≈10 min)

1. Abra a planilha de respostas das **Filiais** → **Extensões → Apps Script**.
2. Apague o conteúdo de `Código.gs` e cole o conteúdo de `Code.gs`.
3. Clique em **+ → HTML**, dê o nome `Index` (sem `.html`) e cole o conteúdo de `Index.html`.
4. (Opcional) Em **Configurações do projeto**, marque "Mostrar arquivo de manifesto appsscript.json" e cole o conteúdo de `appsscript.json`.
5. Ajuste o `CONFIG` no topo do `Code.gs`:
   - `origens.filiais.aba`: nome da aba das respostas (ex.: `Respostas ao formulário 1`).
   - `origens.central.planilhaId`: ID da planilha da Central (o trecho da URL entre `/d/` e `/edit`). Se não houver Central, apague o bloco `central`.
   - `colunas`: se algum cabeçalho tiver nome diferente (ex.: "Nome do Atendente" em vez de "Nome do AR"), acrescente-o à lista do campo. A busca ignora acentos e maiúsculas.
   - `pesos`: quanto vale cada resposta. Hoje "N/C" conta como ponto cheio (é o que bate com as médias do Power BI). Use `null` para que não conte nem a favor nem contra.
   - `dataInicio` e `meta`.
6. **Implantar → Nova implantação → App da Web**
   - Executar como: **Eu**
   - Quem pode acessar: **Qualquer pessoa em [sua empresa]**. Em conta Gmail pessoal, escolha "Qualquer pessoa com conta do Google".
7. Autorize o acesso e copie a URL `/exec`. Esse é o link que vai para os gestores.
   - Para abrir direto na Central: `.../exec?origem=central`

Os dados são lidos a cada abertura (ou ao clicar em **Atualizar**), então nada precisa ser sincronizado. Depois de editar o código, use **Implantar → Gerenciar implantações → Editar → Nova versão** para manter a mesma URL.

## Como as notas são calculadas

- Toda coluna cujo cabeçalho contém `(Vale X)` é uma pergunta que vale X pontos.
- Nota da auditoria: usa a coluna "Nota" da planilha, se existir. Senão, é calculada como `pontos obtidos ÷ pontos possíveis × 100`.
- A coluna "Nota média" da aba Perguntas é a média de pontos daquele item, e o aproveitamento é essa média ÷ Vale.

## Observações

- O PDF é gerado pelo próprio Google (`HTML → PDF`) e baixado direto, sem precisar de "Imprimir".
- O limite é de 60 auditorias por arquivo. Para baixar mais, filtre por mês ou regional.
- Quem tem o link vê todas as filiais. Se cada gestor só puder ver a própria filial, dá para restringir pelo e-mail de quem acessa (`Session.getActiveUser()`), mas é preciso montar uma aba com o mapeamento e-mail → filial.
