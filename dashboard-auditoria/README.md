# Dashboard de Auditoria de Atendimento · IBEU

Painel web em Google Apps Script que lê direto as planilhas de respostas do Google Forms (Filiais e Central de Relacionamento). Com ele o gestor **baixa o PDF da auditoria de um AR com um clique**, e só quem foi liberado vê os dados.

- **Auditorias**: lista com os botões **Ver** (prévia) e **PDF**, além de **Baixar PDF das filtradas**.
- **Por AR**: média, menor nota e última auditoria de cada AR. O botão **PDF** junta todas as auditorias do AR.
- **Perguntas**: aproveitamento de cada item (do pior para o melhor) e evolução mensal.
- **Evolução por Filial** (por Equipe, na Central): tabela filial × mês e ranking com a linha da meta.

## Segurança: como o acesso funciona

1. **Login corporativo obrigatório.** O app é implantado como "Quem pode acessar: usuários do domínio". O Google pede o login da conta @ibeu.com.br, e ninguém de fora abre o link. O código também recusa qualquer e-mail de outro domínio (`CONFIG.dominio`).
2. **Lista de pessoas autorizadas.** Só entra quem está na aba **Acessos** da planilha de controle. Quem não está vê "Você não tem acesso a este painel" e não recebe nenhum dado.
3. **Restrição por painel e por filial ou equipe.** Cada linha da aba diz qual painel a pessoa vê e quais filiais ou equipes. O filtro roda **no servidor**: o navegador só recebe as linhas permitidas, e cada pedido de prévia ou de PDF é conferido de novo, então não adianta alterar a página para pedir outra auditoria.
4. **Log.** Aberturas, visualizações, downloads e tentativas negadas vão para a aba **Log**, com data, e-mail, painel, AR e linhas.
5. **As planilhas continuam fechadas.** O app roda com a conta de quem implantou, então os gestores **não** precisam (nem devem) ter acesso às planilhas do Forms ou ao código.

> Antes de publicar, confirme com o administrador do Google Workspace que o domínio permite web apps do Apps Script. Em algumas empresas isso é bloqueado no Admin Console.

## Planilha de controle

Crie uma planilha nova (ex.: "Controle · Painel de Auditoria") e compartilhe **só** com quem administra os acessos.

**Aba `Acessos`**: a primeira linha é o cabeçalho, e cada linha é uma pessoa em um painel.

| E-mail | Painel | Filiais/Equipes | Ativo |
|---|---|---|---|
| diretora@ibeu.com.br | Ambos | Todas | Sim |
| gestor.botafogo@ibeu.com.br | Filiais | Botafogo | Sim |
| regional.sul@ibeu.com.br | Filiais | Botafogo, Copacabana, Ipanema, Leblon | Sim |
| coord.central@ibeu.com.br | Central | Todas | Sim |
| supervisor.whats@ibeu.com.br | Central | Equipe WhatsApp | Sim |

- **Painel**: `Filiais`, `Central` ou `Ambos`.
- **Filiais/Equipes**: `Todas` ou os nomes **iguais aos da planilha do Forms**, separados por vírgula. Acentos e maiúsculas não importam.
- **Ativo**: `Não` bloqueia sem apagar a linha.
- Para dar acesso a um painel inteiro e a só algumas filiais do outro, use duas linhas.
- As mudanças levam **até 5 minutos** para valer (`CONFIG.cacheMinutos`).

**Aba `Log`**: é criada sozinha no primeiro acesso.

Proteja as duas abas em **Dados → Proteger páginas e intervalos**, deixando só você como editor.

## Como publicar

1. Abra a planilha de respostas das **Filiais** → **Extensões → Apps Script**.
2. Crie estes arquivos e cole o conteúdo de cada um:
   - `Code.gs`, `Relatorio.gs` e `Logo.gs` (arquivos de script);
   - `Index` e `Negado` (arquivos HTML, criados em **+ → HTML**, sem digitar o `.html`).
3. Em **Configurações do projeto**, marque "Mostrar arquivo de manifesto" e cole o `appsscript.json`.
4. No topo do `Code.gs`, ajuste o `CONFIG`:
   - `controle.planilhaId`: ID da planilha de controle (o trecho da URL entre `/d/` e `/edit`);
   - `origens.central.planilhaId`: ID da planilha da Central;
   - `aba` de cada origem: nome da aba de respostas;
   - `colunas`: só se algum cabeçalho tiver nome diferente;
   - `contatoAcesso`, `meta` e `dataInicio`.
5. **Implantar → Nova implantação → App da Web**:
   - Executar como: **Eu**
   - Quem pode acessar: **Qualquer pessoa em IBEU** (a opção do domínio). Não escolha "Qualquer pessoa".
6. Autorize, copie a URL `/exec` e envie só para as pessoas da aba Acessos. Para abrir direto na Central, use `.../exec?origem=central`.

Para atualizar sem mudar o link: **Implantar → Gerenciar implantações → Editar → Nova versão**.

## PDF (frente e verso)

- **Frente**: logo, nota, dados do atendimento e perguntas.
- **Verso**: Pontos de Auditoria, com as quebras de linha preservadas, os títulos em negrito e os marcadores. Se o texto for longo, continua na folha seguinte.
- Num PDF com várias auditorias, cada uma começa numa **frente**. Quando o verso ocupa um número par de páginas, entra uma página "em branco para impressão frente e verso". O número de páginas do verso é **estimado** pelo tamanho do texto (`REL_LINHAS_POR_PAGINA` em `Relatorio.gs`). Se notar desalinhamento com textos muito longos, ajuste esse número.
- O rodapé traz "Confidencial · uso interno IBEU", a data e o e-mail de quem gerou o PDF.
- O limite é de 60 auditorias por arquivo (`CONFIG.limitePdf`).

## Notas

- Toda coluna com `(Vale X)` no cabeçalho é uma pergunta que vale X pontos.
- A nota usa a coluna "Nota" da planilha, se existir. Senão, é calculada como `pontos ÷ pontos possíveis × 100`.
- `CONFIG.pesos` define quanto vale cada resposta. N/C conta como ponto cheio; use `null` para que não conte.
- Testes de permissão (simulam o Apps Script, sem precisar publicar): `node dashboard-auditoria/testes/teste-permissoes.js`.
- **Regra para quem mexer no código:** no Apps Script, toda função cujo nome não termina com `_` pode ser chamada pelo navegador. Só `doGet`, `listarOrigens`, `getDados`, `previaRelatorio` e `gerarPdf` são públicas, e todas conferem o usuário logado. O teste falha se aparecer outra função pública.
