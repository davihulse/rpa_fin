# rpa\_fin

Busca detalhamento de aquisições em consulta a instâncias de pagamento (FIN), disponibilizando de forma organizada para gestão de projetos.





┌─────────────────────────────────────────────────────────────────────────┐

│                        FLUXO GERAL DO RPA   (Atualizado em 17/02/2026                         │

└─────────────────────────────────────────────────────────────────────────┘



1\. INICIALIZAÇÃO

&nbsp;  ├─ Criar driver Chrome (headless)

&nbsp;  ├─ Conectar Google Sheets API

&nbsp;  │  ├─ Planilha Origem: "Acompanhamento\_Aquisições\_RPA"

&nbsp;  │  │  ├─ Aba "Dados"

&nbsp;  │  │  └─ Aba "EPROC" (concatenada)

&nbsp;  │  └─ Planilha Destino: "Acompanhamento\_FIN\_RPA"

&nbsp;  │     ├─ Aba "Dados" (worksheet\_fin)

&nbsp;  │     ├─ Aba "Manuais" (worksheet\_manuais)

&nbsp;  │     ├─ Aba "Ignorar" (worksheet\_ignorar)

&nbsp;  │     └─ Aba "Alertas" (worksheet\_alertas)

&nbsp;  └─ Preparar dados

&nbsp;     ├─ df\_dados\_rpa (Dados + EPROC)

&nbsp;     ├─ fins\_em\_dados (ID\_CARD já processados)

&nbsp;     ├─ fins\_em\_manuais (FINs para forçar extração)

&nbsp;     └─ fins\_ignorados (FINs para pular)



2\. DOWNLOAD PLANNERS (Microsoft Planner)

&nbsp;  ├─ driver.get(Planner URL)

&nbsp;  ├─ Aguardar carregamento completo

&nbsp;  ├─ login\_microsoft(driver)

&nbsp;  │  ├─ Verificar se botão "Opções de plano" existe (já logado)

&nbsp;  │  │  ├─ SIM → return (pula login)

&nbsp;  │  │  └─ NÃO → prossegue

&nbsp;  │  ├─ Inserir email

&nbsp;  │  └─ Inserir senha (3 tentativas)

&nbsp;  ├─ exportar\_planners(driver)

&nbsp;  │  ├─ Para cada um dos 3 Planners:

&nbsp;  │  │  ├─ Clicar "Opções de plano"

&nbsp;  │  │  ├─ Clicar "Exportar para Excel"

&nbsp;  │  │  └─ Aguardar download completo (.crdownload desaparecer)

&nbsp;  └─ consolidar\_planilhas()

&nbsp;     ├─ Ler 3 arquivos .xlsx

&nbsp;     ├─ pd.concat() → df

&nbsp;     └─ Filtrar buckets indesejados (Brementur, PC de Viagem)



3\. PROCESSAMENTO DO DF

&nbsp;  ├─ Extrair "Numero Tarefa" do "Nome da tarefa"

&nbsp;  │  ├─ Regex: números 5-7 dígitos após "Tarefa/Chamado"

&nbsp;  │  └─ Regex: formato "CT 082/25" ou "CT082/25"

&nbsp;  ├─ Extrair "FIN" dos "Itens da lista de verificação"

&nbsp;  │  └─ Regex: FIN.XXXXXX/YY

&nbsp;  └─ Manter campo "Identificação da tarefa" (ID único do card)



4\. LOGIN SE SUITE

&nbsp;  └─ login\_sesuite()

&nbsp;     ├─ driver.get(SE Suite URL)

&nbsp;     ├─ Inserir email

&nbsp;     └─ Inserir senha



5\. EXTRAÇÃO - FASE 1: FINs MANUAIS

&nbsp;  ├─ Remover FINs ignorados da aba Manuais

&nbsp;  └─ Para cada FIN em fins\_em\_manuais:

&nbsp;     ├─ Buscar FIN no df consolidado → linha\_com\_fin

&nbsp;     ├─ Extrair: titulo\_card, numero\_tarefa, id\_card\_atual

&nbsp;     │

&nbsp;     ├─ dados\_fin = extrai\_fin(fin)

&nbsp;     │  └─ SE falha → registrar\_alerta(FALHA\_EXTRACAO) + continue

&nbsp;     │

&nbsp;     ├─ SE linha\_com\_fin vazia → print aviso + continue

&nbsp;     │

&nbsp;     ├─ Buscar aquisição em df\_dados\_rpa (por numero\_tarefa)

&nbsp;     │  └─ SE não encontrar → registrar\_alerta(AQUISICAO\_NAO\_ENCONTRADA) + continue

&nbsp;     │

&nbsp;     ├─ Preparar dados\_aquisicao

&nbsp;     │  ├─ ID\_CARD = id\_card\_atual

&nbsp;     │  └─ Título do Card = titulo\_card

&nbsp;     │

&nbsp;     ├─ registrar\_fin\_google\_sheets(dados\_fin, dados\_aquisicao, worksheet\_fin)

&nbsp;     │  ├─ Verificar se ID\_CARD já existe

&nbsp;     │  │  ├─ SIM → Atualizar linha existente

&nbsp;     │  │  └─ NÃO → Inserir nova linha

&nbsp;     │  └─ Calcular e atualizar SALDO

&nbsp;     │     ├─ SE saldo < -9.99 → registrar\_alerta(SALDO\_NEGATIVO)

&nbsp;     │     └─ SE saldo >= 0 → remover\_alerta(SALDO\_NEGATIVO)

&nbsp;     │

&nbsp;     ├─ Remover alertas resolvidos:

&nbsp;     │  ├─ remover\_alerta(id\_card\_atual, FALHA\_EXTRACAO)

&nbsp;     │  ├─ remover\_alerta(id\_card\_atual, DOC\_DIVERGENTE)

&nbsp;     │  ├─ remover\_alerta(id\_card\_atual, SEM\_NF\_CARD)

&nbsp;     │  └─ remover\_alerta(id\_card\_atual, AQUISICAO\_NAO\_ENCONTRADA)

&nbsp;     │

&nbsp;     └─ Remover FIN da aba Manuais



6\. EXTRAÇÃO - FASE 2: FINs DO PLANNER

&nbsp;  ├─ Para cada FIN único em df\[FIN]:

&nbsp;     │

&nbsp;     ├─ SE FIN em fins\_ignorados → continue

&nbsp;     │

&nbsp;     ├─ Buscar linha\_com\_fin no df

&nbsp;     │  └─ SE vazia → continue

&nbsp;     │

&nbsp;     ├─ Extrair id\_card\_atual

&nbsp;     │  └─ SE id\_card\_atual em fins\_em\_dados → continue (já processado)

&nbsp;     │

&nbsp;     ├─ Extrair: titulo\_card, numero\_doc\_card, numero\_tarefa

&nbsp;     │

&nbsp;     ├─ dados\_fin = extrai\_fin(fin)

&nbsp;     │  └─ SE falha → registrar\_alerta(FALHA\_EXTRACAO) + continue

&nbsp;     │

&nbsp;     ├─ VALIDAÇÃO: Número do Documento

&nbsp;     │  ├─ SE numero\_doc\_card existe:

&nbsp;     │  │  ├─ Comparar com doc\_fiscal\_sesuite (apenas dígitos)

&nbsp;     │  │  └─ SE divergente → registrar\_alerta(DOC\_DIVERGENTE) + continue

&nbsp;     │  └─ SE numero\_doc\_card NÃO existe:

&nbsp;     │     └─ registrar\_alerta(SEM\_NF\_CARD) (mas continua processando)

&nbsp;     │

&nbsp;     ├─ Buscar aquisição em df\_dados\_rpa (por numero\_tarefa)

&nbsp;     │  └─ SE não encontrar → registrar\_alerta(AQUISICAO\_NAO\_ENCONTRADA) + continue

&nbsp;     │

&nbsp;     ├─ Preparar dados\_aquisicao

&nbsp;     │  ├─ ID\_CARD = id\_card\_atual

&nbsp;     │  └─ Título do Card = titulo\_card

&nbsp;     │

&nbsp;     ├─ registrar\_fin\_google\_sheets(dados\_fin, dados\_aquisicao, worksheet\_fin)

&nbsp;     │

&nbsp;     ├─ Remover alertas resolvidos (4 tipos)

&nbsp;     │

&nbsp;     └─ SE FIN estava em Manuais → remover de lá



7\. FINALIZAÇÃO

&nbsp;  └─ driver.quit()



┌─────────────────────────────────────────────────────────────────────────┐

│                    FUNÇÃO: extrai\_fin(numfin)                                                 │

└─────────────────────────────────────────────────────────────────────────┘



├─ driver.get(SE Suite home) + sleep(0.5)

├─ Aceitar alerta se houver

├─ Buscar campo de pesquisa FIN (3 XPaths possíveis)

│  └─ SE não encontrar → return None

├─ Limpar campo, digitar FIN, pressionar ENTER

├─ Aguardar resultado aparecer (timeout 20s)

│  └─ SE timeout → return None

├─ Extrair texto do link (para validar DOC FISCAL depois)

├─ Clicar no primeiro resultado

├─ Trocar para nova janela

├─ Extrair título completo e status

├─ Trocar para frame "ribbonFrame"

├─ Trocar para frame "frame\_form\_\*"

├─ Extrair 21 campos do formulário

├─ Mapear códigos → valores legíveis (4 mapas)

├─ Fechar janelas auxiliares

├─ Retornar para janela principal

└─ return dados\_dos\_chamados (+ \_doc\_fiscal\_validacao)



┌─────────────────────────────────────────────────────────────────────────┐

│          FUNÇÃO: registrar\_fin\_google\_sheets(dados\_fin, ...)                                  │

└─────────────────────────────────────────────────────────────────────────┘



├─ Montar linha com 33 colunas (ID\_CARD, Título do Card, dados aquisição + FIN)

├─ Ler planilha existente → df\_existente

├─ Verificar se ID\_CARD já existe

│  ├─ SIM → Atualizar linha existente

│  └─ NÃO → Inserir nova linha (append)

│

├─ Recarregar df\_existente

│

└─ CALCULAR SALDO (por Identificador/número do chamado):

&nbsp;  ├─ Filtrar registros do mesmo chamado (exceto linha "Saldo")

&nbsp;  ├─ Somar "Valor Líquido a Pagar (R$)"

&nbsp;  ├─ saldo = Valor Aquisição - soma\_fins

&nbsp;  │

&nbsp;  ├─ SE saldo < -9.99:

&nbsp;  │  └─ registrar\_alerta(SALDO\_NEGATIVO)

&nbsp;  ├─ SE saldo >= 0:

&nbsp;  │  └─ remover\_alerta(SALDO\_NEGATIVO)

&nbsp;  │

&nbsp;  ├─ SE saldo != 0:

&nbsp;  │  ├─ Criar linha\_saldo (ID\_CARD vazio, FIN="Saldo", valor em "Valor Bruto")

&nbsp;  │  └─ Atualizar ou inserir linha Saldo

&nbsp;  └─ SE saldo == 0 E existe linha Saldo:

&nbsp;     └─ Deletar linha Saldo



┌─────────────────────────────────────────────────────────────────────────┐

│                    SISTEMA DE ALERTAS                                                         │

└─────────────────────────────────────────────────────────────────────────┘



CHAVE PRIMÁRIA: ID\_CARD + Tipo



TIPOS DE ALERTA:

├─ FALHA\_EXTRACAO

│  ├─ Quando: extrai\_fin() retorna None

│  └─ Remove: Quando extração funciona

│

├─ DOC\_DIVERGENTE

│  ├─ Quando: numero\_doc\_card != doc\_fiscal\_sesuite

│  └─ Remove: Quando documentos batem

│

├─ SEM\_NF\_CARD

│  ├─ Quando: Título do card não tem "NF nº:"

│  └─ Remove: Após processamento (não bloqueia extração)

│

├─ AQUISICAO\_NAO\_ENCONTRADA

│  ├─ Quando: numero\_tarefa não encontrado em df\_dados\_rpa

│  └─ Remove: Quando aquisição é encontrada

│

└─ SALDO\_NEGATIVO

&nbsp;  ├─ Quando: saldo < -9.99 para um Identificador

&nbsp;  └─ Remove: Quando saldo >= 0



ESTRUTURA DA ABA ALERTAS:

ID\_CARD | Número do FIN | Título do Card | Identificador | Tipo | Mensagem | Data



┌─────────────────────────────────────────────────────────────────────────┐

│                    CHAVES PRIMÁRIAS                                                           │

└─────────────────────────────────────────────────────────────────────────┘



ABA DADOS:

├─ Chave: ID\_CARD (Identificação da tarefa do Planner)

├─ Lógica: Se ID\_CARD já existe → atualiza linha

└─ Permite: Mudança de FIN sem criar linha duplicada



ABA ALERTAS:

├─ Chave: ID\_CARD + Tipo

├─ Lógica: Se ID\_CARD + Tipo já existe → atualiza linha

└─ Permite: Múltiplos alertas de tipos diferentes para mesmo card



ABA MANUAIS:

└─ Apenas lista de FINs (coluna A, sem cabeçalho)



ABA IGNORAR:

└─ Apenas lista de FINs (coluna A, sem cabeçalho)



┌─────────────────────────────────────────────────────────────────────────┐

│                    CAMPOS EXTRAÍDOS (33 colunas)                                              │

└─────────────────────────────────────────────────────────────────────────┘



CARD/AQUISIÇÃO:

├─ ID\_CARD (chave primária)

├─ Título do Card

├─ Código Unidade

├─ Identificador (número do chamado/tarefa)

├─ Apelido Projeto

├─ Descrição

├─ Fonte

├─ Rubrica

├─ Valor Aquisição R$

└─ Ordem de Compra (Aquisição)



FIN (do SE Suite):

├─ Número do FIN

├─ Descrição FIN

├─ Status FIN

├─ Data da Abertura do FIN

├─ Tipo de Documento

├─ Especificação

├─ Valor pago por Adiantamento?

├─ Filial Faturada

├─ CNPJ Fornecedor

├─ Número do documento

├─ Tipo de Compra

├─ Ordem de compra (FIN)

├─ Contrato (FIN)

├─ Registro Gerado (Apontamento)

├─ RNs

├─ Observações

├─ Número AP

├─ Data Agendada para Pagamento

├─ Competência

├─ Valor Bruto a Pagar (R$)

├─ Valor a deduzir (R$)

├─ Valor Líquido a Pagar (R$)

└─ Nr. do documento (CAP)



┌─────────────────────────────────────────────────────────────────────────┐

│                    VALIDAÇÕES IMPLEMENTADAS                                                   │

└─────────────────────────────────────────────────────────────────────────┘



1\. Documento Divergente (BLOQUEIA extração):

&nbsp;  ├─ Extrai "NF nº: XXXXX" do título do card

&nbsp;  ├─ Extrai "DOC FISCAL: XXXXX" do link no SE Suite

&nbsp;  ├─ Compara apenas dígitos (remove pontuação)

&nbsp;  └─ Se divergir → alerta + continue (não extrai)



2\. Título sem NF (NÃO bloqueia):

&nbsp;  ├─ Se título não tem "NF nº:"

&nbsp;  └─ Alerta + processa normalmente



3\. Saldo Negativo:

&nbsp;  ├─ Agrupa por Identificador (número do chamado)

&nbsp;  ├─ Soma todos os FINs do mesmo chamado

&nbsp;  ├─ Compara com Valor Aquisição

&nbsp;  └─ Se saldo < -9.99 → alerta



4\. ID\_CARD duplicado:

&nbsp;  ├─ Sempre atualiza em vez de duplicar

&nbsp;  └─ Permite mudança de FIN no card



5\. FINs Ignorados:

&nbsp;  └─ Remove da aba Manuais antes de processar

