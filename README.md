# Demonstrativo de Faturamento — DSV (Nouryon)

Aplicação web (estática, sem backend) que substitui o arquivo `Demonstrativo_-_Nouryon.xlsm`.
Roda inteiramente no navegador e pode ser hospedada gratuitamente no **GitHub Pages**.

## Arquivos deste pacote

| Arquivo | Descrição |
|---|---|
| `index.html` | Aplicação final (HTML + CSS + JS num único arquivo). É o que vai para o GitHub Pages. |
| `motor_calculo.py` | Mesma lógica de negócio em Python, usada para validar as regras antes de portar para JS. Rode com `python3 motor_calculo.py` — executa testes e imprime os resultados. |
| `README.md` | Este arquivo. |

## O que foi mapeado no arquivo original

Abri o `.xlsm` e extraí, célula a célula, todas as fórmulas e as macros VBA reais (usei
`openpyxl` para as fórmulas e `oletools`/`olevba` para o VBA — o VBA colado na sua mensagem
era de outro arquivo com lógica parecida; usei o VBA **real** deste arquivo como referência
de comportamento):

- **Aba "Capa Faturamento"**: cabeçalho, dados do cliente (`XLOOKUP` contra `Tabela Clientes`,
  chave = CNPJ na coluna F), dados do processo (`VLOOKUP`+`MATCH` contra `Relatório FollowNet`,
  chave = `Ref. Cliente` em `F22`), despesas não tributáveis (linhas 33–60) e serviços
  tributáveis (linhas 64–71) com **gross-up de 0,8575** (`Bruto = Líquido / 0,8575`).
- **Aba "CW1"**: tabela nomeada `Table1` com o layout real de exportação do sistema
  operacional (`Debtor`, `Job Num.`, `Posted - Revenue`, `Inv. Type`, `OS Sell Amt`,
  `Estimated Revenue`, `Description` etc.). As fórmulas `FILTER()` em `E48`/`N48`/`N75`
  geram automaticamente as despesas pagas pela Comissária e os adiantamentos.
- **Aba "Relatório FollowNet"**: dados do embarque/processo.
- **Aba "Tabela Clientes"**: cadastro de clientes (CNPJ na coluna F).
- **Aba "Tabela de Despesas"**: de-para entre descrição da despesa e sua categoria
  (`XLOOKUP` nas colunas D33:D60).
- **Aba "Custeio Mult"**: soma cruzada por categoria de despesa, usada como validação
  (célula `R7`/`R9`/`R10` na Capa comparam o total do demonstrativo com o Custeio Mult).
- **VBA real do arquivo**: `LimparEFormatarCNPJ_FollowNet` (formata CNPJ), `Mult` (exporta
  linha do Custeio Mult), `Exportar_PDF_Demonstrativo` e `SalvarCapaFaturamentoDownloads`
  (exportação de arquivos — na versão web isso virou os botões **Imprimir/Salvar PDF** e
  **Exportar dados (JSON)**).

## As 4 mudanças solicitadas pelo cliente — como foram implementadas

1. **Migração de dados só com Ref. Cliente preenchida** — a aplicação bloqueia a busca e
   exibe um alerta caso o campo esteja vazio (replica o comportamento de `IFERROR(...,"")`
   do Excel, mas de forma explícita para o usuário).
2. **Gerenciamento PO editável** — antes era um valor fixo digitado direto na célula.
   Agora é um campo de formulário (`Valor Bruto`, padrão R$ 327,60) que qualquer processo
   pode ajustar conforme a proposta comercial. O valor líquido é calculado automaticamente.
3. **Compensation Fee = 1,5% sobre despesas pagas pela DSV** — antes era digitado manualmente
   (`L66 = 1,44`, sem fórmula). Agora é calculado automaticamente como 1,5% da soma de
   "Pago pela Comissária" e pode ser recalculado a qualquer momento com o botão
   **Recalcular** (o campo continua editável para ajuste manual, se necessário).
4. **Frete internacional faturado pela DSV → valor líquido** — adicionei um aviso destacado
   na seção de despesas orientando que, ao lançar a categoria "Frete Intl. Pago", o valor
   informado deve ser o líquido (sem impostos).
5. **Geração correta dos valores do processo a partir da CW1** — reimplementei fielmente a
   lógica das fórmulas `FILTER()` originais: as despesas pagas pela Comissária e os
   adiantamentos são filtrados automaticamente por `Debtor = Vendor`, `Job Num. = Nr. Processo`,
   e as condições de `Posted - Revenue` / `Inv. Type` / `Estimated Revenue` exatamente como
   na planilha.

## Como usar

1. Abra `index.html` no navegador (ou publique no GitHub Pages — veja abaixo).
2. **Passo 1** — envie os arquivos (.xlsx ou .csv) exportados do sistema:
   Relatório FollowNet, Tabela Clientes, CW1 e Tabela de Despesas.
   (Pode clicar em **"Carregar dados de exemplo"** no topo para ver a aplicação funcionando
   com dados fictícios, sem precisar enviar arquivos.)
3. **Passo 2** — digite a `Ref. Cliente` e clique em **Buscar dados do processo**.
4. Confira os dados do cliente/processo, as despesas importadas da CW1 (destacadas em azul),
   ajuste o Gerenciamento PO se necessário, e adicione linhas manuais quando preciso.
5. Use **Imprimir/Salvar PDF** para gerar o documento final, ou **Exportar dados (JSON)**
   para guardar um registro dos valores usados.

## Publicar no GitHub Pages

```bash
# dentro da pasta do projeto
git init
git add index.html README.md
git commit -m "Demonstrativo de Faturamento - versão web"
git branch -M main
git remote add origin https://github.com/SEU_USUARIO/SEU_REPOSITORIO.git
git push -u origin main
```

Depois, no GitHub: **Settings → Pages → Branch: main / (root)** → Save.
Em alguns minutos o site estará disponível em
`https://SEU_USUARIO.github.io/SEU_REPOSITORIO/`.

## Observações e limitações conhecidas

- Como o GitHub Pages é hospedagem **estática** (sem servidor/banco de dados), a aplicação
  não se conecta diretamente ao CW1/FollowNet/ERP — os arquivos precisam ser exportados e
  enviados manualmente a cada uso. Os dados ficam apenas na memória do navegador (nada é
  enviado para nenhum servidor).
- A regra de PIS/COFINS/CSLL incidindo apenas sobre "Despachante + Gerenciamento PO" e o
  IRRF sobre a soma de todos os serviços tributáveis brutos reproduz exatamente as fórmulas
  `N79:N82` do arquivo original. Se essa regra fiscal mudar, ajuste as constantes
  `ALIQ_PIS`, `ALIQ_COFINS`, `ALIQ_CSLL`, `ALIQ_IRRF` no topo do `<script>` do `index.html`
  (e as equivalentes em `motor_calculo.py`).
- A validação contra o "Custeio Mult" hoje é apenas informativa (mostra o total apurado);
  se quiser que a aplicação também leia um arquivo de Custeio Mult e compare
  automaticamente, é só pedir — dá para adicionar um quinto upload.
