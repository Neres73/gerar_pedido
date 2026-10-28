# Coupa — Geração de Requisições a partir da Nota Fiscal

Script em Python + Playwright que lança requisições de compra no Coupa a partir da nota fiscal do fornecedor (XML, PDF ou até foto) ou de uma planilha. Ele escolhe o código de item (PLU) de cada linha, confere no Coupa se o código pode ser usado, preenche a requisição inteira e, no fim, confere o que foi lançado contra a nota.

Uso no meu trabalho, no back-office de Manutenção. Nas minhas medições, o lançamento manual leva 3 vezes o tempo do script.

---

## Por que existe

Lançar uma requisição a partir de uma nota é trabalhoso e cheio de detalhes que geram erro:

- Cada item da nota vira uma linha com PLU, fornecedor, quantidade, preço, data, empresa, centro de custo da loja e conta
- O PLU certo **não pode ter fornecedor nem valor atrelado** e precisa ter o centro de custo liberado, o que só se descobre abrindo o PLU no Coupa
- Frete, ST e IPI precisam ser distribuídos entre as linhas para a requisição **fechar exatamente no valor total da nota**
- O centro de custo depende da loja, que vem do CNPJ do destinatário da nota

---

## O que o script faz

**1. Lê a nota (ou a planilha)**
- **XML da NF-e**: leitura exata de itens, CNPJs e totais, igual para qualquer fornecedor. Explica por que um XML não serve (resumo sem itens, evento de cancelamento, página da internet salva como `.xml`)
- **PDF ou texto da DANFE**: extrai os itens com expressões regulares e confere a soma com o total dos produtos
- **PDF escaneado ou foto**: usa o **OCR do próprio Windows** (nada sai do computador), com o Tesseract como reserva
- **NFS-e** (nota de serviço): acha prestador, tomador e valor, validando os CNPJs pelo dígito verificador
- **Planilha**: aba `Linhas` com PLU, quantidade, valor, CNPJ e centro de custo

**2. Monta as linhas**
- Distribui frete, ST, IPI, seguro e descontos proporcionalmente entre os itens, fechando no valor total da nota
- Quando valor ÷ quantidade não dá um preço exato em centavos, lança 1 unidade com o valor total da linha, para o total no Coupa não sair com diferença
- Descobre a loja pelo CNPJ do destinatário e busca o centro de custo numa planilha de regiões (com cache, porque o arquivo é grande)
- Valida tudo antes de abrir o Coupa: PLU, quantidade, valor, CNPJ e centro de custo de cada linha

**3. Escolhe o PLU de cada item**, nesta ordem:
- **De-Para**: o código do fornecedor que já foi associado a um PLU antes
- **Confiáveis**: PLUs já aprovados, usados sem perguntar
- **Parecidos**: semelhança entre a descrição do item e a dos PLUs já usados
- **Mais baratos**: PLUs de valor baixo usados com frequência no último ano
- **Genérico**: um PLU coringa, oferecido por último

Cada candidato é **verificado no Coupa** (sem fornecedor, sem valor e com o centro de custo exigido). Os aprovados vão para a aba `Confiáveis`, os reprovados para `Bloqueados` e as associações novas para `De-Para`. **O script aprende com o uso** e pergunta cada vez menos.

**4. Lança no Coupa**, uma requisição atrás da outra
- Limpa linhas que tenham sobrado no carrinho
- Preenche o cabeçalho a partir de um perfil pronto (`perfis.json`) ou deixa você preencher
- Em cada linha: PLU, fornecedor (testando o CNPJ com e sem o zero da frente, como o Coupa às vezes guarda), quantidade, preço, data, empresa + centro de custo e conta, com paginação quando a requisição tem muitas linhas
- No fim, **lê as linhas lançadas e confere valor e centro de custo de cada uma contra a nota**

**5. Controle da execução**
- **F8** pausa e continua; **Ctrl+Alt+F8** para com segurança no próximo ponto seguro
- O progresso é salvo a cada linha: se algo cair, a próxima execução continua de onde parou
- Erro numa etapa: tenta de novo sozinho uma vez e, se continuar, pergunta o que fazer
- Cada requisição ganha uma pasta com resumo, conferência, log, a nota original e prints

---

## Instalação

Requisitos: **Windows** (o OCR usa o do sistema), Python 3.10+ e Microsoft Edge.

```bash
git clone https://github.com/Neres73/gerar_pedido.git
cd gerar_pedido

python -m venv .venv
.venv\Scripts\activate

pip install -r requirements.txt
playwright install msedge
```

## Configuração

```bash
copy .env.example .env
copy perfis.example.json perfis.json      # opcional: perfis de cabeçalho
```

| Variável | O que é |
|---|---|
| `COUPA_URL` | Página inicial do seu Coupa (obrigatória) |
| `COD_EMPRESA` | Código da empresa pesquisado antes do centro de custo (obrigatório) |
| `CAMINHO_PLANILHA` | Planilha da macro (abas `Linhas`, `PLUs`, `De-Para`, `Confiáveis`, `Bloqueados`) |
| `CAMINHO_REGIOES` | Planilha `.xlsb` com CNPJ e centro de custo das lojas (arquivo ou pasta) |
| `CNPJ_RAIZ_EMPRESA` | Raiz do CNPJ da própria empresa, para nunca ser confundida com o fornecedor |
| `CC_EXIGIDO` | Centro de custo que o PLU precisa ter liberado (opcional) |
| `PLU_GENERICO` | PLU coringa oferecido por último (opcional) |

As abas `De-Para`, `Confiáveis` e `Bloqueados` são criadas pelo próprio script. Ajustes de comportamento (quantas sugestões mostrar, verificar ou não no Coupa, teclas de pausa) ficam nas constantes do topo do `main.py`, todas comentadas.

## Uso

```bash
python main.py
```

O script pergunta a entrada (planilha ou nota) e o perfil do cabeçalho. Para a nota, cole o caminho do `.xml` ou `.pdf`, ou copie o texto do PDF e aperte Enter. Na primeira vez, faça o login no Coupa na janela do Edge que abrir; a sessão fica salva. O script **não lê, guarda nem digita senha**.

---

## Limitações

- Os seletores, campos personalizados e a estrutura das planilhas são os do ambiente onde o script foi feito; outro ambiente exige ajustes
- A leitura do PDF da DANFE depende do layout; o XML é sempre a entrada mais confiável
- A NFS-e é lida como uma linha só (não há itens padronizados nesse tipo de nota)
- Sem modo de simulação: o script **cria as requisições de verdade** (mas não as envia para aprovação; isso fica com você, depois de revisar)

## Estrutura

```
├── main.py               # leitura da nota, escolha de PLU, lançamento e conferência no Coupa
├── requirements.txt
├── .env.example          # modelo de configuração
├── perfis.example.json   # modelo de perfil de cabeçalho
└── .gitignore            # ignora .env, perfis.json, planilhas, notas e relatórios
```

---

Desenvolvido por [Gustavo Neres](https://github.com/Neres73), com apoio do Claude (IA) para modelar a lógica, escrever e revisar o código. As regras de escolha de PLU vêm do processo real e foram verificadas em execução. Projeto compartilhado como portfólio: nenhuma URL, credencial, código interno ou dado da empresa está no repositório. Licença MIT.
