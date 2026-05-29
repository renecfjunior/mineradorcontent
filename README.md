# ELEMENTAR — Minerador de Pautas

Bot com interface web local que coleta notícias do **G1**, **Nexo** e
**Agência IBGE**, pontua cada matéria pela **Matriz de Validação** do dossiê do
canal e sugere um **ângulo de vídeo** no estilo Elementar. Exporta o resultado
para **Word (.docx)**.

## Como rodar (no Antigravity ou em qualquer terminal)

1. Abra a pasta `elementar-bot` no Antigravity.
2. Instale as dependências:

   ```bash
   pip install -r requirements.txt
   ```

3. Rode o servidor:

   ```bash
   python app.py
   ```

4. Abra no navegador: **http://127.0.0.1:5000**

5. Escolha as fontes, clique em **Minerar pautas** e, se quiser, **Exportar para Word**.

> Dica no Antigravity: você pode pedir ao agente "rode `python app.py` e abra o
> preview" — ele sobe o servidor e mostra a página no navegador embutido.

## Como funciona a pontuação (0 a 12)

Cada notícia recebe 0–3 em quatro critérios do dossiê:

| Critério | O que mede |
|---|---|
| **Poder de Gancho** | Palavras de indignação/curiosidade (domina, monopólio, escândalo…) — pesa mais no título |
| **Densidade Estatística** | Números, %, R$, "bilhões", "milhões" |
| **Rastreabilidade da Prova** | Operação policial, CADE, TCU, processo, multa, auditoria, censo… |
| **Apelo de Massa** | Mexe no bolso/vida (preço, aluguel, salário, energia…) |

Soma **≥ 9** → marcada como **RECOMENDADA** (regra do dossiê).

## Onde ajustar

Tudo fica em `app.py`, fácil de editar:

- **`FONTES`** — adicionar/trocar feeds RSS. Cada fonte aceita vários
  candidatos de URL; usa o primeiro que responder.
- **`CATEGORIAS`** — palavras-chave das 4 categorias temáticas do dossiê.
- **`GANCHO`, `PROVA`, `MASSA`** — listas de palavras de cada critério.
- **`ANGULOS`** — os textos de sugestão de ângulo por categoria.

## Observações

- Roda 100% local; nada é enviado para fora além das requisições aos feeds.
- O **Nexo** é parcialmente pago e nem sempre publica RSS aberto — por isso o
  app tenta vários endereços e mostra o status de cada fonte na tela. Se o Nexo
  vier vazio, G1 e IBGE seguem funcionando normalmente.
- Os feeds do G1/IBGE não abrem de dentro de alguns sandboxes restritos, mas
  funcionam na sua máquina local.
